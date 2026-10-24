import { drizzleAdapter } from "@better-auth/drizzle-adapter";
import { drizzle } from "drizzle-orm/d1";
import { betterAuth } from "better-auth";
import * as authSchema from "./auth-schema";

export interface Env {
  DB: D1Database;
  ASSETS: Fetcher;
  BETTER_AUTH_SECRET?: string;
  BETTER_AUTH_URL?: string;
}

type SessionUser = { id: string; name: string; email: string };
type TxRow = {
  id: string; account_id: string; category_id: string | null;
  transaction_type: "income" | "expense" | "savings_contribution" | "transfer";
  amount_centavos: number; transaction_date: string; description: string;
  payment_method: "cash" | "bank_transfer" | "e_wallet" | "credit_card" | "other";
  notes: string | null;
  account_name?: string; category_name?: string | null;
};

const json = (body: unknown, status = 200) => Response.json(body, {
  status,
  headers: { "Cache-Control": "no-store", "X-Content-Type-Options": "nosniff" },
});
const fail = (message: string, status = 400) => json({ error: message }, status);
const id = () => crypto.randomUUID();

function manilaDateParts(date = new Date()) {
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone: "Asia/Manila", year: "numeric", month: "2-digit", day: "2-digit",
  }).formatToParts(date);
  const get = (type: string) => parts.find((part) => part.type === type)?.value ?? "";
  return { year: get("year"), month: get("month"), day: get("day") };
}
function todayInManila() {
  const p = manilaDateParts();
  return `${p.year}-${p.month}-${p.day}`;
}
function monthStartInManila() {
  const p = manilaDateParts();
  return `${p.year}-${p.month}-01`;
}
function validDate(value: unknown): value is string {
  if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
  const date = new Date(`${value}T00:00:00Z`);
  return Number.isFinite(date.getTime()) && date.toISOString().slice(0, 10) === value;
}
function validPaymentMethod(value: unknown): value is "cash" | "bank_transfer" | "e_wallet" | "credit_card" | "other" {
  return value === "cash" || value === "bank_transfer" || value === "e_wallet" || value === "credit_card" || value === "other";
}
function amountCentavos(value: unknown): number | null {
  if (typeof value !== "number" || !Number.isFinite(value) || value < 0.01 || value > 1000000000) return null;
  const cents = Math.round(value * 100);
  return Number.isSafeInteger(cents) && cents > 0 ? cents : null;
}
async function readJson(request: Request): Promise<Record<string, unknown> | null> {
  try {
    const value: unknown = await request.json();
    return value !== null && typeof value === "object" && !Array.isArray(value)
      ? value as Record<string, unknown> : null;
  } catch { return null; }
}

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    const url = new URL(request.url);
    if (url.pathname === "/api/health") return json({ ok: true, service: "budgetwise-api" });

    if (url.pathname.startsWith("/api/auth/")) {
      if (!env.BETTER_AUTH_SECRET) return fail("Authentication is not configured. Set BETTER_AUTH_SECRET in Worker secrets.", 503);
      const baseURL = env.BETTER_AUTH_URL || url.origin;
      const auth = betterAuth({
        database: drizzleAdapter(drizzle(env.DB, { schema: authSchema }), {
          provider: "sqlite",
          schema: authSchema,
        }),
        secret: env.BETTER_AUTH_SECRET,
        baseURL,
        trustedOrigins: [new URL(baseURL).origin],
        emailAndPassword: {
          enabled: true,
          minPasswordLength: 10,
          maxPasswordLength: 128,
        },
        rateLimit: { enabled: true },
        advanced: {
          defaultCookieAttributes: {
            httpOnly: true,
            sameSite: "lax",
            secure: new URL(baseURL).protocol === "https:",
          },
        },
      });
      return auth.handler(request);
    }

    if (!url.pathname.startsWith("/api/")) return env.ASSETS.fetch(request);
    if (!env.BETTER_AUTH_SECRET) return fail("Authentication is not configured.", 503);

    const isMutation = !["GET", "HEAD", "OPTIONS"].includes(request.method);
    if (isMutation) {
      const origin = request.headers.get("Origin");
      const expectedOrigin = new URL(env.BETTER_AUTH_URL || url.origin).origin;
      if (!origin || origin !== expectedOrigin) return fail("Request origin is not allowed.", 403);
      if (!(request.headers.get("Content-Type") || "").toLowerCase().startsWith("application/json")) {
        return fail("Content-Type must be application/json.", 415);
      }
    }

    const auth = betterAuth({
      database: drizzleAdapter(drizzle(env.DB, { schema: authSchema }), {
        provider: "sqlite",
        schema: authSchema,
      }),
      secret: env.BETTER_AUTH_SECRET,
      baseURL: env.BETTER_AUTH_URL || url.origin,
      trustedOrigins: [new URL(env.BETTER_AUTH_URL || url.origin).origin],
      emailAndPassword: { enabled: true, minPasswordLength: 10, maxPasswordLength: 128 },
      rateLimit: { enabled: true },
    });
    const sessionResult = await auth.api.getSession({ headers: request.headers }).catch(() => null);
    const currentUser = sessionResult?.user as SessionUser | undefined;
    if (!currentUser?.id) return fail("Please sign in to access your financial records.", 401);
    const userId = currentUser.id;
    const db = env.DB;

    try {
      if (url.pathname === "/api/me" && request.method === "GET") {
        return json({ user: { id: currentUser.id, name: currentUser.name, email: currentUser.email }, timezone: "Asia/Manila", currency: "PHP" });
      }

      if (url.pathname === "/api/dashboard" && request.method === "GET") {
        const start = monthStartInManila();
        const today = todayInManila();
        const sums = await db.prepare(`
          SELECT
            COALESCE(SUM(CASE WHEN transaction_type='income' THEN amount_centavos ELSE 0 END),0) AS income,
            COALESCE(SUM(CASE WHEN transaction_type='expense' THEN amount_centavos ELSE 0 END),0) AS expenses,
            COALESCE(SUM(CASE WHEN transaction_type='savings_contribution' THEN amount_centavos ELSE 0 END),0) AS savings_contributions
          FROM transactions WHERE user_id=? AND transaction_date BETWEEN ? AND ?
        `).bind(userId, start, today).first<{ income: number; expenses: number; savings_contributions: number }>();
        const goals = await db.prepare("SELECT COALESCE(SUM(current_amount_centavos),0) AS saved FROM savings_goals WHERE user_id=? AND status != 'paused'").bind(userId).first<{ saved: number }>();
        const recent = await db.prepare(`
          SELECT t.id,t.account_id,t.category_id,t.goal_id,t.transaction_type,t.amount_centavos,t.transaction_date,t.description,t.payment_method,t.notes,
                 a.name AS account_name,c.name AS category_name
          FROM transactions t JOIN accounts a ON a.id=t.account_id AND a.user_id=t.user_id
          LEFT JOIN categories c ON c.id=t.category_id AND c.user_id=t.user_id
          WHERE t.user_id=? ORDER BY t.transaction_date DESC,t.created_at DESC LIMIT 6
        `).bind(userId).all<TxRow>();
        return json({
          period: { start, end: today, timezone: "Asia/Manila" },
          totals: {
            income: Number(sums?.income ?? 0),
            expenses: Number(sums?.expenses ?? 0),
            savings: Number(sums?.savings_contributions ?? 0),
            savingsContributions: Number(sums?.savings_contributions ?? 0),
            remainingBalance: Number(sums?.income ?? 0) - Number(sums?.expenses ?? 0) - Number(sums?.savings_contributions ?? 0),
            goalSavings: Number(goals?.saved ?? 0),
          },
          recentTransactions: recent.results ?? [],
        });
      }

      if (url.pathname === "/api/accounts" && request.method === "GET") {
        const rows = await db.prepare(`
          SELECT a.id,a.name,a.account_type,a.opening_balance_centavos,
            a.opening_balance_centavos + COALESCE(SUM(CASE WHEN t.transaction_type='income' THEN t.amount_centavos WHEN t.transaction_type IN ('expense','savings_contribution') THEN -t.amount_centavos ELSE 0 END),0) AS balance_centavos
          FROM accounts a LEFT JOIN transactions t ON t.account_id=a.id AND t.user_id=a.user_id
          WHERE a.user_id=? GROUP BY a.id ORDER BY a.created_at
        `).bind(userId).all();
        return json({ data: rows.results ?? [] });
      }

      if (url.pathname === "/api/accounts" && request.method === "POST") {
        const body = await readJson(request);
        const name = typeof body?.name === "string" ? body.name.trim() : "";
        const type = body?.accountType;
        const opening = body?.openingBalance === undefined || body?.openingBalance === "" ? 0 :
          (typeof body.openingBalance === "number" && Number.isFinite(body.openingBalance) && body.openingBalance >= 0 && body.openingBalance <= 1000000000 ? Math.round(body.openingBalance * 100) : null);
        if (!name || name.length > 60) return fail("Enter an account name of 1–60 characters.");
        if (type !== "cash" && type !== "bank" && type !== "e-wallet" && type !== "credit") return fail("Choose a valid account type.");
        if (opening === null) return fail("Opening balance must be a non-negative amount of up to ₱1,000,000,000.");
        const accountId = id();
        await db.prepare("INSERT INTO accounts (id,user_id,name,account_type,opening_balance_centavos) VALUES (?,?,?,?,?)").bind(accountId,userId,name,type,opening).run();
        return json({ data: { id: accountId, name, account_type: type, opening_balance_centavos: opening } }, 201);
      }

      if (url.pathname === "/api/categories" && request.method === "GET") {
        const rows = await db.prepare("SELECT id,name,kind,color,budget_group FROM categories WHERE user_id=? ORDER BY budget_group,kind,name").bind(userId).all();
        return json({ data: rows.results ?? [] });
      }

      if (url.pathname === "/api/categories" && request.method === "POST") {
        const body = await readJson(request);
        const name = typeof body?.name === "string" ? body.name.trim() : "";
        const kind = body?.kind;
        const group = body?.budgetGroup;
        if (!name || name.length > 60) return fail("Enter a category name of 1–60 characters.");
        if (kind !== "income" && kind !== "expense") return fail("Choose income or expense for the category.");
        if (!["needs","wants","savings","income"].includes(group) || (kind === "income" && group !== "income") || (kind === "expense" && group === "income")) return fail("Choose a category group matching its type.");
        const categoryId = id();
        try {
          await db.prepare("INSERT INTO categories (id,user_id,name,kind,budget_group) VALUES (?,?,?,?,?)").bind(categoryId,userId,name,kind,group).run();
        } catch { return fail("That category already exists for this type.", 409); }
        return json({ data: { id: categoryId, name, kind, budget_group: group } }, 201);
      }

      const categoryMatch = url.pathname.match(/^\/api\/categories\/([a-f0-9-]+)$/i);
      if (categoryMatch && request.method === "PATCH") {
        const body = await readJson(request);
        const name = typeof body?.name === "string" ? body.name.trim() : "";
        const kind = body?.kind;
        const group = body?.budgetGroup;
        if (!name || name.length > 60) return fail("Enter a category name of 1–60 characters.");
        if (kind !== "income" && kind !== "expense") return fail("Choose income or expense for the category.");
        if (!["needs","wants","savings","income"].includes(group) || (kind === "income" && group !== "income") || (kind === "expense" && group === "income")) return fail("Choose a category group matching its type.");
        const result = await db.prepare("UPDATE categories SET name=?,kind=?,budget_group=? WHERE id=? AND user_id=?").bind(name,kind,group,categoryMatch[1],userId).run();
        if (!result.meta.changes) return fail("Category not found.",404);
        return json({ok:true});
      }
      if (categoryMatch && request.method === "DELETE") {
        const used = await db.prepare("SELECT id FROM transactions WHERE category_id=? AND user_id=? LIMIT 1").bind(categoryMatch[1],userId).first();
        const budgeted = await db.prepare("SELECT id FROM budgets WHERE category_id=? AND user_id=? LIMIT 1").bind(categoryMatch[1],userId).first();
        if (used || budgeted) return fail("This category is used by transactions or budgets. Edit its name/group instead, or remove those records first.",409);
        const result = await db.prepare("DELETE FROM categories WHERE id=? AND user_id=?").bind(categoryMatch[1],userId).run();
        if (!result.meta.changes) return fail("Category not found.",404);
        return json({ok:true});
      }

      if (url.pathname === "/api/transactions" && request.method === "GET") {
        const from = url.searchParams.get("from");
        const to = url.searchParams.get("to");
        const kind = url.searchParams.get("type");
        const categoryId = url.searchParams.get("categoryId");
        const paymentMethod = url.searchParams.get("paymentMethod");
        if ((from && !validDate(from)) || (to && !validDate(to))) return fail("Use dates in YYYY-MM-DD format.");
        if (from && to && from > to) return fail("The end date must be on or after the start date.");
        if (kind && !["income", "expense", "savings_contribution", "transfer"].includes(kind)) return fail("Choose a valid transaction type.");
        if (paymentMethod && !["cash", "bank_transfer", "e_wallet", "credit_card", "other"].includes(paymentMethod)) return fail("Choose a valid payment method.");
        let sql = `SELECT t.id,t.account_id,t.category_id,t.transaction_type,t.amount_centavos,t.transaction_date,t.description,t.payment_method,t.notes,
          a.name AS account_name,c.name AS category_name FROM transactions t
          JOIN accounts a ON a.id=t.account_id AND a.user_id=t.user_id
          LEFT JOIN categories c ON c.id=t.category_id AND c.user_id=t.user_id
          WHERE t.user_id=?`;
        const args: (string | number)[] = [userId];
        if (from) { sql += " AND t.transaction_date>=?"; args.push(from); }
        if (to) { sql += " AND t.transaction_date<=?"; args.push(to); }
        if (kind) { sql += " AND t.transaction_type=?"; args.push(kind); }
        if (categoryId) { sql += " AND t.category_id=?"; args.push(categoryId); }
        if (paymentMethod) { sql += " AND t.payment_method=?"; args.push(paymentMethod); }
        sql += " ORDER BY t.transaction_date DESC,t.created_at DESC";
        const rows = await db.prepare(sql).bind(...args).all<TxRow>();
        return json({ data: rows.results ?? [] });
      }

      if (url.pathname === "/api/transactions" && request.method === "POST") {
        const body = await readJson(request);
        const accountId = typeof body?.accountId === "string" ? body.accountId : "";
        const categoryId = typeof body?.categoryId === "string" ? body.categoryId : "";
        const kind = body?.type;
        const cents = amountCentavos(body?.amount);
        const date = body?.date;
        const description = typeof body?.description === "string" ? body.description.trim() : "";
        const paymentMethod = body?.paymentMethod;
        const notes = typeof body?.notes === "string" ? body.notes.trim() : "";
        const goalId = typeof body?.goalId === "string" && body.goalId ? body.goalId : null;
        if (kind !== "income" && kind !== "expense" && kind !== "savings_contribution") return fail("Choose income, expense, or savings contribution.");
        if (kind !== "savings_contribution" && goalId) return fail("Only savings contributions can be linked to a goal.");
        if (goalId && !await db.prepare("SELECT id FROM savings_goals WHERE id=? AND user_id=?").bind(goalId,userId).first()) return fail("Choose one of your own savings goals.",403);
        if (cents === null) return fail("Enter an amount greater than ₱0.00 and no more than ₱1,000,000,000.00.");
        if (!validDate(date)) return fail("Choose a valid transaction date.");
        if (date > todayInManila()) return fail("Transaction date cannot be in the future (Asia/Manila).");
        if (!description || description.length > 180) return fail("Description is required and must be 180 characters or fewer.");
        if (notes.length > 1000) return fail("Notes must be 1,000 characters or fewer.");
        if (!validPaymentMethod(paymentMethod)) return fail("Choose a valid payment method.");
        const ownedAccount = await db.prepare("SELECT id FROM accounts WHERE id=? AND user_id=?").bind(accountId,userId).first();
        if (!ownedAccount) return fail("Choose one of your own accounts.", 403);
        const categoryKind = kind === "savings_contribution" ? "expense" : kind;
        const category = await db.prepare("SELECT id FROM categories WHERE id=? AND user_id=? AND kind=?").bind(categoryId,userId,categoryKind).first();
        if (!category) return fail("Choose a category belonging to your account and transaction type.", 403);
        const transactionId = id();
        await db.prepare("INSERT INTO transactions (id,user_id,account_id,category_id,goal_id,transaction_type,amount_centavos,transaction_date,description,payment_method,notes) VALUES (?,?,?,?,?,?,?,?,?,?,?)")
          .bind(transactionId,userId,accountId,categoryId,goalId,kind,cents,date,description,paymentMethod,notes || null).run();
        return json({ data: { id: transactionId, account_id: accountId, category_id: categoryId, goal_id: goalId, transaction_type: kind, amount_centavos: cents, transaction_date: date, description, payment_method: paymentMethod, notes: notes || null } }, 201);
      }

      const transactionMatch = url.pathname.match(/^\/api\/transactions\/([a-f0-9-]+)$/i);
      if (transactionMatch && request.method === "PUT") {
        const body = await readJson(request);
        const accountId = typeof body?.accountId === "string" ? body.accountId : "";
        const categoryId = typeof body?.categoryId === "string" ? body.categoryId : "";
        const kind = body?.type;
        const cents = amountCentavos(body?.amount);
        const date = body?.date;
        const description = typeof body?.description === "string" ? body.description.trim() : "";
        const paymentMethod = body?.paymentMethod;
        const notes = typeof body?.notes === "string" ? body.notes.trim() : "";
        if (kind !== "income" && kind !== "expense" && kind !== "savings_contribution") return fail("Choose income, expense, or savings contribution.");
        if (cents === null) return fail("Enter an amount greater than ₱0.00 and no more than ₱1,000,000,000.00.");
        if (!validDate(date)) return fail("Choose a valid transaction date.");
        if (date > todayInManila()) return fail("Transaction date cannot be in the future (Asia/Manila).");
        if (!description || description.length > 180) return fail("Description is required and must be 180 characters or fewer.");
        if (notes.length > 1000) return fail("Notes must be 1,000 characters or fewer.");
        if (!["cash", "bank_transfer", "e_wallet", "credit_card", "other"].includes(String(paymentMethod))) return fail("Choose a valid payment method.");
        const ownedAccount = await db.prepare("SELECT id FROM accounts WHERE id=? AND user_id=?").bind(accountId,userId).first();
        if (!ownedAccount) return fail("Choose one of your own accounts.", 403);
        const categoryKind = kind === "savings_contribution" ? "expense" : kind;
        const category = await db.prepare("SELECT id FROM categories WHERE id=? AND user_id=? AND kind=?").bind(categoryId,userId,categoryKind).first();
        if (!category) return fail("Choose a category belonging to your account and transaction type.", 403);
        const result = await db.prepare("UPDATE transactions SET account_id=?,category_id=?,goal_id=?,transaction_type=?,amount_centavos=?,transaction_date=?,description=?,payment_method=?,notes=?,updated_at=CURRENT_TIMESTAMP WHERE id=? AND user_id=?")
          .bind(accountId,categoryId,goalId,kind,cents,date,description,paymentMethod,notes || null,transactionMatch[1],userId).run();
        if (!result.meta.changes) return fail("Transaction not found.", 404);
        return json({ ok: true });
      }
      if (transactionMatch && request.method === "DELETE") {
        const result = await db.prepare("DELETE FROM transactions WHERE id=? AND user_id=?").bind(transactionMatch[1],userId).run();
        if (!result.meta.changes) return fail("Transaction not found.", 404);
        return json({ ok: true });
      }

      if (url.pathname === "/api/budgets" && request.method === "GET") {
        const rows = await db.prepare(`
          SELECT b.id,b.category_id,c.name AS category_name,b.period_start,b.period_end,b.amount_centavos,
            COALESCE(SUM(t.amount_centavos),0) AS spent_centavos
          FROM budgets b JOIN categories c ON c.id=b.category_id AND c.user_id=b.user_id
          LEFT JOIN transactions t ON t.user_id=b.user_id AND t.category_id=b.category_id AND t.transaction_type='expense' AND t.transaction_date BETWEEN b.period_start AND b.period_end
          WHERE b.user_id=? GROUP BY b.id ORDER BY b.period_start DESC,b.created_at DESC
        `).bind(userId).all();
        return json({ data: rows.results ?? [] });
      }

      if (url.pathname === "/api/budgets" && request.method === "POST") {
        const body = await readJson(request);
        const categoryId = typeof body?.categoryId === "string" ? body.categoryId : "";
        const cents = amountCentavos(body?.amount);
        const start = body?.periodStart, end = body?.periodEnd;
        if (cents === null) return fail("Enter a budget amount greater than ₱0.00.");
        if (!validDate(start) || !validDate(end) || start > end) return fail("Choose a valid period; the end date must not be before the start date.");
        const category = await db.prepare("SELECT id FROM categories WHERE id=? AND user_id=? AND kind='expense'").bind(categoryId,userId).first();
        if (!category) return fail("Choose one of your own expense categories.", 403);
        const budgetId = id();
        try {
          await db.prepare("INSERT INTO budgets (id,user_id,category_id,period_start,period_end,amount_centavos) VALUES (?,?,?,?,?,?)").bind(budgetId,userId,categoryId,start,end,cents).run();
        } catch { return fail("A budget already exists for that category and date range.", 409); }
        return json({ data: { id: budgetId, category_id: categoryId, period_start: start, period_end: end, amount_centavos: cents } }, 201);
      }

      if (url.pathname === "/api/savings-goals" && request.method === "GET") {
        const rows = await db.prepare("SELECT id,name,target_amount_centavos,target_date,current_amount_centavos,status FROM savings_goals WHERE user_id=? ORDER BY created_at DESC").bind(userId).all();
        return json({ data: rows.results ?? [] });
      }

      if (url.pathname === "/api/savings-goals" && request.method === "POST") {
        const body = await readJson(request);
        const name = typeof body?.name === "string" ? body.name.trim() : "";
        const target = amountCentavos(body?.targetAmount);
        const rawTargetDate = body?.targetDate;
        if (!name || name.length > 80) return fail("Enter a savings goal name of 1–80 characters.");
        if (target === null) return fail("Target amount must be greater than ₱0.00.");
        if (rawTargetDate !== undefined && rawTargetDate !== null && rawTargetDate !== "" && !validDate(rawTargetDate)) return fail("Choose a valid target date.");
        const targetDate = typeof rawTargetDate === "string" && rawTargetDate ? rawTargetDate : null;
        const goalId = id();
        await db.prepare("INSERT INTO savings_goals (id,user_id,name,target_amount_centavos,target_date) VALUES (?,?,?,?,?)").bind(goalId,userId,name,target,targetDate).run();
        return json({ data: { id: goalId, name, target_amount_centavos: target, target_date: targetDate, current_amount_centavos: 0, status: "active" } }, 201);
      }

      const goalMatch = url.pathname.match(/^\/api\/savings-goals\/([a-f0-9-]+)$/i);
      if (goalMatch && request.method === "PATCH") {
        const body = await readJson(request);
        const saved = body?.currentAmount;
        const cents = typeof saved === "number" && Number.isFinite(saved) && saved >= 0 && saved <= 1000000000 ? Math.round(saved * 100) : null;
        const status = body?.status;
        if (cents === null) return fail("Saved amount must be from ₱0.00 to ₱1,000,000,000.00.");
        if (status !== undefined && status !== "active" && status !== "paused" && status !== "completed") return fail("Choose a valid goal status.");
        const existing = await db.prepare("SELECT id FROM savings_goals WHERE id=? AND user_id=?").bind(goalMatch[1],userId).first();
        if (!existing) return fail("Savings goal not found.", 404);
        if (status) {
          await db.prepare("UPDATE savings_goals SET current_amount_centavos=?,status=?,updated_at=CURRENT_TIMESTAMP WHERE id=? AND user_id=?").bind(cents,status,goalMatch[1],userId).run();
        } else {
          await db.prepare("UPDATE savings_goals SET current_amount_centavos=?,updated_at=CURRENT_TIMESTAMP WHERE id=? AND user_id=?").bind(cents,goalMatch[1],userId).run();
        }
        return json({ ok: true });
      }

      return fail("API endpoint not found.", 404);
    } catch (error) {
      console.error("Budgetwise API error", error instanceof Error ? error.message : "Unknown error");
      return fail("The request could not be completed. Please try again.", 500);
    }
  },
} satisfies ExportedHandler<Env>;
