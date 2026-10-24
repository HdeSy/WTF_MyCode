import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { DatabaseSync } from "node:sqlite";

const migration = (name: string) => readFileSync(new URL(`../migrations/${name}`, import.meta.url), "utf8");

describe("Cloudflare D1 migration safety", () => {
  it("applies the complete schema and preserves existing financial records through the transaction-table rebuild", () => {
    const db = new DatabaseSync(":memory:");
    try {
      db.exec(migration("0000_auth_tables.sql"));
      db.exec(migration("0001_finance_tables.sql"));

      db.prepare("INSERT INTO user (id,name,email,emailVerified,createdAt,updatedAt) VALUES (?,?,?,?,?,?)")
        .run("user-1", "Migration Test", "migration-test@example.invalid", 0, 1000, 1000);
      db.prepare("INSERT INTO accounts (id,user_id,name,account_type,opening_balance_centavos) VALUES (?,?,?,?,?)")
        .run("account-1", "user-1", "Cash", "cash", 50000);
      db.prepare("INSERT INTO categories (id,user_id,name,kind) VALUES (?,?,?,?)")
        .run("category-1", "user-1", "Groceries", "expense");
      db.prepare("INSERT INTO transactions (id,user_id,account_id,category_id,transaction_type,amount_centavos,transaction_date,description) VALUES (?,?,?,?,?,?,?,?)")
        .run("transaction-1", "user-1", "account-1", "category-1", "expense", 12500, "2026-10-20", "Legacy groceries");
      db.prepare("INSERT INTO budgets (id,user_id,category_id,period_start,period_end,amount_centavos) VALUES (?,?,?,?,?,?)")
        .run("budget-1", "user-1", "category-1", "2026-10-19", "2026-10-25", 30000);
      db.prepare("INSERT INTO savings_goals (id,user_id,name,target_amount_centavos,current_amount_centavos) VALUES (?,?,?,?,?)")
        .run("goal-1", "user-1", "Emergency fund", 100000, 25000);

      db.exec(migration("0002_transaction_enhancements.sql"));
      const preserved = db.prepare("SELECT id,user_id,account_id,category_id,transaction_type,amount_centavos,transaction_date,description,payment_method,notes FROM transactions WHERE id=?")
        .get("transaction-1") as Record<string, unknown>;
      assert.equal(preserved.id, "transaction-1");
      assert.equal(preserved.user_id, "user-1");
      assert.equal(preserved.amount_centavos, 12500);
      assert.equal(preserved.transaction_date, "2026-10-20");
      assert.equal(preserved.description, "Legacy groceries");
      assert.equal(preserved.payment_method, "cash");
      assert.equal(preserved.notes, null);

      db.exec(migration("0003_budget_categories_goals.sql"));
      const category = db.prepare("SELECT budget_group FROM categories WHERE id=?").get("category-1") as { budget_group: string };
      assert.equal(category.budget_group, "needs");
      const goal = db.prepare("SELECT target_amount_centavos,current_amount_centavos,description FROM savings_goals WHERE id=?").get("goal-1") as Record<string, unknown>;
      assert.equal(goal.target_amount_centavos, 100000);
      assert.equal(goal.current_amount_centavos, 25000);
      assert.equal(goal.description, null);

      db.prepare("INSERT INTO transactions (id,user_id,account_id,category_id,goal_id,transaction_type,amount_centavos,transaction_date,description,payment_method) VALUES (?,?,?,?,?,?,?,?,?,?)")
        .run("transaction-2", "user-1", "account-1", "category-1", "goal-1", "savings_contribution", 5000, "2026-10-21", "Goal contribution", "bank_transfer");
      const count = db.prepare("SELECT COUNT(*) AS count FROM transactions WHERE user_id=?").get("user-1") as { count: number };
      assert.equal(count.count, 2);
      const budget = db.prepare("SELECT amount_centavos FROM budgets WHERE id=?").get("budget-1") as { amount_centavos: number };
      assert.equal(budget.amount_centavos, 30000);
      const indexes = db.prepare("SELECT name FROM sqlite_master WHERE type='index'").all() as Array<{ name: string }>;
      assert.ok(indexes.some((index) => index.name === "idx_transactions_user_goal_date"));
    } finally {
      db.close();
    }
  });
});
