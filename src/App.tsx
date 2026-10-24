import { FormEvent, useCallback, useEffect, useMemo, useState } from "react";
import ExcelJS from "exceljs";
import {
  ArrowDownLeft, ArrowDownRight, ArrowLeftRight, ArrowUpRight, BadgeCheck,
  BarChart3, CalendarDays, Check, ChevronRight, CircleHelp, Clock3, CreditCard,
  Download, Ellipsis, FileSpreadsheet, Goal, LayoutDashboard, LoaderCircle,
  LogOut, Menu, Plus, Search, Settings, ShieldCheck, Trash2, TrendingUp,
  Wallet, X,
} from "lucide-react";

type Page = "Dashboard" | "Transactions" | "Budgets" | "Savings Goals" | "Reports" | "Settings";
type User = { id: string; name: string; email: string };
type Account = { id: string; name: string; account_type: string; opening_balance_centavos: number; balance_centavos?: number };
type Category = { id: string; name: string; kind: "income" | "expense"; color?: string | null };
type Transaction = { id: string; account_id: string; category_id: string | null; transaction_type: "income" | "expense" | "transfer"; amount_centavos: number; transaction_date: string; description: string | null; account_name?: string; category_name?: string | null };
type Budget = { id: string; category_id: string; category_name: string; period_start: string; period_end: string; amount_centavos: number; spent_centavos: number };
type Goal = { id: string; name: string; target_amount_centavos: number; target_date: string | null; current_amount_centavos: number; status: "active" | "completed" | "paused" };
type Dashboard = { period: { start: string; end: string; timezone: string }; totals: { income: number; expenses: number; savings: number; remainingBalance: number; goalSavings: number }; recentTransactions: Transaction[] };
type ApiResult<T> = { data: T; error?: string };

const navItems: { label: Page; icon: typeof LayoutDashboard; description: string }[] = [
  { label: "Dashboard", icon: LayoutDashboard, description: "Your financial overview" },
  { label: "Transactions", icon: ArrowLeftRight, description: "Income and expenses" },
  { label: "Budgets", icon: CreditCard, description: "Plan your spending" },
  { label: "Savings Goals", icon: Goal, description: "Build toward what matters" },
  { label: "Reports", icon: BarChart3, description: "Your financial reports" },
  { label: "Settings", icon: Settings, description: "Accounts and preferences" },
];
const money = (centavos: number) => new Intl.NumberFormat("en-PH", { style: "currency", currency: "PHP", minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(centavos / 100);
const peso = (amount: number) => new Intl.NumberFormat("en-PH", { style: "currency", currency: "PHP", minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(amount);
function manilaToday() {
  const p = new Intl.DateTimeFormat("en-CA", { timeZone: "Asia/Manila", year: "numeric", month: "2-digit", day: "2-digit" }).formatToParts(new Date());
  const get = (type: string) => p.find((x) => x.type === type)?.value ?? "";
  return `${get("year")}-${get("month")}-${get("day")}`;
}
function manilaMonthStart() { return manilaToday().slice(0, 8) + "01"; }
function reportWindow(period: "Daily" | "Weekly" | "Monthly") {
  const today = manilaToday();
  if (period === "Daily") return { start: today, end: today };
  if (period === "Monthly") return { start: today.slice(0, 8) + "01", end: today };
  const cursor = new Date(today + "T00:00:00Z");
  const mondayOffset = (cursor.getUTCDay() + 6) % 7;
  cursor.setUTCDate(cursor.getUTCDate() - mondayOffset);
  return { start: cursor.toISOString().slice(0, 10), end: today };
}
function prettyDate(value: string) {
  if (!value) return "—";
  return new Intl.DateTimeFormat("en-PH", { dateStyle: "medium", timeZone: "Asia/Manila" }).format(new Date(value.slice(0, 10) + "T12:00:00+08:00"));
}
async function api<T>(path: string, init: RequestInit = {}): Promise<T> {
  const response = await fetch(path, {
    ...init,
    credentials: "same-origin",
    headers: { ...(!["GET", "HEAD"].includes((init.method || "GET").toUpperCase()) ? { "Content-Type": "application/json" } : {}), ...(init.headers || {}) },
  });
  const body = await response.json().catch(() => ({})) as { error?: string } & T;
  if (!response.ok) throw new Error(body.error || "Something went wrong. Please try again.");
  return body;
}

export default function App() {
  const [user, setUser] = useState<User | null>(null);
  const [authMode, setAuthMode] = useState<"signin" | "signup">("signin");
  const [authForm, setAuthForm] = useState({ name: "", email: "", password: "" });
  const [page, setPage] = useState<Page>("Dashboard");
  const [reportPeriod, setReportPeriod] = useState<"Daily" | "Weekly" | "Monthly">("Monthly");
  const [mobileNav, setMobileNav] = useState(false);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const [success, setSuccess] = useState("");
  const [dashboard, setDashboard] = useState<Dashboard | null>(null);
  const [transactions, setTransactions] = useState<Transaction[]>([]);
  const [accounts, setAccounts] = useState<Account[]>([]);
  const [categories, setCategories] = useState<Category[]>([]);
  const [budgets, setBudgets] = useState<Budget[]>([]);
  const [goals, setGoals] = useState<Goal[]>([]);
  const [search, setSearch] = useState("");
  const [transactionFilter, setTransactionFilter] = useState("all");
  const [transactionForm, setTransactionForm] = useState({ type: "expense", amount: "", date: manilaToday(), accountId: "", categoryId: "", description: "" });
  const [budgetForm, setBudgetForm] = useState({ categoryId: "", amount: "", periodStart: manilaMonthStart(), periodEnd: manilaToday() });
  const [goalForm, setGoalForm] = useState({ name: "", targetAmount: "", targetDate: "" });
  const [goalUpdates, setGoalUpdates] = useState<Record<string, string>>({});
  const [accountForm, setAccountForm] = useState({ name: "", accountType: "bank", openingBalance: "0" });
  const [categoryForm, setCategoryForm] = useState({ name: "", kind: "expense" });

  const refreshAll = useCallback(async () => {
    setLoading(true);
    setError("");
    try {
      const [me, dash, tx, acc, cats, buds, savings] = await Promise.all([
        api<{ user: User }>("/api/me"),
        api<Dashboard>("/api/dashboard"),
        api<ApiResult<Transaction[]>>("/api/transactions"),
        api<ApiResult<Account[]>>("/api/accounts"),
        api<ApiResult<Category[]>>("/api/categories"),
        api<ApiResult<Budget[]>>("/api/budgets"),
        api<ApiResult<Goal[]>>("/api/savings-goals"),
      ]);
      setUser(me.user);
      setDashboard(dash);
      setTransactions(tx.data);
      setAccounts(acc.data);
      setCategories(cats.data);
      setBudgets(buds.data);
      setGoals(savings.data);
      setTransactionForm((old) => ({ ...old, accountId: old.accountId || acc.data[0]?.id || "" }));
      setBudgetForm((old) => ({ ...old, categoryId: old.categoryId || cats.data.find((c) => c.kind === "expense")?.id || "" }));
    } catch (e) {
      if (e instanceof Error && e.message.toLowerCase().includes("sign in")) {
        setUser(null);
      } else {
        setError(e instanceof Error ? e.message : "Unable to load your financial records.");
      }
    } finally { setLoading(false); }
  }, []);

  useEffect(() => {
    let cancelled = false;
    api<{ user: User }>("/api/me")
      .then((result) => { if (!cancelled) { setUser(result.user); void refreshAll(); } })
      .catch(() => { if (!cancelled) setLoading(false); });
    return () => { cancelled = true; };
  }, [refreshAll]);

  const flash = (message: string) => { setSuccess(message); setError(""); };
  const handleAuth = async (event: FormEvent) => {
    event.preventDefault(); setError(""); setSuccess("");
    if (!authForm.email.trim() || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(authForm.email)) { setError("Enter a valid email address."); return; }
    if (authForm.password.length < 10) { setError("Use a password with at least 10 characters."); return; }
    if (authMode === "signup" && !authForm.name.trim()) { setError("Enter your name."); return; }
    setSaving(true);
    try {
      await api("/api/auth/" + (authMode === "signup" ? "sign-up/email" : "sign-in/email"), {
        method: "POST",
        body: JSON.stringify(authMode === "signup"
          ? { name: authForm.name.trim(), email: authForm.email.trim(), password: authForm.password }
          : { email: authForm.email.trim(), password: authForm.password }),
      });
      await refreshAll();
      flash(authMode === "signup" ? "Your account was created successfully." : "Welcome back.");
    } catch (e) { setError(e instanceof Error ? e.message : "Unable to authenticate."); }
    finally { setSaving(false); }
  };

  const signOut = async () => {
    setSaving(true);
    try { await api("/api/auth/sign-out", { method: "POST", body: JSON.stringify({}) }); }
    catch { /* local state is cleared even if the server session already expired */ }
    setUser(null); setDashboard(null); setTransactions([]); setAccounts([]); setCategories([]); setBudgets([]); setGoals([]);
    setPage("Dashboard"); setSuccess(""); setError(""); setSaving(false);
  };

  const submitTransaction = async (event: FormEvent) => {
    event.preventDefault(); setError(""); setSuccess("");
    const amount = Number(transactionForm.amount);
    if (!Number.isFinite(amount) || amount <= 0 || amount > 1_000_000_000) { setError("Enter an amount greater than ₱0.00 and no more than ₱1,000,000,000.00."); return; }
    if (!transactionForm.accountId) { setError("Create an account in Settings before adding transactions."); return; }
    if (!transactionForm.date) { setError("Choose a transaction date."); return; }
    setSaving(true);
    try {
      await api("/api/transactions", { method: "POST", body: JSON.stringify({ type: transactionForm.type, amount, date: transactionForm.date, accountId: transactionForm.accountId, categoryId: transactionForm.categoryId || null, description: transactionForm.description.trim() }) });
      setTransactionForm((old) => ({ ...old, amount: "", description: "", categoryId: "" }));
      await refreshAll(); setPage("Transactions"); flash("Transaction saved to your private account.");
    } catch (e) { setError(e instanceof Error ? e.message : "Could not save transaction."); }
    finally { setSaving(false); }
  };

  const deleteTransaction = async (transactionId: string) => {
    if (!window.confirm("Delete this transaction? This cannot be undone.")) return;
    setSaving(true); setError(""); setSuccess("");
    try { await api("/api/transactions/" + transactionId, { method: "DELETE" }); await refreshAll(); flash("Transaction deleted."); }
    catch (e) { setError(e instanceof Error ? e.message : "Could not delete transaction."); }
    finally { setSaving(false); }
  };

  const submitBudget = async (event: FormEvent) => {
    event.preventDefault(); setError(""); setSuccess("");
    const amount = Number(budgetForm.amount);
    if (!budgetForm.categoryId) { setError("Create an expense category in Settings first."); return; }
    if (!Number.isFinite(amount) || amount <= 0) { setError("Enter a budget greater than ₱0.00."); return; }
    if (!budgetForm.periodStart || !budgetForm.periodEnd || budgetForm.periodEnd < budgetForm.periodStart) { setError("The end date must be on or after the start date."); return; }
    setSaving(true);
    try { await api("/api/budgets", { method: "POST", body: JSON.stringify({ categoryId: budgetForm.categoryId, amount, periodStart: budgetForm.periodStart, periodEnd: budgetForm.periodEnd }) }); setBudgetForm((old) => ({ ...old, amount: "" })); await refreshAll(); flash("Budget saved."); }
    catch (e) { setError(e instanceof Error ? e.message : "Could not save budget."); }
    finally { setSaving(false); }
  };

  const submitGoal = async (event: FormEvent) => {
    event.preventDefault(); setError(""); setSuccess("");
    const targetAmount = Number(goalForm.targetAmount);
    if (!goalForm.name.trim()) { setError("Enter a name for your savings goal."); return; }
    if (!Number.isFinite(targetAmount) || targetAmount <= 0) { setError("Enter a target amount greater than ₱0.00."); return; }
    setSaving(true);
    try { await api("/api/savings-goals", { method: "POST", body: JSON.stringify({ name: goalForm.name.trim(), targetAmount, targetDate: goalForm.targetDate || null }) }); setGoalForm({ name: "", targetAmount: "", targetDate: "" }); await refreshAll(); flash("Savings goal created."); }
    catch (e) { setError(e instanceof Error ? e.message : "Could not create savings goal."); }
    finally { setSaving(false); }
  };

  const updateGoal = async (goal: Goal) => {
    const raw = goalUpdates[goal.id] ?? (goal.current_amount_centavos / 100).toFixed(2);
    const currentAmount = Number(raw);
    if (!Number.isFinite(currentAmount) || currentAmount < 0 || currentAmount > 1_000_000_000) { setError("Saved amount must be between ₱0.00 and ₱1,000,000,000.00."); return; }
    setSaving(true); setError(""); setSuccess("");
    try { await api("/api/savings-goals/" + goal.id, { method: "PATCH", body: JSON.stringify({ currentAmount, status: currentAmount >= goal.target_amount_centavos / 100 ? "completed" : "active" }) }); await refreshAll(); flash("Savings progress updated."); }
    catch (e) { setError(e instanceof Error ? e.message : "Could not update savings goal."); }
    finally { setSaving(false); }
  };

  const createAccount = async (event: FormEvent) => {
    event.preventDefault(); setError(""); setSuccess("");
    const openingBalance = Number(accountForm.openingBalance);
    if (!accountForm.name.trim()) { setError("Enter an account name."); return; }
    if (!Number.isFinite(openingBalance) || openingBalance < 0 || openingBalance > 1_000_000_000) { setError("Enter a valid non-negative opening balance."); return; }
    setSaving(true);
    try { await api("/api/accounts", { method: "POST", body: JSON.stringify({ name: accountForm.name.trim(), accountType: accountForm.accountType, openingBalance }) }); setAccountForm({ name: "", accountType: "bank", openingBalance: "0" }); await refreshAll(); flash("Account saved."); }
    catch (e) { setError(e instanceof Error ? e.message : "Could not create account."); }
    finally { setSaving(false); }
  };

  const createCategory = async (event: FormEvent) => {
    event.preventDefault(); setError(""); setSuccess("");
    if (!categoryForm.name.trim()) { setError("Enter a category name."); return; }
    setSaving(true);
    try { await api("/api/categories", { method: "POST", body: JSON.stringify({ name: categoryForm.name.trim(), kind: categoryForm.kind }) }); setCategoryForm((old) => ({ ...old, name: "" })); await refreshAll(); flash("Category saved."); }
    catch (e) { setError(e instanceof Error ? e.message : "Could not create category."); }
    finally { setSaving(false); }
  };

  const visibleTransactions = useMemo(() => transactions.filter((tx) => {
    const text = [tx.description, tx.category_name, tx.account_name, tx.transaction_date, tx.transaction_type].join(" ").toLowerCase();
    return text.includes(search.toLowerCase()) && (transactionFilter === "all" || tx.transaction_type === transactionFilter);
  }), [transactions, search, transactionFilter]);

  const exportExcel = async () => {
    setSaving(true); setError(""); setSuccess("");
    try {
      const range = reportWindow(reportPeriod);
      const reportTransactions = transactions.filter((tx) => tx.transaction_date >= range.start && tx.transaction_date <= range.end);
      const reportIncome = reportTransactions.filter((tx) => tx.transaction_type === "income").reduce((sum, tx) => sum + tx.amount_centavos, 0);
      const reportExpenses = reportTransactions.filter((tx) => tx.transaction_type === "expense").reduce((sum, tx) => sum + tx.amount_centavos, 0);
      const reportBudgets = budgets.filter((b) => b.period_end >= range.start && b.period_start <= range.end);
      const workbook = new ExcelJS.Workbook();
      workbook.creator = "Budgetwise"; workbook.created = new Date();
      const summary = workbook.addWorksheet("Financial Summary");
      summary.columns = [{ header: "Metric", key: "metric", width: 30 }, { header: "Amount (PHP)", key: "amount", width: 22 }];
      summary.addRows([
        { metric: "Report generated (Asia/Manila)", amount: new Intl.DateTimeFormat("en-PH", { dateStyle: "medium", timeStyle: "short", timeZone: "Asia/Manila" }).format(new Date()) },
        { metric: "Report frequency", amount: reportPeriod },
        { metric: "Period", amount: range.start + " to " + range.end },
        { metric: "Total income", amount: reportIncome / 100 },
        { metric: "Total expenses", amount: reportExpenses / 100 },
        { metric: "Net cash flow", amount: (reportIncome - reportExpenses) / 100 },
        { metric: "Savings across active goals", amount: (dashboard?.totals.goalSavings ?? 0) / 100 },
        { metric: "Remaining account balance", amount: (dashboard?.totals.remainingBalance ?? 0) / 100 },
      ]);
      const txSheet = workbook.addWorksheet("Transaction Details");
      txSheet.columns = [{ header: "Date", key: "date", width: 16 }, { header: "Description", key: "description", width: 30 }, { header: "Category", key: "category", width: 22 }, { header: "Account", key: "account", width: 22 }, { header: "Type", key: "type", width: 14 }, { header: "Amount (PHP)", key: "amount", width: 18 }];
      txSheet.addRows(reportTransactions.map((tx) => ({ date: tx.transaction_date, description: tx.description || "", category: tx.category_name || "", account: tx.account_name || "", type: tx.transaction_type, amount: (tx.transaction_type === "expense" ? -tx.amount_centavos : tx.amount_centavos) / 100 })));
      txSheet.autoFilter = { from: "A1", to: "F1" };
      const budgetSheet = workbook.addWorksheet("Budget Analysis");
      budgetSheet.columns = [{ header: "Category", key: "category", width: 24 }, { header: "Period start", key: "start", width: 16 }, { header: "Period end", key: "end", width: 16 }, { header: "Budget (PHP)", key: "budget", width: 18 }, { header: "Spent (PHP)", key: "spent", width: 18 }, { header: "Remaining (PHP)", key: "remaining", width: 20 }];
      budgetSheet.addRows(reportBudgets.map((b) => ({ category: b.category_name, start: b.period_start, end: b.period_end, budget: b.amount_centavos / 100, spent: b.spent_centavos / 100, remaining: (b.amount_centavos - b.spent_centavos) / 100 })));
      const goalsSheet = workbook.addWorksheet("Savings Goals");
      goalsSheet.columns = [{ header: "Goal", key: "goal", width: 28 }, { header: "Target (PHP)", key: "target", width: 18 }, { header: "Saved (PHP)", key: "saved", width: 18 }, { header: "Progress", key: "progress", width: 14 }, { header: "Target date", key: "date", width: 18 }, { header: "Status", key: "status", width: 14 }];
      goalsSheet.addRows(goals.map((g) => ({ goal: g.name, target: g.target_amount_centavos / 100, saved: g.current_amount_centavos / 100, progress: g.target_amount_centavos ? g.current_amount_centavos / g.target_amount_centavos : 0, date: g.target_date || "", status: g.status })));
      for (const sheet of [summary, txSheet, budgetSheet, goalsSheet]) {
        sheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
        sheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF285A43" } };
        sheet.views = [{ state: "frozen", ySplit: 1 }];
        for (let i = 2; i <= sheet.rowCount; i++) sheet.getRow(i).alignment = { vertical: "middle" };
      }
      for (const sheet of [summary, txSheet, budgetSheet, goalsSheet]) {
        sheet.columns.forEach((column) => { if (column.key !== "metric" && column.key !== "date" && column.key !== "start" && column.key !== "end" && column.key !== "description" && column.key !== "category" && column.key !== "account" && column.key !== "type" && column.key !== "goal" && column.key !== "status") column.numFmt = '"₱"#,##0.00;[Red]-"₱"#,##0.00'; });
      }
      goalsSheet.getColumn("progress").numFmt = "0%";
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const objectUrl = URL.createObjectURL(blob);
      const link = document.createElement("a"); link.href = objectUrl; link.download = `budgetwise-report-${manilaToday()}.xlsx`; link.click(); URL.revokeObjectURL(objectUrl);
      flash("Your private financial report was exported.");
    } catch { setError("Excel export failed. Please try again."); }
    finally { setSaving(false); }
  };

  if (loading && !user) return <div className="full-state"><LoaderCircle className="spin" size={28} /><span>Checking your secure session…</span></div>;

  if (!user) return (
    <main className="auth-screen">
      <section className="auth-intro">
        <div className="auth-brand"><span className="brand-icon"><Wallet size={23} /></span><span>budgetwise<span className="brand-period">.</span></span></div>
        <div className="auth-copy"><span className="overline">PERSONAL FINANCE, MADE CLEAR</span><h1>A little more clarity.<br /><em>A lot more confidence.</em></h1><p>Understand your money, build better habits, and work toward the future you want — one step at a time.</p>
          <div className="auth-points"><span><BadgeCheck size={18} /> Your own private financial space</span><span><TrendingUp size={18} /> Budgets and savings goals in one place</span><span><FileSpreadsheet size={18} /> Downloadable Excel reports</span></div>
        </div>
        <p className="auth-footer">Designed for everyday money in the Philippines <span>·</span> PHP <span>·</span> Asia/Manila</p>
      </section>
      <section className="auth-form-side"><div className="auth-form-wrap">
        <div className="mobile-brand auth-brand"><span className="brand-icon"><Wallet size={23} /></span><span>budgetwise<span className="brand-period">.</span></span></div>
        <span className="overline">{authMode === "signin" ? "WELCOME BACK" : "YOUR FINANCIAL JOURNEY STARTS HERE"}</span>
        <h2>{authMode === "signin" ? "Sign in to Budgetwise" : "Create your account"}</h2>
        <p className="muted">{authMode === "signin" ? "Your finances, just where you left them." : "Set up your own secure personal finance space."}</p>
        {error && <div className="alert alert-error" role="alert">{error}</div>}
        {success && <div className="alert alert-success" role="status">{success}</div>}
        <form onSubmit={handleAuth} className="stack-form">
          {authMode === "signup" && <label>Full name<input autoComplete="name" value={authForm.name} onChange={(e) => setAuthForm({ ...authForm, name: e.target.value })} placeholder="Your name" maxLength={80} required /></label>}
          <label>Email address<input type="email" autoComplete="email" value={authForm.email} onChange={(e) => setAuthForm({ ...authForm, email: e.target.value })} placeholder="you@example.com" required /></label>
          <label>Password<input type="password" autoComplete={authMode === "signin" ? "current-password" : "new-password"} value={authForm.password} onChange={(e) => setAuthForm({ ...authForm, password: e.target.value })} placeholder="At least 10 characters" minLength={10} maxLength={128} required /></label>
          <button className="btn btn-primary btn-full" type="submit" disabled={saving}>{saving ? <LoaderCircle size={17} className="spin" /> : null}{authMode === "signin" ? "Sign in" : "Create account"}<ChevronRight size={17} /></button>
        </form>
        <p className="switch-auth">{authMode === "signin" ? "New to Budgetwise?" : "Already have an account?"} <button onClick={() => { setAuthMode(authMode === "signin" ? "signup" : "signin"); setError(""); setSuccess(""); }}> {authMode === "signin" ? "Create an account" : "Sign in"}</button></p>
        <div className="auth-security"><ShieldCheck size={17} /><span>Your records are stored privately. Each account can access only its own financial data.</span></div>
      </div></section>
    </main>
  );

  const pageDescription: Record<Page, string> = {
    Dashboard: "A calm, clear view of where your money stands.",
    Transactions: "Every inflow and outflow, all in one place.",
    Budgets: "Give every peso a purpose.",
    "Savings Goals": "Make progress toward the things that matter.",
    Reports: "Turn your financial activity into a clearer picture.",
    Settings: "Manage your accounts and personal preferences.",
  };

  return (
    <div className="app-shell">
      <aside className={mobileNav ? "sidebar sidebar-open" : "sidebar"}>
        <a className="brand" href="#dashboard" onClick={(e) => { e.preventDefault(); setPage("Dashboard"); setMobileNav(false); }}><span className="brand-icon"><Wallet size={21} /></span><span>budgetwise<span className="brand-period">.</span><small>PERSONAL FINANCE</small></span></a>
        <span className="nav-heading">WORKSPACE</span>
        <nav className="nav-list" aria-label="Main navigation">
          {navItems.map(({ label, icon: Icon }) => <button key={label} className={page === label ? "nav-item active" : "nav-item"} onClick={() => { setPage(label); setMobileNav(false); setError(""); setSuccess(""); }} aria-current={page === label ? "page" : undefined}><Icon size={18} /><span>{label}</span>{page === label && <span className="nav-active-dot" />}</button>)}
        </nav>
        <div className="sidebar-bottom">
          <div className="privacy-card"><ShieldCheck size={18} /><div><b>Your money, your space</b><small>Private records for your account.</small></div></div>
          <div className="user-row"><div className="user-avatar">{(user.name || user.email).slice(0, 1).toUpperCase()}</div><div className="user-info"><b>{user.name}</b><small>{user.email}</small></div><button className="icon-btn" onClick={signOut} aria-label="Sign out" title="Sign out"><LogOut size={17} /></button></div>
        </div>
      </aside>
      {mobileNav && <button className="mobile-scrim" aria-label="Close navigation" onClick={() => setMobileNav(false)} />}
      <main className="main-content">
        <header className="topbar"><div className="topbar-left"><button className="mobile-menu icon-btn" aria-label="Open navigation" onClick={() => setMobileNav(true)}><Menu size={20} /></button><div className="breadcrumbs"><span>Workspace</span><ChevronRight size={14} /><b>{page}</b></div></div><div className="topbar-right"><span className="timezone-pill"><span className="live-dot" /> Asia/Manila</span><button className="icon-btn help-btn" title="Help" aria-label="Help" onClick={() => { setPage("Settings"); flash("All amounts use Philippine pesos. Dates use Asia/Manila."); }}><CircleHelp size={18} /></button><div className="user-avatar small-avatar">{(user.name || user.email).slice(0, 1).toUpperCase()}</div></div></header>
        <div className="page-content">
          <div className="page-heading"><div><span className="overline">{page === "Dashboard" ? "YOUR MONEY AT A GLANCE" : "PERSONAL FINANCE"}</span><h1>{page === "Dashboard" ? "Your money, in balance." : page}</h1><p>{pageDescription[page]}</p></div>{page === "Transactions" && <button className="btn btn-primary" onClick={() => document.getElementById("transaction-entry")?.scrollIntoView({ behavior: "smooth", block: "start" })}><Plus size={17} /> Add transaction</button>}</div>

          {error && <div className="alert alert-error global-alert" role="alert"><X size={17} /><span>{error}</span><button onClick={() => setError("")} aria-label="Dismiss error"><X size={16} /></button></div>}
          {success && <div className="alert alert-success global-alert" role="status"><Check size={17} /><span>{success}</span><button onClick={() => setSuccess("")} aria-label="Dismiss success"><X size={16} /></button></div>}

          {loading ? <div className="loading-panel"><LoaderCircle className="spin" size={24} /><span>Loading your private financial records…</span></div> : <>
            {page === "Dashboard" && <DashboardPage dashboard={dashboard} accounts={accounts} goals={goals} transactions={transactions} setPage={setPage} />}
            {page === "Transactions" && <div className="page-columns">
              <section className="panel form-panel" id="transaction-entry"><PanelTitle eyebrow="NEW RECORD" title="Add a transaction" description="Amounts are stored securely in your account." /><form onSubmit={submitTransaction} className="stack-form">
                <label>Transaction type<div className="segmented"><button type="button" className={transactionForm.type === "expense" ? "selected expense-selected" : ""} onClick={() => setTransactionForm({ ...transactionForm, type: "expense", categoryId: "" })}><ArrowUpRight size={16} /> Expense</button><button type="button" className={transactionForm.type === "income" ? "selected income-selected" : ""} onClick={() => setTransactionForm({ ...transactionForm, type: "income", categoryId: "" })}><ArrowDownLeft size={16} /> Income</button></div></label>
                <label>Amount (₱)<input type="number" inputMode="decimal" min="0.01" max="1000000000" step="0.01" placeholder="0.00" value={transactionForm.amount} onChange={(e) => setTransactionForm({ ...transactionForm, amount: e.target.value })} required /></label>
                <label>Date<input type="date" value={transactionForm.date} max={manilaToday()} onChange={(e) => setTransactionForm({ ...transactionForm, date: e.target.value })} required /></label>
                <label>Account<select value={transactionForm.accountId} onChange={(e) => setTransactionForm({ ...transactionForm, accountId: e.target.value })} required><option value="">Select an account</option>{accounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}</select>{accounts.length === 0 && <small className="field-help">Add your first account in Settings.</small>}</label>
                <label>Category <span className="optional">(optional)</span><select value={transactionForm.categoryId} onChange={(e) => setTransactionForm({ ...transactionForm, categoryId: e.target.value })}><option value="">No category</option>{categories.filter((c) => c.kind === transactionForm.type).map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}</select></label>
                <label>Description <span className="optional">(optional)</span><input value={transactionForm.description} onChange={(e) => setTransactionForm({ ...transactionForm, description: e.target.value })} placeholder="What was this for?" maxLength={180} /></label>
                <button className="btn btn-primary btn-full" type="submit" disabled={saving || accounts.length === 0}>{saving ? <LoaderCircle className="spin" size={17} /> : <Plus size={17} />} Save transaction</button>
              </form></section>
              <section className="panel table-panel"><div className="panel-inner"><PanelTitle eyebrow="YOUR RECORDS" title="Transactions" description={`${transactions.length} saved transaction${transactions.length === 1 ? "" : "s"}`} /><TransactionTable transactions={visibleTransactions} search={search} setSearch={setSearch} filter={transactionFilter} setFilter={setTransactionFilter} onDelete={deleteTransaction} /></div></section>
            </div>}
            {page === "Budgets" && <div className="page-columns"><section className="panel form-panel"><PanelTitle eyebrow="SPENDING PLAN" title="Create a budget" description="Set a limit for one of your expense categories." /><form className="stack-form" onSubmit={submitBudget}><label>Expense category<select value={budgetForm.categoryId} onChange={(e) => setBudgetForm({ ...budgetForm, categoryId: e.target.value })} required><option value="">Choose a category</option>{categories.filter((c) => c.kind === "expense").map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}</select>{categories.filter((c) => c.kind === "expense").length === 0 && <small className="field-help">Add an expense category in Settings first.</small>}</label><label>Budget amount (₱)<input type="number" inputMode="decimal" min="0.01" step="0.01" max="1000000000" value={budgetForm.amount} onChange={(e) => setBudgetForm({ ...budgetForm, amount: e.target.value })} placeholder="0.00" required /></label><div className="form-two"><label>Start date<input type="date" value={budgetForm.periodStart} onChange={(e) => setBudgetForm({ ...budgetForm, periodStart: e.target.value })} required /></label><label>End date<input type="date" value={budgetForm.periodEnd} onChange={(e) => setBudgetForm({ ...budgetForm, periodEnd: e.target.value })} required /></label></div><button className="btn btn-primary btn-full" type="submit" disabled={saving || categories.every((c) => c.kind !== "expense")}><Plus size={17} /> Save budget</button></form></section><section className="panel table-panel"><div className="panel-inner"><PanelTitle eyebrow="YOUR PLAN" title="Budget overview" description="Spending is matched to saved expenses in each budget period." />{budgets.length ? <div className="budget-list">{budgets.map((b) => { const pct = b.amount_centavos ? Math.min(100, b.spent_centavos / b.amount_centavos * 100) : 0; return <article className="budget-item" key={b.id}><div className="budget-top"><div><b>{b.category_name}</b><small>{prettyDate(b.period_start)} – {prettyDate(b.period_end)}</small></div><span className={pct >= 100 ? "budget-percent over" : "budget-percent"}>{Math.round(pct)}%</span></div><div className="progress-track"><span className={pct >= 100 ? "progress-over" : ""} style={{ width: pct + "%" }} /></div><div className="budget-bottom"><span>{money(b.spent_centavos)} spent</span><b>{money(b.amount_centavos)} budget</b></div><small className={b.amount_centavos - b.spent_centavos < 0 ? "over-text" : "field-help"}>{b.amount_centavos >= b.spent_centavos ? money(b.amount_centavos - b.spent_centavos) + " remaining" : money(b.spent_centavos - b.amount_centavos) + " over budget"}</small></article>; })}</div> : <EmptyState icon={CreditCard} title="No budgets yet" description="Create a budget to give your spending a plan." />}</div></section></div>}
            {page === "Savings Goals" && <div className="page-columns"><section className="panel form-panel"><PanelTitle eyebrow="YOUR FUTURE" title="Create a savings goal" description="Give your next milestone a name and a target." /><form className="stack-form" onSubmit={submitGoal}><label>Goal name<input value={goalForm.name} onChange={(e) => setGoalForm({ ...goalForm, name: e.target.value })} placeholder="e.g. Emergency fund" maxLength={80} required /></label><label>Target amount (₱)<input type="number" inputMode="decimal" min="0.01" max="1000000000" step="0.01" value={goalForm.targetAmount} onChange={(e) => setGoalForm({ ...goalForm, targetAmount: e.target.value })} placeholder="0.00" required /></label><label>Target date <span className="optional">(optional)</span><input type="date" value={goalForm.targetDate} onChange={(e) => setGoalForm({ ...goalForm, targetDate: e.target.value })} /></label><button className="btn btn-primary btn-full" type="submit" disabled={saving}><Plus size={17} /> Create goal</button></form></section><section className="panel table-panel"><div className="panel-inner"><PanelTitle eyebrow="MILESTONES" title="Your savings goals" description="Update the amount saved as you make progress." />{goals.length ? <div className="goal-list">{goals.map((g) => { const pct = g.target_amount_centavos ? Math.min(100, g.current_amount_centavos / g.target_amount_centavos * 100) : 0; return <article className="goal-card" key={g.id}><div className="goal-card-icon"><Goal size={20} /></div><div className="goal-card-main"><div className="goal-title-row"><b>{g.name}</b><span className={"status-pill " + g.status}>{g.status}</span></div><div className="goal-amounts"><b>{money(g.current_amount_centavos)}</b><span>of {money(g.target_amount_centavos)}</span></div><div className="progress-track"><span style={{ width: pct + "%" }} /></div><div className="goal-edit-row"><label className="saved-input">Saved amount (₱)<input type="number" min="0" max="1000000000" step="0.01" value={goalUpdates[g.id] ?? (g.current_amount_centavos / 100).toFixed(2)} onChange={(e) => setGoalUpdates({ ...goalUpdates, [g.id]: e.target.value })} aria-label={"Saved amount for " + g.name} /></label><button className="btn btn-secondary btn-small" onClick={() => updateGoal(g)} disabled={saving}>Update</button></div>{g.target_date && <small className="field-help">Target date: {prettyDate(g.target_date)}</small>}</div></article>; })}</div> : <EmptyState icon={Goal} title="Your goals start here" description="Create a goal for an emergency fund, a trip, or another future plan." />}</div></section></div>}
            {page === "Reports" && <section className="panel report-panel"><div className="report-hero"><div className="report-icon"><FileSpreadsheet size={25} /></div><div className="report-hero-copy"><span className="overline">YOUR DATA, YOUR REPORT</span><h2>Financial report in Excel</h2><p>Export your saved records into a workbook you can review, filter, or keep for your personal records.</p></div><button className="btn btn-primary" onClick={exportExcel} disabled={saving}><Download size={17} /> Export .xlsx</button></div><div className="report-period-select"><div><b>Report frequency</b><small>Daily and weekly periods use Asia/Manila dates.</small></div><div className="segmented report-segmented">{(["Daily", "Weekly", "Monthly"] as const).map((period) => <button type="button" key={period} className={reportPeriod === period ? "selected income-selected" : ""} onClick={() => setReportPeriod(period)}>{period}</button>)}</div><span className="report-range">{reportWindow(reportPeriod).start} – {reportWindow(reportPeriod).end}</span></div><div className="report-sheet-grid"><div><FileSpreadsheet size={20} /><b>Financial Summary</b><span>Current month income, expenses, savings and remaining balance.</span></div><div><ArrowLeftRight size={20} /><b>Transaction Details</b><span>Your saved transactions, categories, accounts and dates.</span></div><div><CreditCard size={20} /><b>Budget Analysis</b><span>Budget limits, actual spending and remaining amounts.</span></div><div><Goal size={20} /><b>Savings Goals</b><span>Targets, saved amounts, progress and goal status.</span></div></div><div className="report-note"><ShieldCheck size={17} /><span>Reports are generated from your signed-in account's saved records. No sample transactions are included.</span></div></section>}
            {page === "Settings" && <div className="settings-grid"><section className="panel form-panel"><PanelTitle eyebrow="MONEY SOURCES" title="Accounts" description="Add cash, bank, e-wallet, or credit accounts." /><form className="stack-form" onSubmit={createAccount}><label>Account name<input value={accountForm.name} onChange={(e) => setAccountForm({ ...accountForm, name: e.target.value })} placeholder="e.g. Main bank account" maxLength={60} required /></label><label>Account type<select value={accountForm.accountType} onChange={(e) => setAccountForm({ ...accountForm, accountType: e.target.value })}><option value="cash">Cash</option><option value="bank">Bank account</option><option value="e-wallet">E-wallet</option><option value="credit">Credit account</option></select></label><label>Opening balance (₱)<input type="number" min="0" max="1000000000" step="0.01" value={accountForm.openingBalance} onChange={(e) => setAccountForm({ ...accountForm, openingBalance: e.target.value })} required /></label><button className="btn btn-primary btn-full" type="submit" disabled={saving}><Plus size={17} /> Add account</button></form><div className="settings-records">{accounts.length ? accounts.map((a) => <div className="setting-row" key={a.id}><span className="setting-record-icon"><Wallet size={17} /></span><div><b>{a.name}</b><small>{a.account_type}</small></div><strong>{money(a.balance_centavos ?? a.opening_balance_centavos)}</strong></div>) : <EmptyState icon={Wallet} title="No accounts yet" description="Add an account to start saving transactions." />}</div></section><section className="panel form-panel"><PanelTitle eyebrow="ORGANIZE YOUR MONEY" title="Categories" description="Create the categories you want to use." /><form className="stack-form" onSubmit={createCategory}><label>Category name<input value={categoryForm.name} onChange={(e) => setCategoryForm({ ...categoryForm, name: e.target.value })} placeholder="e.g. Groceries" maxLength={60} required /></label><label>Category type<select value={categoryForm.kind} onChange={(e) => setCategoryForm({ ...categoryForm, kind: e.target.value })}><option value="expense">Expense</option><option value="income">Income</option></select></label><button className="btn btn-primary btn-full" type="submit" disabled={saving}><Plus size={17} /> Add category</button></form><div className="category-list">{categories.length ? categories.map((c) => <span className="category-chip" key={c.id}><span className={"category-dot " + c.kind} />{c.name}<small>{c.kind}</small></span>) : <EmptyState icon={Ellipsis} title="No categories yet" description="Add categories to organize transactions." />}</div></section><section className="panel preferences-panel"><PanelTitle eyebrow="PREFERENCES" title="Regional settings" description="These settings are used throughout your finance tracker." /><div className="preference-row"><span className="preference-icon"><Wallet size={18} /></span><div><b>Currency</b><small>Philippine peso</small></div><strong>PHP (₱)</strong></div><div className="preference-row"><span className="preference-icon"><Clock3 size={18} /></span><div><b>Time zone</b><small>Dates and reporting period</small></div><strong>Asia/Manila</strong></div><div className="preference-row"><span className="preference-icon"><ShieldCheck size={18} /></span><div><b>Data privacy</b><small>Only your authenticated account can access these records.</small></div><strong className="private-label">Private</strong></div><button className="btn btn-secondary signout-btn" onClick={signOut} disabled={saving}><LogOut size={17} /> Sign out</button></section></div>}
          </>}
          <footer className="page-footer"><span>© {new Date().getFullYear()} Budgetwise <i>·</i> Built for better money habits</span><span><ShieldCheck size={14} /> Private by design <i>·</i> PHP <i>·</i> Asia/Manila</span></footer>
        </div>
      </main>
    </div>
  );
}

function PanelTitle({ eyebrow, title, description }: { eyebrow: string; title: string; description: string }) {
  return <div className="panel-title"><span className="overline">{eyebrow}</span><h2>{title}</h2><p>{description}</p></div>;
}
function EmptyState({ icon: Icon, title, description }: { icon: typeof Wallet; title: string; description: string }) {
  return <div className="empty-state"><span className="empty-icon"><Icon size={22} /></span><b>{title}</b><p>{description}</p></div>;
}
function TransactionTable({ transactions, search, setSearch, filter, setFilter, onDelete }: { transactions: Transaction[]; search: string; setSearch: (v: string) => void; filter: string; setFilter: (v: string) => void; onDelete: (id: string) => void }) {
  return <>
    <div className="table-toolbar"><label className="search-input"><Search size={16} /><input placeholder="Search transactions" value={search} onChange={(e) => setSearch(e.target.value)} aria-label="Search transactions" /></label><select value={filter} onChange={(e) => setFilter(e.target.value)} aria-label="Filter transaction type"><option value="all">All types</option><option value="income">Income</option><option value="expense">Expenses</option></select></div>
    {transactions.length ? <div className="table-scroll"><table><thead><tr><th>TRANSACTION</th><th>CATEGORY</th><th>DATE</th><th>AMOUNT</th><th><span className="sr-only">Actions</span></th></tr></thead><tbody>{transactions.map((tx) => <tr key={tx.id}><td><div className="tx-title"><span className={"tx-icon " + tx.transaction_type}>{tx.transaction_type === "income" ? <ArrowDownLeft size={16} /> : tx.transaction_type === "expense" ? <ArrowUpRight size={16} /> : <ArrowLeftRight size={16} />}</span><div><b>{tx.description || (tx.transaction_type === "income" ? "Income" : "Expense")}</b><small>{tx.account_name || "Account"} · {tx.transaction_type}</small></div></div></td><td><span className="category-label">{tx.category_name || "Uncategorized"}</span></td><td className="date-col">{prettyDate(tx.transaction_date)}</td><td className={"money-col " + tx.transaction_type}>{tx.transaction_type === "expense" ? "−" : "+"}{money(tx.amount_centavos)}</td><td><button className="icon-btn delete-btn" onClick={() => onDelete(tx.id)} aria-label="Delete transaction" title="Delete transaction"><Trash2 size={15} /></button></td></tr>)}</tbody></table></div> : <EmptyState icon={ArrowLeftRight} title={search || filter !== "all" ? "No matching transactions" : "No transactions yet"} description={search || filter !== "all" ? "Try changing your search or filter." : "Add your first transaction to see your money activity here."} />}
    <div className="table-foot"><span>Showing {transactions.length} saved record{transactions.length === 1 ? "" : "s"} (up to 500)</span></div>
  </>;
}
function DashboardPage({ dashboard, accounts, goals, transactions, setPage }: { dashboard: Dashboard | null; accounts: Account[]; goals: Goal[]; transactions: Transaction[]; setPage: (page: Page) => void }) {
  const totals = dashboard?.totals;
  const activeGoals = goals.filter((g) => g.status !== "paused");
  const goalTotal = activeGoals.reduce((sum, g) => sum + g.target_amount_centavos, 0);
  const goalSaved = activeGoals.reduce((sum, g) => sum + g.current_amount_centavos, 0);
  const expenseCategories = transactions.filter((tx) => tx.transaction_type === "expense").reduce<Record<string, number>>((acc, tx) => { const key = tx.category_name || "Uncategorized"; acc[key] = (acc[key] || 0) + tx.amount_centavos; return acc; }, {});
  const topCategories = Object.entries(expenseCategories).sort((a, b) => b[1] - a[1]).slice(0, 4);
  return <>
    <section className="summary-grid">
      <MetricCard title="Total Income" value={money(totals?.income ?? 0)} icon={ArrowDownLeft} tone="green" caption="This month" />
      <MetricCard title="Total Expenses" value={money(totals?.expenses ?? 0)} icon={ArrowUpRight} tone="orange" caption="This month" />
      <MetricCard title="Savings" value={money(totals?.goalSavings ?? 0)} icon={Goal} tone="violet" caption="Across active goals" />
      <MetricCard title="Remaining Balance" value={money(totals?.remainingBalance ?? 0)} icon={Wallet} tone="blue" caption="Across your accounts" />
    </section>
    <div className="dashboard-grid"><section className="panel overview-panel"><div className="section-head"><div><span className="overline">MONTHLY OVERVIEW</span><h2>Income & expenses</h2><p>{dashboard?.period.start ? prettyDate(dashboard.period.start) + " – " + prettyDate(dashboard.period.end) : "Current month"} · Asia/Manila</p></div><span className="month-chip"><CalendarDays size={14} /> This month</span></div>
      <div className="cashflow-numbers"><div><span className="legend-mark income-mark" />Income <b>{money(totals?.income ?? 0)}</b></div><div><span className="legend-mark expense-mark" />Expenses <b>{money(totals?.expenses ?? 0)}</b></div></div>
      <div className="cashflow-bars"><div className="bar-group"><span>Income</span><div className="bar-bg"><span className="bar-income" style={{ width: (totals?.income ? Math.max(2, totals.income / Math.max(totals.income, totals.expenses, 1) * 100) : 0) + "%" }} /></div><b>{money(totals?.income ?? 0)}</b></div><div className="bar-group"><span>Expenses</span><div className="bar-bg"><span className="bar-expense" style={{ width: (totals?.expenses ? Math.max(2, totals.expenses / Math.max(totals.income, totals.expenses, 1) * 100) : 0) + "%" }} /></div><b>{money(totals?.expenses ?? 0)}</b></div></div>
      <div className={"net-card " + ((totals?.savings ?? 0) >= 0 ? "net-positive" : "net-negative")}><span>{(totals?.savings ?? 0) >= 0 ? <TrendingUp size={18} /> : <ArrowDownRight size={18} />}Net cash flow this month</span><b>{money(totals?.savings ?? 0)}</b></div>
    </section>
    <section className="panel goal-overview-panel"><div className="section-head"><div><span className="overline">YOUR NEXT MILESTONES</span><h2>Savings goals</h2><p>Every little bit counts.</p></div><button className="text-link" onClick={() => setPage("Savings Goals")}>View all <ChevronRight size={15} /></button></div>
      {activeGoals.length ? <><div className="goal-summary"><div className="goal-ring" style={{ background: "conic-gradient(#69a681 " + Math.min(100, goalTotal ? goalSaved / goalTotal * 100 : 0) + "%, #edf2ec 0)" }}><span><b>{goalTotal ? Math.round(goalSaved / goalTotal * 100) : 0}%</b><small>saved</small></span></div><div><b>{money(goalSaved)}</b><small>of {money(goalTotal)} target</small><div className="goal-summary-note">{activeGoals.length} active goal{activeGoals.length === 1 ? "" : "s"}</div></div></div><div className="mini-goal-list">{activeGoals.slice(0, 3).map((g) => <div className="mini-goal" key={g.id}><div><span>{g.name}</span><b>{Math.round(g.current_amount_centavos / Math.max(1, g.target_amount_centavos) * 100)}%</b></div><div className="progress-track"><span style={{ width: Math.min(100, g.current_amount_centavos / Math.max(1, g.target_amount_centavos) * 100) + "%" }} /></div></div>)}</div></> : <EmptyState icon={Goal} title="Start a savings goal" description="Set a target and keep track of your progress." />}
      <button className="outline-full" onClick={() => setPage("Savings Goals")}><Plus size={16} /> Manage savings goals</button>
    </section></div>
    <div className="dashboard-grid lower-grid"><section className="panel recent-panel"><div className="section-head"><div><span className="overline">LATEST ACTIVITY</span><h2>Recent transactions</h2><p>Your latest saved money movements.</p></div><button className="text-link" onClick={() => setPage("Transactions")}>View all <ChevronRight size={15} /></button></div>
      {transactions.length ? <div className="recent-list">{transactions.slice(0, 5).map((tx) => <div className="recent-row" key={tx.id}><span className={"tx-icon " + tx.transaction_type}>{tx.transaction_type === "income" ? <ArrowDownLeft size={16} /> : <ArrowUpRight size={16} />}</span><div className="recent-desc"><b>{tx.description || (tx.transaction_type === "income" ? "Income" : "Expense")}</b><small>{tx.category_name || "Uncategorized"} · {prettyDate(tx.transaction_date)}</small></div><b className={"recent-amount " + tx.transaction_type}>{tx.transaction_type === "expense" ? "−" : "+"}{money(tx.amount_centavos)}</b></div>)}</div> : <EmptyState icon={ArrowLeftRight} title="Nothing to show yet" description="Your saved transactions will appear here." />}
    </section><section className="panel category-overview"><div className="section-head"><div><span className="overline">SPENDING BREAKDOWN</span><h2>Expense categories</h2><p>From the latest 500 saved transactions.</p></div></div>
      {topCategories.length ? <div className="category-breakdown">{topCategories.map(([name, amount], i) => { const total = topCategories.reduce((s, x) => s + x[1], 0); return <div className="breakdown-row" key={name}><div className="breakdown-label"><span className={"breakdown-dot dot-" + i} /><span>{name}</span><b>{money(amount)}</b></div><div className="progress-track"><span className={"breakdown-bar bar-" + i} style={{ width: (amount / Math.max(total, 1) * 100) + "%" }} /></div></div>; })}</div> : <EmptyState icon={BarChart3} title="No spending breakdown" description="Categorize expenses to see where your money goes." />}
    </section></div>
    {accounts.length === 0 && <div className="getting-started"><div className="getting-icon"><Wallet size={20} /></div><div><b>First, add a money account</b><p>Set up a cash, bank, or e-wallet account before recording transactions.</p></div><button className="btn btn-primary btn-small" onClick={() => setPage("Settings")}>Set up account <ChevronRight size={15} /></button></div>}
  </>;
}
function MetricCard({ title, value, icon: Icon, tone, caption }: { title: string; value: string; icon: typeof Wallet; tone: string; caption: string }) {
  return <article className="metric-card"><div className="metric-top"><span>{title}</span><span className={"metric-icon " + tone}><Icon size={18} /></span></div><div className="metric-value">{value}</div><div className="metric-caption"><span className={"caption-dot " + tone} />{caption}</div></article>;
}
