import { useCallback, useEffect, useMemo, useState } from "react";
import type { FormEvent } from "react";
import { getPeriodRange, shiftPeriod, summarizeTransactions } from "./lib/finance";
import type { PeriodKind } from "./lib/finance";
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
type Category = { id: string; name: string; kind: "income" | "expense"; budget_group: "needs" | "wants" | "savings" | "income"; color?: string | null };
type Transaction = { id: string; account_id: string; category_id: string | null; goal_id?: string | null; transaction_type: "income" | "expense" | "savings_contribution" | "transfer"; amount_centavos: number; transaction_date: string; description: string; payment_method: "cash" | "bank_transfer" | "e_wallet" | "credit_card" | "other"; notes: string | null; account_name?: string; category_name?: string | null };
type Budget = { id: string; category_id: string; category_name: string; period_start: string; period_end: string; amount_centavos: number; spent_centavos: number };
type Goal = { id: string; name: string; description?: string | null; target_amount_centavos: number; target_date: string | null; current_amount_centavos: number; status: "active" | "completed" | "paused" };
type Dashboard = { period: { start: string; end: string; timezone: string }; totals: { income: number; expenses: number; savings: number; savingsContributions: number; remainingBalance: number; goalSavings: number }; recentTransactions: Transaction[] };
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
function manilaToday() {
  const p = new Intl.DateTimeFormat("en-CA", { timeZone: "Asia/Manila", year: "numeric", month: "2-digit", day: "2-digit" }).formatToParts(new Date());
  const get = (type: string) => p.find((x) => x.type === type)?.value ?? "";
  return `${get("year")}-${get("month")}-${get("day")}`;
}
function manilaMonthStart() { return manilaToday().slice(0, 8) + "01"; }
function reportWindow(period: "Daily" | "Weekly" | "Monthly") { return getPeriodRange(manilaToday(), period); }

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
  const [transactionForm, setTransactionForm] = useState({ type: "expense", amount: "", date: manilaToday(), accountId: "", categoryId: "", description: "", paymentMethod: "cash", notes: "", goalId: "" });
  const [editingTransactionId, setEditingTransactionId] = useState<string | null>(null);
  const [viewTransaction, setViewTransaction] = useState<Transaction | null>(null);
  const [transactionPeriod, setTransactionPeriod] = useState<PeriodKind>("Monthly");
  const [periodAnchor, setPeriodAnchor] = useState(manilaToday());
  const [customStart, setCustomStart] = useState(manilaMonthStart());
  const [customEnd, setCustomEnd] = useState(manilaToday());
  const [filterCategory, setFilterCategory] = useState("all");
  const [filterPaymentMethod, setFilterPaymentMethod] = useState("all");
  const [budgetForm, setBudgetForm] = useState({ categoryId: "", amount: "", periodStart: manilaMonthStart(), periodEnd: manilaToday() });
  const [goalForm, setGoalForm] = useState({ name: "", description: "", targetAmount: "", targetDate: "" });
  const [goalUpdates, setGoalUpdates] = useState<Record<string, string>>({});
  const [accountForm, setAccountForm] = useState({ name: "", accountType: "bank", openingBalance: "0" });
  const [categoryForm, setCategoryForm] = useState({ name: "", kind: "expense", budgetGroup: "needs" });
  const [editingCategoryId, setEditingCategoryId] = useState<string | null>(null);
  const [editingBudgetId, setEditingBudgetId] = useState<string | null>(null);
  const [budgetPeriodKind, setBudgetPeriodKind] = useState<"Daily" | "Weekly" | "Monthly">("Monthly");
  const [budgetAnchor, setBudgetAnchor] = useState(manilaToday());

  const refreshAll = useCallback(async () => {
    setLoading(true);
    setError("");
    try {
      const me = await api<{ user: User }>("/api/me");
      setUser(me.user);
      const [dash, tx, acc, cats, buds, savings] = await Promise.all([
        api<Dashboard>("/api/dashboard"),
        api<ApiResult<Transaction[]>>("/api/transactions"),
        api<ApiResult<Account[]>>("/api/accounts"),
        api<ApiResult<Category[]>>("/api/categories"),
        api<ApiResult<Budget[]>>("/api/budgets"),
        api<ApiResult<Goal[]>>("/api/savings-goals"),
      ]);
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
    if (!transactionForm.accountId) { setError("Create or select an account in Settings."); return; }
    if (!transactionForm.date || transactionForm.date > manilaToday()) { setError("Choose a valid transaction date no later than today (Asia/Manila)."); return; }
    if (!transactionForm.description.trim()) { setError("Description is required."); return; }
    if (!transactionForm.categoryId) { setError("Choose a category."); return; }
    if (!transactionForm.paymentMethod) { setError("Choose a payment method."); return; }
    setSaving(true);
    try {
      const payload = { type: transactionForm.type, amount, date: transactionForm.date, accountId: transactionForm.accountId, categoryId: transactionForm.categoryId, description: transactionForm.description.trim(), paymentMethod: transactionForm.paymentMethod, notes: transactionForm.notes.trim(), goalId: transactionForm.type === "savings_contribution" ? transactionForm.goalId || null : null };
      const wasEditing = Boolean(editingTransactionId);
      await api(editingTransactionId ? "/api/transactions/" + editingTransactionId : "/api/transactions", { method: editingTransactionId ? "PUT" : "POST", body: JSON.stringify(payload) });
      setTransactionForm((old) => ({ ...old, amount: "", description: "", categoryId: "", notes: "" }));
      setEditingTransactionId(null);
      await refreshAll(); setPage("Transactions"); flash(wasEditing ? "Transaction updated." : "Transaction saved to your private account.");
    } catch (e) { setError(e instanceof Error ? e.message : "Could not save transaction."); }
    finally { setSaving(false); }
  };

  const startEditTransaction = (tx: Transaction) => {
    setEditingTransactionId(tx.id);
    setTransactionForm({ type: tx.transaction_type === "transfer" ? "expense" : tx.transaction_type, amount: (tx.amount_centavos / 100).toFixed(2), date: tx.transaction_date, accountId: tx.account_id, categoryId: tx.category_id || "", description: tx.description || "", paymentMethod: tx.payment_method || "cash", notes: tx.notes || "", goalId: tx.goal_id || "" });
    document.getElementById("transaction-entry")?.scrollIntoView({ behavior: "smooth", block: "start" });
  };

  const deleteTransaction = async (transactionId: string) => {
    if (!window.confirm("Delete this transaction? This cannot be undone.")) return;
    setSaving(true); setError(""); setSuccess("");
    try {
      await api("/api/transactions/" + transactionId, { method: "DELETE" });
      if (editingTransactionId === transactionId) setEditingTransactionId(null);
      await refreshAll();
      flash("Transaction deleted.");
    } catch (e) { setError(e instanceof Error ? e.message : "Could not delete transaction."); }
    finally { setSaving(false); }
  };

  const submitBudget = async (event: FormEvent) => {
    event.preventDefault(); setError(""); setSuccess("");
    const amount = Number(budgetForm.amount);
    if (!budgetForm.categoryId) { setError("Choose a Needs or Wants category."); return; }
    if (!Number.isFinite(amount) || amount <= 0 || amount > 1_000_000_000) { setError("Enter a budget greater than ₱0.00."); return; }
    if (!budgetForm.periodStart || !budgetForm.periodEnd || budgetForm.periodEnd < budgetForm.periodStart) { setError("The end date must be on or after the start date."); return; }
    setSaving(true);
    try {
      const payload = { categoryId: budgetForm.categoryId, amount, periodStart: budgetForm.periodStart, periodEnd: budgetForm.periodEnd };
      await api(editingBudgetId ? "/api/budgets/" + editingBudgetId : "/api/budgets", { method: editingBudgetId ? "PUT" : "POST", body: JSON.stringify(payload) });
      setBudgetForm((old) => ({ ...old, amount: "" })); setEditingBudgetId(null);
      await refreshAll(); flash(editingBudgetId ? "Budget updated." : "Budget saved.");
    } catch (e) { setError(e instanceof Error ? e.message : "Could not save budget."); }
    finally { setSaving(false); }
  };
  const editBudget = (budget: Budget) => {
    setEditingBudgetId(budget.id);
    setBudgetForm({ categoryId: budget.category_id, amount: (budget.amount_centavos / 100).toFixed(2), periodStart: budget.period_start, periodEnd: budget.period_end });
    document.getElementById("budget-entry")?.scrollIntoView({ behavior: "smooth", block: "start" });
  };
  const deleteBudget = async (budget: Budget) => {
    if (!window.confirm("Delete this budget for " + budget.category_name + "?")) return;
    setSaving(true); setError(""); setSuccess("");
    try { await api("/api/budgets/" + budget.id, { method: "DELETE" }); await refreshAll(); flash("Budget deleted."); }
    catch (e) { setError(e instanceof Error ? e.message : "Could not delete budget."); }
    finally { setSaving(false); }
  };

  const submitGoal = async (event: FormEvent) => {
    event.preventDefault(); setError(""); setSuccess("");
    const targetAmount = Number(goalForm.targetAmount);
    if (!goalForm.name.trim()) { setError("Enter a name for your savings goal."); return; }
    if (!Number.isFinite(targetAmount) || targetAmount <= 0) { setError("Enter a target amount greater than ₱0.00."); return; }
    setSaving(true);
    try { await api("/api/savings-goals", { method: "POST", body: JSON.stringify({ name: goalForm.name.trim(), description: goalForm.description.trim(), targetAmount, targetDate: goalForm.targetDate || null }) }); setGoalForm({ name: "", description: "", targetAmount: "", targetDate: "" }); await refreshAll(); flash("Savings goal created."); }
    catch (e) { setError(e instanceof Error ? e.message : "Could not create savings goal."); }
    finally { setSaving(false); }
  };

  const updateGoal = async (goal: Goal) => {
    const raw = goalUpdates[goal.id] ?? (goal.current_amount_centavos / 100).toFixed(2);
    const totalSaved = Number(raw);
    const linkedContributions = transactions.filter((tx) => tx.transaction_type === "savings_contribution" && tx.goal_id === goal.id).reduce((sum,tx) => sum + tx.amount_centavos,0);
    if (!Number.isFinite(totalSaved) || totalSaved < linkedContributions / 100 || totalSaved > 1_000_000_000) { setError("Total saved must be at least the amount already tracked in linked contributions."); return; }
    const baseSaved = totalSaved - linkedContributions / 100;
    setSaving(true); setError(""); setSuccess("");
    try { await api("/api/savings-goals/" + goal.id, { method: "PATCH", body: JSON.stringify({ currentAmount: baseSaved, status: totalSaved >= goal.target_amount_centavos / 100 ? "completed" : "active" }) }); await refreshAll(); flash("Savings progress updated."); }
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
    try {
      const payload = { name: categoryForm.name.trim(), kind: categoryForm.kind, budgetGroup: categoryForm.kind === "income" ? "income" : categoryForm.budgetGroup };
      await api(editingCategoryId ? "/api/categories/" + editingCategoryId : "/api/categories", { method: editingCategoryId ? "PATCH" : "POST", body: JSON.stringify(payload) });
      setCategoryForm({ name: "", kind: "expense", budgetGroup: "needs" }); setEditingCategoryId(null);
      await refreshAll(); flash(editingCategoryId ? "Category updated." : "Category created.");
    } catch (e) { setError(e instanceof Error ? e.message : "Could not save category."); }
    finally { setSaving(false); }
  };

  const editCategory = (category: Category) => {
    setEditingCategoryId(category.id);
    setCategoryForm({ name: category.name, kind: category.kind, budgetGroup: category.budget_group });
    document.getElementById("category-entry")?.scrollIntoView({ behavior: "smooth", block: "start" });
  };
  const deleteCategory = async (category: Category) => {
    if (!window.confirm("Delete category “" + category.name + "”? Categories used by transactions or budgets cannot be deleted.")) return;
    setSaving(true); setError(""); setSuccess("");
    try { await api("/api/categories/" + category.id, { method: "DELETE" }); await refreshAll(); flash("Category deleted."); }
    catch (e) { setError(e instanceof Error ? e.message : "Could not delete category."); }
    finally { setSaving(false); }
  };


  const activeTransactionRange = useMemo(() => transactionPeriod === "Custom"
    ? { start: customStart, end: customEnd }
    : getPeriodRange(periodAnchor || manilaToday(), transactionPeriod), [transactionPeriod, periodAnchor, customStart, customEnd]);
  const transactionSummary = useMemo(() => summarizeTransactions(transactions, activeTransactionRange), [transactions, activeTransactionRange]);
  const visibleTransactions = useMemo(() => transactions.filter((tx) => {
    const text = [tx.description, tx.category_name, tx.account_name, tx.transaction_date, tx.transaction_type, tx.payment_method, tx.notes].join(" ").toLowerCase();
    return text.includes(search.toLowerCase()) && tx.transaction_date >= activeTransactionRange.start && tx.transaction_date <= activeTransactionRange.end
      && (transactionFilter === "all" || tx.transaction_type === transactionFilter)
      && (filterCategory === "all" || tx.category_id === filterCategory)
      && (filterPaymentMethod === "all" || tx.payment_method === filterPaymentMethod);
  }), [transactions, search, transactionFilter, filterCategory, filterPaymentMethod, activeTransactionRange]);

  const exportExcel = async () => {
    setSaving(true); setError(""); setSuccess("");
    try {
      const range = getPeriodRange(manilaToday(), reportPeriod);
      const reportTransactions = transactions.filter((tx) => tx.transaction_date >= range.start && tx.transaction_date <= range.end);
      const reportTotals = summarizeTransactions(reportTransactions, range);
      const reportIncome = reportTotals.income;
      const reportExpenses = reportTotals.expenses;
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
        { metric: "Savings contributions", amount: reportTotals.savingsContributions / 100 },
        { metric: "Remaining balance", amount: reportTotals.remainingBalance / 100 },
        { metric: "Savings across goals", amount: (dashboard?.totals.goalSavings ?? 0) / 100 },
      ]);
      const txSheet = workbook.addWorksheet("Transaction Details");
      txSheet.columns = [{ header: "Date", key: "date", width: 16 }, { header: "Description", key: "description", width: 30 }, { header: "Category", key: "category", width: 22 }, { header: "Account", key: "account", width: 22 }, { header: "Type", key: "type", width: 22 }, { header: "Payment method", key: "payment", width: 20 }, { header: "Amount (PHP)", key: "amount", width: 18 }, { header: "Notes", key: "notes", width: 36 }];
      txSheet.addRows(reportTransactions.map((tx) => ({ date: tx.transaction_date, description: tx.description || "", category: tx.category_name || "", account: tx.account_name || "", type: tx.transaction_type, payment: tx.payment_method, amount: (tx.transaction_type === "expense" || tx.transaction_type === "savings_contribution" ? -tx.amount_centavos : tx.amount_centavos) / 100, notes: tx.notes || "" })));
      txSheet.autoFilter = { from: "A1", to: "H1" };
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
      const blob = new Blob([new Uint8Array(buffer)], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
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
            {page === "Dashboard" && <DashboardPage accounts={accounts} categories={categories} goals={goals} transactions={transactions} setPage={setPage} />}
            {page === "Transactions" && <div className="transaction-workspace">
              <section className="panel period-panel">
                <div className="section-head"><div><span className="overline">PERIOD SUMMARY</span><h2>Your money movement</h2><p>All dates are calendar dates in Asia/Manila.</p></div></div>
                <div className="period-toolbar">
                  <div className="segmented period-segmented">{(["Daily", "Weekly", "Monthly", "Custom"] as PeriodKind[]).map((period) => <button key={period} type="button" className={transactionPeriod === period ? "selected income-selected" : ""} onClick={() => setTransactionPeriod(period)}>{period}</button>)}</div>
                  {transactionPeriod === "Custom" ? <div className="period-date-controls"><label>From<input type="date" value={customStart} onChange={(e) => setCustomStart(e.target.value)} /></label><label>To<input type="date" value={customEnd} min={customStart} onChange={(e) => setCustomEnd(e.target.value)} /></label></div> : <div className="period-date-controls"><button className="btn btn-secondary btn-small" type="button" onClick={() => setPeriodAnchor(shiftPeriod(periodAnchor || manilaToday(), transactionPeriod, -1))}><ChevronRight size={14} style={{ transform: "rotate(180deg)" }} /> Previous</button><label>Anchor date<input type="date" value={periodAnchor} onChange={(e) => setPeriodAnchor(e.target.value)} /></label><button className="btn btn-secondary btn-small" type="button" onClick={() => setPeriodAnchor(shiftPeriod(periodAnchor || manilaToday(), transactionPeriod, 1))}>Next <ChevronRight size={14} /></button></div>}
                </div>
                <p className="active-range-label">Showing {prettyDate(activeTransactionRange.start)} – {prettyDate(activeTransactionRange.end)}</p>
                <div className="summary-grid transaction-summary-grid">
                  <MetricCard title="Total Income" value={money(transactionSummary.income)} icon={ArrowDownLeft} tone="green" caption="Income transactions" />
                  <MetricCard title="Total Expenses" value={money(transactionSummary.expenses)} icon={ArrowUpRight} tone="orange" caption="Expense transactions" />
                  <MetricCard title="Savings Contributions" value={money(transactionSummary.savingsContributions)} icon={Goal} tone="violet" caption="Allocated, not expenses" />
                  <MetricCard title="Remaining Balance" value={money(transactionSummary.remainingBalance)} icon={Wallet} tone="blue" caption="Income − expenses − savings" />
                </div>
                {customStart > customEnd && transactionPeriod === "Custom" && <div className="alert alert-error">The end date must be on or after the start date.</div>}
              </section>
              <div className="page-columns">
                <section className="panel form-panel" id="transaction-entry"><PanelTitle eyebrow={editingTransactionId ? "EDIT RECORD" : "NEW RECORD"} title={editingTransactionId ? "Edit transaction" : "Add a transaction"} description="Required fields are validated before saving." />
                  <form onSubmit={submitTransaction} className="stack-form">
                    <label>Transaction type<div className="segmented transaction-type-segmented">
                      <button type="button" className={transactionForm.type === "expense" ? "selected expense-selected" : ""} onClick={() => setTransactionForm({ ...transactionForm, type: "expense", categoryId: "" })}><ArrowUpRight size={15} /> Expense</button>
                      <button type="button" className={transactionForm.type === "income" ? "selected income-selected" : ""} onClick={() => setTransactionForm({ ...transactionForm, type: "income", categoryId: "" })}><ArrowDownLeft size={15} /> Income</button>
                      <button type="button" className={transactionForm.type === "savings_contribution" ? "selected income-selected" : ""} onClick={() => setTransactionForm({ ...transactionForm, type: "savings_contribution", categoryId: "" })}><Goal size={15} /> Savings</button>
                    </div></label>
                    <label>Transaction date<input type="date" value={transactionForm.date} max={manilaToday()} onChange={(e) => setTransactionForm({ ...transactionForm, date: e.target.value })} required /></label>
                    <label>Description<input value={transactionForm.description} onChange={(e) => setTransactionForm({ ...transactionForm, description: e.target.value })} placeholder="e.g. Monthly salary" maxLength={180} required /></label>
                    <label>Category<select value={transactionForm.categoryId} onChange={(e) => setTransactionForm({ ...transactionForm, categoryId: e.target.value })} required><option value="">Select a category</option>{categories.filter((cat) => transactionForm.type === "income" ? cat.kind === "income" : cat.kind === "expense" && (transactionForm.type === "savings_contribution" ? cat.budget_group === "savings" : cat.budget_group !== "savings")).map((cat) => <option key={cat.id} value={cat.id}>{cat.name}</option>)}</select>{categories.filter((cat) => transactionForm.type === "income" ? cat.kind === "income" : cat.kind === "expense" && (transactionForm.type === "savings_contribution" ? cat.budget_group === "savings" : cat.budget_group !== "savings")).length === 0 && <small className="field-help">Create a matching category in Settings first.</small>}</label>
                    {transactionForm.type === "savings_contribution" && <label>Link to savings goal <span className="optional">(optional)</span><select value={transactionForm.goalId} onChange={(e) => setTransactionForm({ ...transactionForm, goalId: e.target.value })}><option value="">General savings</option>{goals.map((goal) => <option key={goal.id} value={goal.id}>{goal.name}</option>)}</select></label>}
                    <label>Amount (₱)<input type="number" inputMode="decimal" min="0.01" max="1000000000" step="0.01" placeholder="0.00" value={transactionForm.amount} onChange={(e) => setTransactionForm({ ...transactionForm, amount: e.target.value })} required /></label>
                    <label>Payment method<select value={transactionForm.paymentMethod} onChange={(e) => setTransactionForm({ ...transactionForm, paymentMethod: e.target.value })} required><option value="cash">Cash</option><option value="bank_transfer">Bank transfer</option><option value="e_wallet">E-wallet</option><option value="credit_card">Credit card</option><option value="other">Other</option></select></label>
                    <label>Account<select value={transactionForm.accountId} onChange={(e) => setTransactionForm({ ...transactionForm, accountId: e.target.value })} required><option value="">Select an account</option>{accounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}</select>{accounts.length === 0 && <small className="field-help">Add your first account in Settings.</small>}</label>
                    <label>Notes <span className="optional">(optional)</span><textarea value={transactionForm.notes} onChange={(e) => setTransactionForm({ ...transactionForm, notes: e.target.value })} placeholder="Additional details" maxLength={1000} rows={3} /></label>
                    <button className="btn btn-primary btn-full" type="submit" disabled={saving || accounts.length === 0 || !categories.some((cat) => cat.id === transactionForm.categoryId && (transactionForm.type === "income" ? cat.kind === "income" : cat.kind === "expense"))}>{saving ? <LoaderCircle className="spin" size={17} /> : editingTransactionId ? <Check size={17} /> : <Plus size={17} />}{editingTransactionId ? "Update transaction" : "Save transaction"}</button>
                    {editingTransactionId && <button className="btn btn-secondary btn-full" type="button" onClick={() => { setEditingTransactionId(null); setTransactionForm((old) => ({ ...old, amount: "", description: "", categoryId: "", notes: "" })); }}>Cancel edit</button>}
                  </form>
                </section>
                <section className="panel table-panel"><div className="panel-inner"><PanelTitle eyebrow="YOUR RECORDS" title="Transactions" description={`${visibleTransactions.length} matching record${visibleTransactions.length === 1 ? "" : "s"}`} />
                  <TransactionTable transactions={visibleTransactions} search={search} setSearch={setSearch} filter={transactionFilter} setFilter={setTransactionFilter} categories={categories} filterCategory={filterCategory} setFilterCategory={setFilterCategory} filterPaymentMethod={filterPaymentMethod} setFilterPaymentMethod={setFilterPaymentMethod} onView={setViewTransaction} onEdit={startEditTransaction} onDelete={deleteTransaction} />
                </div></section>
              </div>
            </div>}
            {page === "Budgets" && <div className="budget-workspace">
              <section className="panel period-panel">
                <div className="section-head"><div><span className="overline">SPENDING LIMITS</span><h2>Budget period</h2><p>Only expense transactions count toward spending. Savings contributions are excluded.</p></div></div>
                <div className="period-toolbar">
                  <div className="segmented period-segmented">{(["Daily","Weekly","Monthly"] as const).map((period) => <button type="button" key={period} className={budgetPeriodKind === period ? "selected income-selected" : ""} onClick={() => { setBudgetPeriodKind(period); const range = getPeriodRange(budgetAnchor || manilaToday(), period); setBudgetForm((old) => ({ ...old, periodStart: range.start, periodEnd: range.end })); }}>{period}</button>)}</div>
                  <div className="period-date-controls"><button className="btn btn-secondary btn-small" type="button" onClick={() => { const anchor = shiftPeriod(budgetAnchor || manilaToday(), budgetPeriodKind, -1); setBudgetAnchor(anchor); const range = getPeriodRange(anchor, budgetPeriodKind); setBudgetForm((old) => ({ ...old, periodStart: range.start, periodEnd: range.end })); }}><ChevronRight size={14} style={{ transform: "rotate(180deg)" }} /> Previous</button><label>Anchor date<input type="date" value={budgetAnchor} onChange={(e) => { const anchor = e.target.value; setBudgetAnchor(anchor); if (anchor) { const range = getPeriodRange(anchor, budgetPeriodKind); setBudgetForm((old) => ({ ...old, periodStart: range.start, periodEnd: range.end })); } }} /></label><button className="btn btn-secondary btn-small" type="button" onClick={() => { const anchor = shiftPeriod(budgetAnchor || manilaToday(), budgetPeriodKind, 1); setBudgetAnchor(anchor); const range = getPeriodRange(anchor, budgetPeriodKind); setBudgetForm((old) => ({ ...old, periodStart: range.start, periodEnd: range.end })); }}>Next <ChevronRight size={14} /></button></div>
                </div>
                <p className="active-range-label">New budget period: {prettyDate(budgetForm.periodStart)} – {prettyDate(budgetForm.periodEnd)}</p>
              </section>
              <div className="page-columns">
                <section className="panel form-panel" id="budget-entry"><PanelTitle eyebrow={editingBudgetId ? "EDIT LIMIT" : "SPENDING PLAN"} title={editingBudgetId ? "Edit budget" : "Create a budget"} description="Set a daily, weekly, or monthly limit for a Needs or Wants category." />
                  <form className="stack-form" onSubmit={submitBudget}>
                    <label>Budget category<select value={budgetForm.categoryId} onChange={(e) => setBudgetForm({ ...budgetForm, categoryId: e.target.value })} required><option value="">Choose a category</option>{categories.filter((cat) => cat.kind === "expense" && cat.budget_group !== "savings").map((cat) => <option key={cat.id} value={cat.id}>{cat.name} · {cat.budget_group}</option>)}</select></label>
                    <label>Budget amount (₱)<input type="number" inputMode="decimal" min="0.01" max="1000000000" step="0.01" value={budgetForm.amount} onChange={(e) => setBudgetForm({ ...budgetForm, amount: e.target.value })} placeholder="0.00" required /></label>
                    <div className="form-two"><label>Start date<input type="date" value={budgetForm.periodStart} onChange={(e) => setBudgetForm({ ...budgetForm, periodStart: e.target.value })} required /></label><label>End date<input type="date" value={budgetForm.periodEnd} min={budgetForm.periodStart} onChange={(e) => setBudgetForm({ ...budgetForm, periodEnd: e.target.value })} required /></label></div>
                    <button className="btn btn-primary btn-full" type="submit" disabled={saving || !categories.some((cat) => cat.id === budgetForm.categoryId && cat.kind === "expense" && cat.budget_group !== "savings")}>{saving ? <LoaderCircle className="spin" size={17} /> : editingBudgetId ? <Check size={17} /> : <Plus size={17} />}{editingBudgetId ? "Update budget" : "Save budget"}</button>
                    {editingBudgetId && <button className="btn btn-secondary btn-full" type="button" onClick={() => { setEditingBudgetId(null); setBudgetForm((old) => ({ ...old, amount: "" })); }}>Cancel edit</button>}
                  </form>
                </section>
                <section className="panel table-panel"><div className="panel-inner"><PanelTitle eyebrow="YOUR PLAN" title="Budget overview" description="Actual spending updates from saved expense transactions." />
                  {budgets.length ? <div className="budget-list">{budgets.map((b) => { const pct = b.amount_centavos ? b.spent_centavos / b.amount_centavos * 100 : 0; const remaining = b.amount_centavos - b.spent_centavos; return <article className="budget-item" key={b.id}>
                    <div className="budget-top"><div><b>{b.category_name}</b><small>{prettyDate(b.period_start)} – {prettyDate(b.period_end)}</small></div><span className={pct >= 100 ? "budget-percent over" : "budget-percent"}>{Math.round(pct)}% consumed</span></div>
                    <div className="progress-track"><span className={pct >= 100 ? "progress-over" : ""} style={{ width: Math.min(100,Math.max(0,pct)) + "%" }} /></div>
                    <div className="budget-bottom"><span>Actual: {money(b.spent_centavos)}</span><b>Limit: {money(b.amount_centavos)}</b></div>
                    <small className={remaining < 0 ? "over-text" : "field-help"}>{remaining < 0 ? "Limit exceeded by " + money(Math.abs(remaining)) : money(remaining) + " remaining"}</small>
                    {pct >= 100 && <div className="budget-alert"><CircleHelp size={15} /> {pct === 100 ? "Spending limit reached." : "Spending limit exceeded."}</div>}
                    <div className="budget-actions"><button className="btn btn-secondary btn-small" type="button" onClick={() => editBudget(b)}><Settings size={14} /> Edit</button><button className="btn btn-secondary btn-small" type="button" onClick={() => deleteBudget(b)}><Trash2 size={14} /> Delete</button></div>
                  </article>; })}</div> : <EmptyState icon={CreditCard} title="No budgets yet" description="Create a budget to give your spending a plan." />}
                </div></section>
              </div>
            </div>}
            {page === "Savings Goals" && <div className="page-columns">
              <section className="panel form-panel"><PanelTitle eyebrow="YOUR FUTURE" title="Create a savings goal" description="Give your next milestone a name, target, and plan." />
                <form className="stack-form" onSubmit={submitGoal}>
                  <label>Goal name<input value={goalForm.name} onChange={(e) => setGoalForm({ ...goalForm, name: e.target.value })} placeholder="e.g. Emergency fund" maxLength={80} required /></label>
                  <label>Target amount (₱)<input type="number" inputMode="decimal" min="0.01" max="1000000000" step="0.01" value={goalForm.targetAmount} onChange={(e) => setGoalForm({ ...goalForm, targetAmount: e.target.value })} placeholder="0.00" required /></label>
                  <label>Target date <span className="optional">(optional)</span><input type="date" value={goalForm.targetDate} onChange={(e) => setGoalForm({ ...goalForm, targetDate: e.target.value })} /></label>
                  <label>Description <span className="optional">(optional)</span><textarea value={goalForm.description} onChange={(e) => setGoalForm({ ...goalForm, description: e.target.value })} placeholder="What are you saving for?" maxLength={500} rows={3} /></label>
                  <button className="btn btn-primary btn-full" type="submit" disabled={saving}><Plus size={17} /> Create goal</button>
                </form>
              </section>
              <section className="panel table-panel"><div className="panel-inner"><PanelTitle eyebrow="MILESTONES" title="Your savings goals" description="Goal-linked savings contributions automatically increase saved totals." />
                {goals.length ? <div className="goal-list">{goals.map((g) => {
                  const pct = g.target_amount_centavos ? Math.min(100, g.current_amount_centavos / g.target_amount_centavos * 100) : 0;
                  const remaining = Math.max(0,g.target_amount_centavos-g.current_amount_centavos);
                  const linkedContributions = transactions.filter((tx) => tx.transaction_type === "savings_contribution" && tx.goal_id === g.id).reduce((sum,tx) => sum + tx.amount_centavos,0);
                  return <article className="goal-card" key={g.id}><div className="goal-card-icon"><Goal size={20} /></div><div className="goal-card-main">
                    <div className="goal-title-row"><b>{g.name}</b><span className={"status-pill " + g.status}>{g.status}</span></div>
                    {g.description && <p className="goal-description">{g.description}</p>}
                    <div className="goal-amounts"><b>{money(g.current_amount_centavos)}</b><span>of {money(g.target_amount_centavos)}</span></div>
                    <div className="progress-track goal-progress"><span style={{ width: pct + "%" }} /></div>
                    <div className="goal-metrics"><span>{Math.round(pct)}% complete</span><b>{money(remaining)} to go</b></div>
                    <div className="goal-edit-row"><label className="saved-input">Total saved amount (₱)<input type="number" min={(linkedContributions/100).toFixed(2)} max="1000000000" step="0.01" value={goalUpdates[g.id] ?? (g.current_amount_centavos / 100).toFixed(2)} onChange={(e) => setGoalUpdates({ ...goalUpdates, [g.id]: e.target.value })} aria-label={"Saved amount for " + g.name} /></label><button className="btn btn-secondary btn-small" onClick={() => updateGoal(g)} disabled={saving}>Update</button></div>
                    {linkedContributions > 0 && <small className="field-help">{money(linkedContributions)} tracked through linked contribution transactions.</small>}
                    {g.target_date && <small className="field-help">Target date: {prettyDate(g.target_date)}</small>}
                  </div></article>;
                })}</div> : <EmptyState icon={Goal} title="Your goals start here" description="Create a goal for an emergency fund, a trip, or another future plan." />}
              </div></section>
            </div>}
            {page === "Reports" && <section className="panel report-panel"><div className="report-hero"><div className="report-icon"><FileSpreadsheet size={25} /></div><div className="report-hero-copy"><span className="overline">YOUR DATA, YOUR REPORT</span><h2>Financial report in Excel</h2><p>Export your saved records into a workbook you can review, filter, or keep for your personal records.</p></div><button className="btn btn-primary" onClick={exportExcel} disabled={saving}><Download size={17} /> Export .xlsx</button></div><div className="report-period-select"><div><b>Report frequency</b><small>Daily and weekly periods use Asia/Manila dates.</small></div><div className="segmented report-segmented">{(["Daily", "Weekly", "Monthly"] as const).map((period) => <button type="button" key={period} className={reportPeriod === period ? "selected income-selected" : ""} onClick={() => setReportPeriod(period)}>{period}</button>)}</div><span className="report-range">{reportWindow(reportPeriod).start} – {reportWindow(reportPeriod).end}</span></div><div className="report-sheet-grid"><div><FileSpreadsheet size={20} /><b>Financial Summary</b><span>Current month income, expenses, savings and remaining balance.</span></div><div><ArrowLeftRight size={20} /><b>Transaction Details</b><span>Your saved transactions, categories, accounts and dates.</span></div><div><CreditCard size={20} /><b>Budget Analysis</b><span>Budget limits, actual spending and remaining amounts.</span></div><div><Goal size={20} /><b>Savings Goals</b><span>Targets, saved amounts, progress and goal status.</span></div></div><div className="report-note"><ShieldCheck size={17} /><span>Reports are generated from your signed-in account's saved records. No sample transactions are included.</span></div></section>}
            {page === "Settings" && <div className="settings-grid"><section className="panel form-panel"><PanelTitle eyebrow="MONEY SOURCES" title="Accounts" description="Add cash, bank, e-wallet, or credit accounts." /><form className="stack-form" onSubmit={createAccount}><label>Account name<input value={accountForm.name} onChange={(e) => setAccountForm({ ...accountForm, name: e.target.value })} placeholder="e.g. Main bank account" maxLength={60} required /></label><label>Account type<select value={accountForm.accountType} onChange={(e) => setAccountForm({ ...accountForm, accountType: e.target.value })}><option value="cash">Cash</option><option value="bank">Bank account</option><option value="e-wallet">E-wallet</option><option value="credit">Credit account</option></select></label><label>Opening balance (₱)<input type="number" min="0" max="1000000000" step="0.01" value={accountForm.openingBalance} onChange={(e) => setAccountForm({ ...accountForm, openingBalance: e.target.value })} required /></label><button className="btn btn-primary btn-full" type="submit" disabled={saving}><Plus size={17} /> Add account</button></form><div className="settings-records">{accounts.length ? accounts.map((a) => <div className="setting-row" key={a.id}><span className="setting-record-icon"><Wallet size={17} /></span><div><b>{a.name}</b><small>{a.account_type}</small></div><strong>{money(a.balance_centavos ?? a.opening_balance_centavos)}</strong></div>) : <EmptyState icon={Wallet} title="No accounts yet" description="Add an account to start saving transactions." />}</div></section><section className="panel form-panel" id="category-entry"><PanelTitle eyebrow="ORGANIZE YOUR MONEY" title="Categories" description="Manage Needs, Wants, Savings, and income categories." /><form className="stack-form" onSubmit={createCategory}><label>Category name<input value={categoryForm.name} onChange={(e) => setCategoryForm({ ...categoryForm, name: e.target.value })} placeholder="e.g. Groceries" maxLength={60} required /></label><label>Category type<select value={categoryForm.kind} onChange={(e) => setCategoryForm({ ...categoryForm, kind: e.target.value, budgetGroup: e.target.value === "income" ? "income" : "needs" })}><option value="expense">Expense / allocation</option><option value="income">Income</option></select></label>{categoryForm.kind === "expense" && <label>Budget group<select value={categoryForm.budgetGroup} onChange={(e) => setCategoryForm({ ...categoryForm, budgetGroup: e.target.value })}><option value="needs">Needs — necessities</option><option value="wants">Wants — discretionary</option><option value="savings">Savings — savings categories</option></select></label>}<button className="btn btn-primary btn-full" type="submit" disabled={saving}>{editingCategoryId ? <Check size={17} /> : <Plus size={17} />}{editingCategoryId ? "Update category" : "Add category"}</button>{editingCategoryId && <button className="btn btn-secondary btn-full" type="button" onClick={() => { setEditingCategoryId(null); setCategoryForm({ name: "", kind: "expense", budgetGroup: "needs" }); }}>Cancel edit</button>}</form><div className="category-list category-manage-list">{categories.length ? categories.map((cat) => <article className="category-manage-row" key={cat.id}><span className={"category-dot " + cat.kind} /><div className="category-manage-label"><b>{cat.name}</b><small>{cat.kind === "income" ? "Income" : cat.budget_group[0].toUpperCase() + cat.budget_group.slice(1)}</small></div><button className="icon-btn" type="button" onClick={() => editCategory(cat)} aria-label={"Edit " + cat.name}><Settings size={14} /></button><button className="icon-btn delete-btn" type="button" onClick={() => deleteCategory(cat)} aria-label={"Delete " + cat.name}><Trash2 size={14} /></button></article>) : <EmptyState icon={Ellipsis} title="No categories yet" description="Add categories to organize transactions." />}</div></section><section className="panel preferences-panel"><PanelTitle eyebrow="PREFERENCES" title="Regional settings" description="These settings are used throughout your finance tracker." /><div className="preference-row"><span className="preference-icon"><Wallet size={18} /></span><div><b>Currency</b><small>Philippine peso</small></div><strong>PHP (₱)</strong></div><div className="preference-row"><span className="preference-icon"><Clock3 size={18} /></span><div><b>Time zone</b><small>Dates and reporting period</small></div><strong>Asia/Manila</strong></div><div className="preference-row"><span className="preference-icon"><ShieldCheck size={18} /></span><div><b>Data privacy</b><small>Only your authenticated account can access these records.</small></div><strong className="private-label">Private</strong></div><button className="btn btn-secondary signout-btn" onClick={signOut} disabled={saving}><LogOut size={17} /> Sign out</button></section></div>}
          </>}
          {viewTransaction && <div className="modal-backdrop" role="presentation" onClick={() => setViewTransaction(null)}>
            <section className="panel transaction-detail-modal" role="dialog" aria-modal="true" aria-labelledby="transaction-detail-title" onClick={(event) => event.stopPropagation()}>
              <div className="section-head"><div><span className="overline">TRANSACTION DETAILS</span><h2 id="transaction-detail-title">{viewTransaction.description}</h2></div><button className="icon-btn" onClick={() => setViewTransaction(null)} aria-label="Close transaction details"><X size={18} /></button></div>
              <dl className="transaction-detail-grid">
                <div><dt>Date</dt><dd>{prettyDate(viewTransaction.transaction_date)}</dd></div>
                <div><dt>Type</dt><dd>{viewTransaction.transaction_type.replace("_", " ")}</dd></div>
                <div><dt>Amount</dt><dd>{money(viewTransaction.amount_centavos)}</dd></div>
                <div><dt>Category</dt><dd>{viewTransaction.category_name || "Uncategorized"}</dd></div>
                <div><dt>Payment method</dt><dd>{viewTransaction.payment_method.replace("_", " ")}</dd></div>
                <div><dt>Account</dt><dd>{viewTransaction.account_name || "Account"}</dd></div>
                <div className="detail-notes"><dt>Notes</dt><dd>{viewTransaction.notes || "No notes added."}</dd></div>
              </dl>
              <div className="modal-actions"><button className="btn btn-secondary" onClick={() => setViewTransaction(null)}>Close</button><button className="btn btn-primary" onClick={() => { const tx = viewTransaction; setViewTransaction(null); startEditTransaction(tx); }}><Settings size={15} /> Edit transaction</button></div>
            </section>
          </div>}
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
function TransactionTable({ transactions, search, setSearch, filter, setFilter, categories, filterCategory, setFilterCategory, filterPaymentMethod, setFilterPaymentMethod, onView, onEdit, onDelete }: {
  transactions: Transaction[]; search: string; setSearch: (v: string) => void; filter: string; setFilter: (v: string) => void;
  categories: Category[]; filterCategory: string; setFilterCategory: (v: string) => void;
  filterPaymentMethod: string; setFilterPaymentMethod: (v: string) => void;
  onView: (tx: Transaction) => void; onEdit: (tx: Transaction) => void; onDelete: (id: string) => void;
}) {
  const paymentLabels: Record<string, string> = { cash: "Cash", bank_transfer: "Bank transfer", e_wallet: "E-wallet", credit_card: "Credit card", other: "Other" };
  return <>
    <div className="table-toolbar transaction-filter-toolbar">
      <label className="search-input"><Search size={16} /><input placeholder="Search description, notes, category…" value={search} onChange={(e) => setSearch(e.target.value)} aria-label="Search transactions" /></label>
      <select value={filter} onChange={(e) => setFilter(e.target.value)} aria-label="Filter transaction type"><option value="all">All types</option><option value="income">Income</option><option value="expense">Expense</option><option value="savings_contribution">Savings contribution</option></select>
      <select value={filterCategory} onChange={(e) => setFilterCategory(e.target.value)} aria-label="Filter category"><option value="all">All categories</option>{categories.map((cat) => <option key={cat.id} value={cat.id}>{cat.name}</option>)}</select>
      <select value={filterPaymentMethod} onChange={(e) => setFilterPaymentMethod(e.target.value)} aria-label="Filter payment method"><option value="all">All payment methods</option><option value="cash">Cash</option><option value="bank_transfer">Bank transfer</option><option value="e_wallet">E-wallet</option><option value="credit_card">Credit card</option><option value="other">Other</option></select>
    </div>
    {transactions.length ? <div className="table-scroll"><table><thead><tr><th>TRANSACTION</th><th>CATEGORY</th><th>DATE</th><th>PAYMENT</th><th>AMOUNT</th><th><span className="sr-only">Actions</span></th></tr></thead><tbody>{transactions.map((tx) => <tr key={tx.id}>
      <td><div className="tx-title"><span className={"tx-icon " + tx.transaction_type}>{tx.transaction_type === "income" ? <ArrowDownLeft size={16} /> : tx.transaction_type === "expense" ? <ArrowUpRight size={16} /> : <Goal size={16} />}</span><div><b>{tx.description}</b><small>{tx.account_name || "Account"} · {tx.transaction_type.replace("_", " ")}</small></div></div></td>
      <td><span className="category-label">{tx.category_name || "Uncategorized"}</span></td><td className="date-col">{prettyDate(tx.transaction_date)}</td><td>{paymentLabels[tx.payment_method] || tx.payment_method}</td>
      <td className={"money-col " + tx.transaction_type}>{tx.transaction_type === "expense" || tx.transaction_type === "savings_contribution" ? "−" : "+"}{money(tx.amount_centavos)}</td>
      <td><div className="row-actions"><button className="icon-btn" onClick={() => onView(tx)} aria-label="View transaction" title="View details"><Search size={15} /></button><button className="icon-btn" onClick={() => onEdit(tx)} aria-label="Edit transaction" title="Edit transaction"><Settings size={15} /></button><button className="icon-btn delete-btn" onClick={() => onDelete(tx.id)} aria-label="Delete transaction" title="Delete transaction"><Trash2 size={15} /></button></div></td>
    </tr>)}</tbody></table></div> : <EmptyState icon={ArrowLeftRight} title={search || filter !== "all" || filterCategory !== "all" || filterPaymentMethod !== "all" ? "No matching transactions" : "No transactions in this period"} description="Try another date range or adjust your filters." />}
    <div className="table-foot"><span>Showing {transactions.length} matching saved record{transactions.length === 1 ? "" : "s"}</span></div>
  </>;
}
function DashboardPage({ accounts, categories, goals, transactions, setPage }: { accounts: Account[]; categories: Category[]; goals: Goal[]; transactions: Transaction[]; setPage: (page: Page) => void }) {
  const [weekAnchor, setWeekAnchor] = useState(manilaToday());
  const [selectedPieSlice, setSelectedPieSlice] = useState(0);
  const weekRange = getPeriodRange(weekAnchor || manilaToday(), "Weekly");
  const monthRange = getPeriodRange(manilaToday(), "Monthly");
  const weekTransactions = transactions.filter((tx) => tx.transaction_date >= weekRange.start && tx.transaction_date <= weekRange.end);
  const weekTotals = summarizeTransactions(weekTransactions, weekRange);
  const totals = summarizeTransactions(transactions, monthRange);
  const categoryById = new Map(categories.map((category) => [category.id, category]));
  const pieData = [
    { name: "Weekly Savings", amount: weekTotals.savingsContributions, color: "#8d7bc2" },
    { name: "Needs", amount: weekTransactions.filter((tx) => tx.transaction_type === "expense" && categoryById.get(tx.category_id || "")?.budget_group === "needs").reduce((sum,tx) => sum + tx.amount_centavos,0), color: "#5b9c79" },
    { name: "Wants", amount: weekTransactions.filter((tx) => tx.transaction_type === "expense" && categoryById.get(tx.category_id || "")?.budget_group === "wants").reduce((sum,tx) => sum + tx.amount_centavos,0), color: "#e1a15e" },
  ];
  const pieTotal = pieData.reduce((sum, item) => sum + item.amount, 0);
  const pieIncome = weekTotals.income;
  let cursor = 0;
  const pieStops = pieData.map((item) => { const start = cursor; cursor += pieTotal > 0 ? item.amount / pieTotal * 100 : 0; return item.color + " " + start + "% " + cursor + "%"; }).join(", ");
  const unallocated = pieIncome - pieTotal;
  const expenseCategories = weekTransactions.filter((tx) => tx.transaction_type === "expense").reduce<Record<string, number>>((acc, tx) => { const key = tx.category_name || "Uncategorized"; acc[key] = (acc[key] || 0) + tx.amount_centavos; return acc; }, {});
  const topCategories = Object.entries(expenseCategories).sort((a, b) => b[1] - a[1]).slice(0, 6);
  const topExpenseTotal = Object.values(expenseCategories).reduce((sum,amount) => sum + amount,0);
  const activeGoals = goals.filter((g) => g.status !== "paused");
  const goalTotal = activeGoals.reduce((sum, g) => sum + g.target_amount_centavos, 0);
  const goalSaved = activeGoals.reduce((sum, g) => sum + g.current_amount_centavos, 0);
  return <>
    <section className="summary-grid">
      <MetricCard title="Total Income" value={money(totals.income)} icon={ArrowDownLeft} tone="green" caption="This month" />
      <MetricCard title="Total Expenses" value={money(totals.expenses)} icon={ArrowUpRight} tone="orange" caption="This month" />
      <MetricCard title="Savings" value={money(totals.savingsContributions)} icon={Goal} tone="violet" caption="Savings contributions" />
      <MetricCard title="Remaining Balance" value={money(totals.remainingBalance)} icon={Wallet} tone="blue" caption="Income less expenses and savings" />
    </section>
    <div className="dashboard-grid"><section className="panel overview-panel"><div className="section-head"><div><span className="overline">MONTHLY OVERVIEW</span><h2>Income & expenses</h2><p>{prettyDate(monthRange.start) + " – " + prettyDate(monthRange.end)} · Asia/Manila</p></div><span className="month-chip"><CalendarDays size={14} /> This month</span></div>
      <div className="cashflow-numbers"><div><span className="legend-mark income-mark" />Income <b>{money(totals.income)}</b></div><div><span className="legend-mark expense-mark" />Expenses <b>{money(totals.expenses)}</b></div></div>
      <div className="cashflow-bars"><div className="bar-group"><span>Income</span><div className="bar-bg"><span className="bar-income" style={{ width: (totals.income ? Math.max(2, totals.income / Math.max(totals.income, totals.expenses, totals.savingsContributions, 1) * 100) : 0) + "%" }} /></div><b>{money(totals.income)}</b></div><div className="bar-group"><span>Expenses</span><div className="bar-bg"><span className="bar-expense" style={{ width: (totals.expenses ? Math.max(2, totals.expenses / Math.max(totals.income, totals.expenses, totals.savingsContributions, 1) * 100) : 0) + "%" }} /></div><b>{money(totals.expenses)}</b></div><div className="bar-group"><span>Savings</span><div className="bar-bg"><span className="bar-savings" style={{ width: (totals.savingsContributions ? Math.max(2, totals.savingsContributions / Math.max(totals.income, totals.expenses, totals.savingsContributions, 1) * 100) : 0) + "%" }} /></div><b>{money(totals.savingsContributions)}</b></div></div>
      <div className={"net-card " + (totals.remainingBalance >= 0 ? "net-positive" : "net-negative")}><span>{totals.remainingBalance >= 0 ? <TrendingUp size={18} /> : <ArrowDownRight size={18} />}Remaining after expenses and savings</span><b>{money(totals.remainingBalance)}</b></div>
    </section>
    <section className="panel goal-overview-panel"><div className="section-head"><div><span className="overline">YOUR NEXT MILESTONES</span><h2>Savings goals</h2><p>Every little bit counts.</p></div><button className="text-link" onClick={() => setPage("Savings Goals")}>View all <ChevronRight size={15} /></button></div>
      {activeGoals.length ? <><div className="goal-summary"><div className="goal-ring" style={{ background: "conic-gradient(#69a681 " + Math.min(100, goalTotal ? goalSaved / goalTotal * 100 : 0) + "%, #edf2ec 0)" }}><span><b>{goalTotal ? Math.round(goalSaved / goalTotal * 100) : 0}%</b><small>saved</small></span></div><div><b>{money(goalSaved)}</b><small>of {money(goalTotal)} target</small><div className="goal-summary-note">{activeGoals.length} active goal{activeGoals.length === 1 ? "" : "s"}</div></div></div><div className="mini-goal-list">{activeGoals.slice(0, 3).map((g) => <div className="mini-goal" key={g.id}><div><span>{g.name}</span><b>{Math.round(g.current_amount_centavos / Math.max(1, g.target_amount_centavos) * 100)}%</b></div><div className="progress-track"><span style={{ width: Math.min(100, g.current_amount_centavos / Math.max(1, g.target_amount_centavos) * 100) + "%" }} /></div></div>)}</div></> : <EmptyState icon={Goal} title="Start a savings goal" description="Set a target and keep track of your progress." />}
      <button className="outline-full" onClick={() => setPage("Savings Goals")}><Plus size={16} /> Manage savings goals</button>
    </section></div>
    <section className="panel weekly-chart-panel">
      <div className="section-head"><div><span className="overline">WEEKLY ALLOCATION</span><h2>Savings, needs & wants</h2><p>Category amounts are compared with total income for the selected week.</p></div>
        <div className="period-date-controls"><button className="btn btn-secondary btn-small" type="button" onClick={() => setWeekAnchor(shiftPeriod(weekAnchor || manilaToday(),"Weekly",-1))}><ChevronRight size={14} style={{transform:"rotate(180deg)"}} /> Previous</button><label>Week anchor<input type="date" value={weekAnchor} onChange={(e) => setWeekAnchor(e.target.value)} /></label><button className="btn btn-secondary btn-small" type="button" onClick={() => setWeekAnchor(shiftPeriod(weekAnchor || manilaToday(),"Weekly",1))}>Next <ChevronRight size={14} /></button></div>
      </div>
      <p className="active-range-label">{prettyDate(weekRange.start)} – {prettyDate(weekRange.end)} · Monday–Sunday · Asia/Manila</p>
      {pieIncome <= 0 ? <EmptyState icon={BarChart3} title="No income for this week" description="The weekly allocation chart appears when the selected week has income. There is no division by zero." /> : <div className="weekly-chart-layout">
        <div className="allocation-pie-wrap"><button className="allocation-pie" type="button" aria-label={"Weekly allocation pie chart. Selected slice: " + pieData[selectedPieSlice].name} style={{background: pieTotal > 0 ? "conic-gradient(" + pieStops + ")" : "conic-gradient(#edf2ec 0 100%)"}} onClick={() => setSelectedPieSlice((selectedPieSlice+1)%pieData.length)}><span><small>Weekly income</small><b>{money(pieIncome)}</b><small>{pieTotal > pieIncome ? "Allocation exceeds income" : "Select a slice"}</small></span></button><p className="chart-interaction-hint">Select the chart or a legend item to inspect each category.</p></div>
        <div className="allocation-legend">{pieData.map((item,index) => { const pct = pieIncome > 0 ? item.amount / pieIncome * 100 : 0; return <button type="button" key={item.name} className={"allocation-legend-item " + (selectedPieSlice === index ? "active" : "")} onClick={() => setSelectedPieSlice(index)}><span className="allocation-legend-dot" style={{background:item.color}} /><span className="allocation-legend-copy"><b>{item.name}</b><small>{money(item.amount)} · {pct.toFixed(1)}% of income</small></span><strong>{pct.toFixed(1)}%</strong></button>; })}</div>
        <div className="allocation-summary"><span>Selected category</span><b>{pieData[selectedPieSlice].name}</b><strong>{money(pieData[selectedPieSlice].amount)}</strong><small>{(pieIncome > 0 ? pieData[selectedPieSlice].amount / pieIncome * 100 : 0).toFixed(1)}% of weekly income</small></div>
      </div>}
      {pieIncome > 0 && <div className={"unallocated-summary " + (unallocated < 0 ? "negative" : "")}><div><b>{unallocated < 0 ? "Allocation above income" : "Unallocated income"}</b><small>{unallocated < 0 ? "Needs, wants, and savings are greater than income for this week." : "Income not represented by the three chart categories; it is not assigned to another slice."}</small></div><strong>{money(Math.abs(unallocated))}{unallocated < 0 ? " over" : ""}</strong></div>}
    </section>
    <div className="dashboard-grid lower-grid">
      <section className="panel category-overview"><div className="section-head"><div><span className="overline">WEEKLY SPENDING</span><h2>Spending by category</h2><p>{prettyDate(weekRange.start)} – {prettyDate(weekRange.end)}</p></div></div>
        {topCategories.length ? <div className="category-breakdown">{topCategories.map(([name, amount], i) => <button type="button" className="breakdown-row interactive-breakdown" key={name} onClick={() => setPage("Transactions")}><div className="breakdown-label"><span className={"breakdown-dot dot-" + i} /><span>{name}</span><b>{money(amount)}</b></div><div className="progress-track"><span className={"breakdown-bar bar-" + i} style={{ width: (amount / Math.max(topExpenseTotal, 1) * 100) + "%" }} /></div><small>{(amount / Math.max(topExpenseTotal,1) * 100).toFixed(1)}% of weekly expenses</small></button>)}</div> : <EmptyState icon={BarChart3} title="No spending breakdown" description="Categorized expenses from this week will appear here." />}
      </section>
      <section className="panel category-overview"><div className="section-head"><div><span className="overline">SAVINGS PROGRESS</span><h2>Goal progress chart</h2><p>Saved amounts compared with each target.</p></div><button className="text-link" onClick={() => setPage("Savings Goals")}>Manage <ChevronRight size={15} /></button></div>
        {goals.length ? <div className="savings-progress-chart">{goals.map((goal) => { const pct = goal.target_amount_centavos > 0 ? goal.current_amount_centavos / goal.target_amount_centavos * 100 : 0; return <button type="button" className="savings-progress-row" key={goal.id} onClick={() => setPage("Savings Goals")}><div><span>{goal.name}</span><b>{Math.min(100,Math.round(pct))}%</b></div><div className="progress-track"><span style={{width:Math.min(100,Math.max(0,pct))+"%"}} /></div><small>{money(goal.current_amount_centavos)} saved of {money(goal.target_amount_centavos)}</small></button>; })}</div> : <EmptyState icon={Goal} title="No savings goals" description="Create a goal to visualize your progress." />}
      </section>
    </div>
    <div className="dashboard-grid lower-grid"><section className="panel recent-panel"><div className="section-head"><div><span className="overline">LATEST ACTIVITY</span><h2>Recent transactions</h2><p>Your latest saved money movements.</p></div><button className="text-link" onClick={() => setPage("Transactions")}>View all <ChevronRight size={15} /></button></div>
      {transactions.length ? <div className="recent-list">{transactions.slice(0, 5).map((tx) => <div className="recent-row" key={tx.id}><span className={"tx-icon " + tx.transaction_type}>{tx.transaction_type === "income" ? <ArrowDownLeft size={16} /> : tx.transaction_type === "savings_contribution" ? <Goal size={16} /> : <ArrowUpRight size={16} />}</span><div className="recent-desc"><b>{tx.description || (tx.transaction_type === "income" ? "Income" : "Expense")}</b><small>{tx.category_name || "Uncategorized"} · {prettyDate(tx.transaction_date)}</small></div><b className={"recent-amount " + tx.transaction_type}>{tx.transaction_type === "expense" || tx.transaction_type === "savings_contribution" ? "−" : "+"}{money(tx.amount_centavos)}</b></div>)}</div> : <EmptyState icon={ArrowLeftRight} title="Nothing to show yet" description="Your saved transactions will appear here." />}
    </section><section className="panel account-overview-panel"><div className="section-head"><div><span className="overline">YOUR ACCOUNTS</span><h2>Money accounts</h2><p>Available account balances.</p></div><button className="text-link" onClick={() => setPage("Settings")}>Manage <ChevronRight size={15} /></button></div>
      {accounts.length ? <div className="account-list">{accounts.map((a) => <div className="account-row" key={a.id}><span className="account-icon"><Wallet size={16} /></span><span>{a.name}</span><b>{money(a.balance_centavos ?? a.opening_balance_centavos)}</b></div>)}</div> : <EmptyState icon={Wallet} title="No accounts yet" description="Add a money account to get started." />}
    </section></div>
    {accounts.length === 0 && <div className="getting-started"><div className="getting-icon"><Wallet size={20} /></div><div><b>First, add a money account</b><p>Set up a cash, bank, or e-wallet account before recording transactions.</p></div><button className="btn btn-primary btn-small" onClick={() => setPage("Settings")}>Set up account <ChevronRight size={15} /></button></div>}
  </>;
}

function MetricCard({ title, value, icon: Icon, tone, caption }: { title: string; value: string; icon: typeof Wallet; tone: string; caption: string }) {
  return <article className="metric-card"><div className="metric-top"><span>{title}</span><span className={"metric-icon " + tone}><Icon size={18} /></span></div><div className="metric-value">{value}</div><div className="metric-caption"><span className={"caption-dot " + tone} />{caption}</div></article>;
}
