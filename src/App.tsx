import { useMemo, useState } from "react";
import ExcelJS from "exceljs";
import {
  ArrowDownLeft, ArrowUpRight, CalendarDays, ChevronDown, CircleHelp,
  CreditCard, Download, Ellipsis, Goal, LayoutDashboard, LogOut, Plus,
  Search, Settings, ShieldCheck, Sparkles, Wallet, X,
} from "lucide-react";
import {
  Area, AreaChart, CartesianGrid, ResponsiveContainer, Tooltip, XAxis, YAxis,
} from "recharts";

type Period = "Daily" | "Weekly" | "Monthly";

const demoTransactions = [
  { id: "TX-001", name: "Monthly salary", category: "Salary", date: "Oct 24, 2026", amount: 20000, type: "income" as const, icon: "↗" },
  { id: "TX-002", name: "Grocery shopping", category: "Food & groceries", date: "Oct 23, 2026", amount: 1250, type: "expense" as const, icon: "🛒" },
  { id: "TX-003", name: "Electric bill", category: "Utilities", date: "Oct 22, 2026", amount: 1840, type: "expense" as const, icon: "⚡" },
  { id: "TX-004", name: "Commute", category: "Transportation", date: "Oct 21, 2026", amount: 180, type: "expense" as const, icon: "🚌" },
  { id: "TX-005", name: "Freelance project", category: "Side income", date: "Oct 19, 2026", amount: 3500, type: "income" as const, icon: "✦" },
];

const chartData: Record<Period, { label: string; income: number; expenses: number }[]> = {
  Daily: [
    { label: "Mon", income: 0, expenses: 180 }, { label: "Tue", income: 0, expenses: 420 },
    { label: "Wed", income: 3500, expenses: 650 }, { label: "Thu", income: 0, expenses: 180 },
    { label: "Fri", income: 0, expenses: 1840 }, { label: "Sat", income: 0, expenses: 1250 },
    { label: "Sun", income: 20000, expenses: 0 },
  ],
  Weekly: [
    { label: "Week 1", income: 20000, expenses: 5800 }, { label: "Week 2", income: 0, expenses: 4100 },
    { label: "Week 3", income: 3500, expenses: 3260 }, { label: "Week 4", income: 20000, expenses: 3470 },
  ],
  Monthly: [
    { label: "May", income: 23000, expenses: 15500 }, { label: "Jun", income: 23500, expenses: 14800 },
    { label: "Jul", income: 22000, expenses: 16100 }, { label: "Aug", income: 24500, expenses: 14300 },
    { label: "Sep", income: 23500, expenses: 17200 }, { label: "Oct", income: 23500, expenses: 10620 },
  ],
};

const money = (amount: number) => new Intl.NumberFormat("en-PH", {
  style: "currency", currency: "PHP", maximumFractionDigits: 2,
}).format(amount);

export default function App() {
  const [period, setPeriod] = useState<Period>("Monthly");
  const [activeNav, setActiveNav] = useState("Overview");
  const [query, setQuery] = useState("");
  const [notice, setNotice] = useState("");
  const [showDemoNotice, setShowDemoNotice] = useState(true);

  const filteredTransactions = useMemo(() => demoTransactions.filter((item) =>
    [item.name, item.category, item.date].join(" ").toLowerCase().includes(query.toLowerCase())
  ), [query]);

  async function exportReport() {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Budgetwise";
    workbook.created = new Date();

    const summary = workbook.addWorksheet("Financial Summary");
    summary.columns = [
      { header: "Metric", key: "metric", width: 28 },
      { header: "Amount (PHP)", key: "amount", width: 20 },
    ];
    summary.addRows([
      { metric: "Report period", amount: period },
      { metric: "Total income", amount: 23500 },
      { metric: "Total expenses", amount: 10620 },
      { metric: "Net cash flow", amount: 12880 },
      { metric: "Savings goal progress", amount: "42%" },
    ]);
    summary.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
    summary.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF253B35" } };
    summary.getColumn(2).numFmt = '"₱"#,##0.00;[Red]-"₱"#,##0.00';

    const transactions = workbook.addWorksheet("Transaction Details");
    transactions.columns = [
      { header: "Transaction ID", key: "id", width: 18 },
      { header: "Description", key: "name", width: 28 },
      { header: "Category", key: "category", width: 24 },
      { header: "Date", key: "date", width: 20 },
      { header: "Type", key: "type", width: 14 },
      { header: "Amount (PHP)", key: "amount", width: 18 },
    ];
    transactions.addRows(demoTransactions.map((item) => ({
      ...item, amount: item.type === "expense" ? -item.amount : item.amount,
    })));
    transactions.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
    transactions.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF253B35" } };
    transactions.getColumn(6).numFmt = '"₱"#,##0.00;[Red]-"₱"#,##0.00';
    transactions.autoFilter = { from: "A1", to: "F1" };

    const budgets = workbook.addWorksheet("Budget Analysis");
    budgets.columns = [
      { header: "Category", key: "category", width: 24 },
      { header: "Budget (PHP)", key: "budget", width: 18 },
      { header: "Spent (PHP)", key: "spent", width: 18 },
      { header: "Remaining (PHP)", key: "remaining", width: 20 },
    ];
    budgets.addRows([
      { category: "Food & groceries", budget: 6000, spent: 3250, remaining: 2750 },
      { category: "Utilities", budget: 3500, spent: 1840, remaining: 1660 },
      { category: "Transportation", budget: 2000, spent: 980, remaining: 1020 },
    ]);
    budgets.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
    budgets.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF253B35" } };
    for (const col of [2, 3, 4]) budgets.getColumn(col).numFmt = '"₱"#,##0.00;[Red]-"₱"#,##0.00';

    const goals = workbook.addWorksheet("Savings Goals");
    goals.columns = [
      { header: "Goal", key: "goal", width: 28 },
      { header: "Target (PHP)", key: "target", width: 18 },
      { header: "Saved (PHP)", key: "saved", width: 18 },
      { header: "Progress", key: "progress", width: 16 },
    ];
    goals.addRows([
      { goal: "Emergency fund", target: 50000, saved: 21000, progress: 0.42 },
      { goal: "Travel fund", target: 25000, saved: 8500, progress: 0.34 },
    ]);
    goals.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
    goals.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF253B35" } };
    goals.getColumn(2).numFmt = '"₱"#,##0.00';
    goals.getColumn(3).numFmt = '"₱"#,##0.00';
    goals.getColumn(4).numFmt = "0%";

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const objectUrl = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = objectUrl;
    link.download = `budgetwise-demo-report-${period.toLowerCase()}-${new Date().toISOString().slice(0, 10)}.xlsx`;
    link.click();
    URL.revokeObjectURL(objectUrl);
    setNotice("Downloaded a sample Excel report. Connect sign-in and the database before using real records.");
  }

  function navigate(item: string) {
    setActiveNav(item);
    if (item !== "Overview") setNotice(`${item} page is planned for the next implementation step.`);
    else setNotice("");
  }

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <a className="brand" href="#" onClick={(event) => { event.preventDefault(); navigate("Overview"); }}>
          <span className="brand-mark"><Wallet size={21} /></span>
          <span>budgetwise<span className="brand-dot">.</span><small>PERSONAL FINANCE</small></span>
        </a>
        <div className="side-label">WORKSPACE</div>
        <nav className="nav-list">
          <button className={activeNav === "Overview" ? "nav-item active" : "nav-item"} onClick={() => navigate("Overview")}><LayoutDashboard size={18} />Overview</button>
          <button className={activeNav === "Transactions" ? "nav-item active" : "nav-item"} onClick={() => navigate("Transactions")}><ArrowUpRight size={18} />Transactions</button>
          <button className={activeNav === "Budgets" ? "nav-item active" : "nav-item"} onClick={() => navigate("Budgets")}><CreditCard size={18} />Budgets</button>
          <button className={activeNav === "Savings goals" ? "nav-item active" : "nav-item"} onClick={() => navigate("Savings goals")}><Goal size={18} />Savings goals</button>
          <button className={activeNav === "Reports" ? "nav-item active" : "nav-item"} onClick={() => navigate("Reports")}><CalendarDays size={18} />Reports</button>
        </nav>
        <div className="sidebar-bottom">
          <div className="secure-note"><ShieldCheck size={17} /><span><b>Your money, your privacy</b><small>Designed for isolated user data.</small></span></div>
          <button className="nav-item" onClick={() => navigate("Settings")}><Settings size={18} />Settings</button>
          <div className="profile-row"><div className="avatar">D</div><div className="profile-copy"><b>Demo account</b><small>Preview mode</small></div><Ellipsis size={19} /></div>
        </div>
      </aside>

      <main className="main-content">
        <header className="topbar">
          <div className="breadcrumb">Workspace <span>/</span> <b>{activeNav}</b></div>
          <div className="top-actions"><span className="demo-chip"><span /> Demo data</span><button className="icon-button" aria-label="Help" onClick={() => setNotice("This is a UI prototype. Authentication and persistent storage are not enabled yet.")}><CircleHelp size={19} /></button><div className="avatar top-avatar">D</div></div>
        </header>

        <div className="page-content">
          <div className="welcome-row">
            <div><div className="eyebrow">SATURDAY, OCTOBER 24, 2026</div><h1>Your money, in balance.</h1><p className="subtitle">A clearer view of your spending, savings, and what comes next.</p></div>
            <button className="primary-button" onClick={() => setNotice("Transaction entry will be enabled after authentication and database setup.")}><Plus size={17} /> Add transaction</button>
          </div>

          {showDemoNotice && <div className="prototype-banner"><Sparkles size={18} /><div><b>You're viewing a prototype</b><span>These figures are sample data, not saved personal finances. Sign-in and private database storage are the next setup step.</span></div><button aria-label="Dismiss prototype notice" onClick={() => setShowDemoNotice(false)}><X size={17} /></button></div>}
          {notice && <div className="notice-banner" role="status">{notice}<button onClick={() => setNotice("")} aria-label="Dismiss notice"><X size={15} /></button></div>}

          <section className="metric-grid">
            <article className="metric-card"><div className="metric-top"><span>Total balance</span><span className="metric-icon green"><Wallet size={18} /></span></div><div className="metric-value">₱42,580.00</div><div className="metric-foot"><span className="trend up"><ArrowUpRight size={14} /> 8.2%</span><span>vs. last month</span></div><div className="metric-spark green-spark"><i /><i /><i /><i /><i /><i /><i /><i /></div></article>
            <article className="metric-card"><div className="metric-top"><span>Income</span><span className="metric-icon blue"><ArrowDownLeft size={18} /></span></div><div className="metric-value">₱23,500.00</div><div className="metric-foot"><span className="trend up"><ArrowUpRight size={14} /> 4.5%</span><span>vs. last month</span></div><div className="metric-spark blue-spark"><i /><i /><i /><i /><i /><i /><i /><i /></div></article>
            <article className="metric-card"><div className="metric-top"><span>Expenses</span><span className="metric-icon orange"><ArrowUpRight size={18} /></span></div><div className="metric-value">₱10,620.00</div><div className="metric-foot"><span className="trend down">↓ 2.1%</span><span>vs. last month</span></div><div className="metric-spark orange-spark"><i /><i /><i /><i /><i /><i /><i /><i /></div></article>
            <article className="metric-card"><div className="metric-top"><span>Savings rate</span><span className="metric-icon violet"><Goal size={18} /></span></div><div className="metric-value">54.8%</div><div className="metric-foot"><span className="trend up"><ArrowUpRight size={14} /> 3.6%</span><span>vs. last month</span></div><div className="savings-meter"><span /></div></article>
          </section>

          <section className="content-grid">
            <article className="panel cashflow-panel">
              <div className="panel-header"><div><h2>Cash flow</h2><p>Track money coming in and going out</p></div><div className="period-tabs">{(["Daily", "Weekly", "Monthly"] as Period[]).map((item) => <button key={item} className={period === item ? "selected" : ""} onClick={() => setPeriod(item)}>{item}</button>)}</div></div>
              <div className="chart-summary"><div><span className="legend-dot income-dot" /> Income <b>{money(period === "Daily" ? 23500 : period === "Weekly" ? 43500 : 23500)}</b></div><div><span className="legend-dot expense-dot" /> Expenses <b>{money(period === "Daily" ? 4520 : period === "Weekly" ? 16630 : 10620)}</b></div></div>
              <div className="chart-wrap"><ResponsiveContainer width="100%" height="100%"><AreaChart data={chartData[period]} margin={{ top: 10, right: 6, left: -16, bottom: 0 }}><defs><linearGradient id="incomeFill" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#4e9b7a" stopOpacity={0.19} /><stop offset="95%" stopColor="#4e9b7a" stopOpacity={0} /></linearGradient><linearGradient id="expenseFill" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#e6a36f" stopOpacity={0.18} /><stop offset="95%" stopColor="#e6a36f" stopOpacity={0} /></linearGradient></defs><CartesianGrid strokeDasharray="3 5" vertical={false} stroke="#edf0ed" /><XAxis dataKey="label" axisLine={false} tickLine={false} tick={{ fill: "#929b96", fontSize: 11 }} dy={10} /><YAxis axisLine={false} tickLine={false} tick={{ fill: "#929b96", fontSize: 11 }} tickFormatter={(value: number) => value === 0 ? "₱0" : `₱${value / 1000}k`} /><Tooltip formatter={(value) => money(Number(value))} contentStyle={{ borderRadius: 12, border: "1px solid #e8ece9", fontSize: 12 }} /><Area type="monotone" dataKey="income" name="Income" stroke="#4e9b7a" strokeWidth={2.5} fill="url(#incomeFill)" /><Area type="monotone" dataKey="expenses" name="Expenses" stroke="#e6a36f" strokeWidth={2.5} fill="url(#expenseFill)" /></AreaChart></ResponsiveContainer></div>
            </article>

            <article className="panel goals-panel">
              <div className="panel-header"><div><h2>Savings goals</h2><p>Small steps add up</p></div><button className="more-button" onClick={() => navigate("Savings goals")} aria-label="View savings goals"><Ellipsis size={20} /></button></div>
              <div className="goal-item"><div className="goal-illustration goal-green"><ShieldCheck size={20} /></div><div className="goal-details"><div className="goal-title"><b>Emergency fund</b><span>42%</span></div><div className="progress-track"><span style={{ width: "42%" }} /></div><div className="goal-amount"><b>₱21,000</b><span>of ₱50,000</span></div></div></div>
              <div className="goal-item"><div className="goal-illustration goal-blue"><Sparkles size={20} /></div><div className="goal-details"><div className="goal-title"><b>Travel fund</b><span>34%</span></div><div className="progress-track blue-track"><span style={{ width: "34%" }} /></div><div className="goal-amount"><b>₱8,500</b><span>of ₱25,000</span></div></div></div>
              <button className="goal-add-button" onClick={() => navigate("Savings goals")}><Plus size={16} /> Create savings goal</button>
            </article>
          </section>

          <section className="panel transactions-panel">
            <div className="panel-header transaction-header"><div><h2>Recent transactions</h2><p>Your latest money movements</p></div><div className="transaction-actions"><label className="search-box"><Search size={16} /><input aria-label="Search transactions" placeholder="Search transactions" value={query} onChange={(event) => setQuery(event.target.value)} /></label><button className="secondary-button" onClick={exportReport}><Download size={16} /> Export Excel</button><button className="view-all" onClick={() => navigate("Transactions")}>View all <ChevronDown size={15} /></button></div></div>
            <div className="table-scroll"><table><thead><tr><th>TRANSACTION</th><th>CATEGORY</th><th>DATE</th><th>TYPE</th><th className="amount-cell">AMOUNT</th></tr></thead><tbody>{filteredTransactions.map((item) => <tr key={item.id}><td><div className="transaction-name"><span className={item.type === "income" ? "transaction-icon income-bg" : "transaction-icon expense-bg"}>{item.icon}</span><span><b>{item.name}</b><small>{item.id}</small></span></div></td><td><span className="category-pill">{item.category}</span></td><td className="date-cell">{item.date}</td><td><span className={item.type === "income" ? "type-pill income-type" : "type-pill expense-type"}>{item.type}</span></td><td className={item.type === "income" ? "amount-cell positive" : "amount-cell negative"}>{item.type === "income" ? "+" : "−"}{money(item.amount)}</td></tr>)}</tbody></table>{filteredTransactions.length === 0 && <div className="empty-state">No matching sample transactions.</div>}</div>
            <div className="table-footer"><span>Showing {filteredTransactions.length} of {demoTransactions.length} sample transactions</span><button onClick={() => navigate("Transactions")}>Open transactions <ArrowUpRight size={14} /></button></div>
          </section>

          <footer className="page-footer"><span>© 2026 Budgetwise <span className="footer-dot">·</span> Built for better money habits</span><span><ShieldCheck size={14} /> Private by design <span className="footer-dot">·</span> <button onClick={() => setNotice("Privacy and account controls will be documented before public launch.")}>Privacy</button></span></footer>
        </div>
      </main>
    </div>
  );
}