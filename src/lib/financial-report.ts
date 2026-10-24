import ExcelJS from "exceljs";
import type { DateRange, FinancialTransaction, PeriodKind } from "./finance";

export type ReportTransaction = FinancialTransaction & {
  id: string;
  category_id: string | null;
  goal_id?: string | null;
  description: string;
  payment_method: string;
  notes?: string | null;
  category_name?: string | null;
  account_name?: string;
};
export type ReportCategory = { id: string; name: string; kind: "income" | "expense"; budget_group: "needs" | "wants" | "savings" | "income" };
export type ReportBudget = { id: string; category_id: string; category_name: string; period_start: string; period_end: string; amount_centavos: number };
export type ReportGoal = { id: string; name: string; description?: string | null; target_amount_centavos: number; current_amount_centavos: number; target_date: string | null; status: string };
export type FinancialReportInput = {
  period: PeriodKind;
  range: DateRange;
  transactions: ReportTransaction[];
  categories: ReportCategory[];
  budgets: ReportBudget[];
  goals: ReportGoal[];
  exportedAt?: Date;
};

const pesosFormat = '"₱"#,##0.00;[Red]-"₱"#,##0.00';
const percentFormat = "0.0%";
const dateFormat = "yyyy-mm-dd";
const validDate = (value: string) => /^\d{4}-\d{2}-\d{2}$/.test(value) && !Number.isNaN(Date.parse(value + "T00:00:00Z"));
const dayNumber = (value: string) => Date.parse(value + "T00:00:00Z") / 86_400_000;
const intersects = (aStart: string, aEnd: string, bStart: string, bEnd: string) => aStart <= bEnd && aEnd >= bStart;

function styleTable(sheet: ExcelJS.Worksheet, headerRow: number, lastColumn: number, freezeRows: number, filter = false) {
  const header = sheet.getRow(headerRow);
  header.height = 28;
  header.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 10 };
  header.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF285A43" } };
  header.alignment = { vertical: "middle", wrapText: true };
  for (let rowIndex = 1; rowIndex <= sheet.rowCount; rowIndex++) {
    const row = sheet.getRow(rowIndex);
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.border = {
        top: { style: "thin", color: { argb: "FFE2E9E2" } },
        left: { style: "thin", color: { argb: "FFE2E9E2" } },
        bottom: { style: "thin", color: { argb: "FFE2E9E2" } },
        right: { style: "thin", color: { argb: "FFE2E9E2" } },
      };
      if (rowIndex > headerRow && rowIndex % 2 === 0) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF6F8F5" } };
      cell.alignment = { vertical: "middle", wrapText: rowIndex === headerRow };
    });
  }
  sheet.views = [{ state: "frozen", ySplit: freezeRows }];
  sheet.properties.defaultRowHeight = 20;
  if (filter && lastColumn > 0 && sheet.rowCount >= headerRow) sheet.autoFilter = { from: { row: headerRow, column: 1 }, to: { row: Math.max(headerRow, sheet.rowCount), column: lastColumn } };
}

function addReportHeading(sheet: ExcelJS.Worksheet, title: string, range: DateRange, exportDate: string, lastColumn: number) {
  sheet.mergeCells(1, 1, 1, lastColumn);
  sheet.getCell(1, 1).value = title;
  sheet.getCell(1, 1).font = { bold: true, size: 16, color: { argb: "FF285A43" } };
  sheet.getCell(1, 1).alignment = { vertical: "middle" };
  sheet.getRow(1).height = 30;
  sheet.mergeCells(2, 1, 2, lastColumn);
  sheet.getCell(2, 1).value = `Reporting period: ${range.start} to ${range.end} (Asia/Manila)`;
  sheet.getCell(2, 1).font = { italic: true, color: { argb: "FF647367" } };
  sheet.mergeCells(3, 1, 3, lastColumn);
  sheet.getCell(3, 1).value = `Exported: ${exportDate} (Asia/Manila)`;
  sheet.getCell(3, 1).font = { color: { argb: "FF647367" }, size: 9 };
}

function formulaCell(formula: string, result: number | string): ExcelJS.CellValue {
  return { formula, result };
}

export function buildFinancialWorkbook(input: FinancialReportInput): ExcelJS.Workbook {
  if (!validDate(input.range.start) || !validDate(input.range.end) || input.range.start > input.range.end) {
    throw new Error("Choose a valid reporting date range.");
  }
  const exportedAt = input.exportedAt ?? new Date();
  const exportDate = new Intl.DateTimeFormat("en-PH", { dateStyle: "medium", timeStyle: "short", timeZone: "Asia/Manila" }).format(exportedAt);
  const transactions = input.transactions
    .filter((tx) => tx.transaction_date >= input.range.start && tx.transaction_date <= input.range.end)
    .sort((a, b) => a.transaction_date.localeCompare(b.transaction_date) || a.id.localeCompare(b.id));
  const categoryById = new Map(input.categories.map((category) => [category.id, category]));
  const categoryFor = (tx: ReportTransaction) => tx.category_id ? categoryById.get(tx.category_id) : undefined;
  const cents = (values: number[]) => values.reduce((sum, value) => sum + value, 0);
  const incomeCents = cents(transactions.filter((tx) => tx.transaction_type === "income").map((tx) => tx.amount_centavos));
  const expenseCents = cents(transactions.filter((tx) => tx.transaction_type === "expense").map((tx) => tx.amount_centavos));
  const savingsCents = cents(transactions.filter((tx) => tx.transaction_type === "savings_contribution").map((tx) => tx.amount_centavos));
  const needsCents = cents(transactions.filter((tx) => tx.transaction_type === "expense" && categoryFor(tx)?.budget_group === "needs").map((tx) => tx.amount_centavos));
  const wantsCents = cents(transactions.filter((tx) => tx.transaction_type === "expense" && categoryFor(tx)?.budget_group === "wants").map((tx) => tx.amount_centavos));
  const goalSavedCents = cents(input.goals.map((goal) => goal.current_amount_centavos));
  const goalTargetCents = cents(input.goals.map((goal) => goal.target_amount_centavos));

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Budgetwise";
  workbook.subject = "Private financial report";
  workbook.title = `Financial Report - ${input.period} - ${input.range.start} to ${input.range.end}`;
  workbook.created = exportedAt;
  workbook.modified = exportedAt;
  workbook.calcProperties.fullCalcOnLoad = true;
  workbook.calcProperties.forceFullCalc = true;

  const summary = workbook.addWorksheet("Financial Summary");
  summary.columns = [{ width: 35 }, { width: 22 }, { width: 16 }];
  addReportHeading(summary, "BUDGETWISE FINANCIAL SUMMARY", input.range, exportDate, 3);
  summary.getRow(5).values = ["Financial Metric", "Amount (PHP)", "Details"];
  const metrics: Array<[string, ExcelJS.CellValue, string]> = [
    ["Total Income", formulaCell('SUMIF(Transactions!$C:$C,"Income",Transactions!$F:$F)', incomeCents / 100), "Income transactions in the selected period"],
    ["Total Expenses", formulaCell('SUMIF(Transactions!$C:$C,"Expense",Transactions!$F:$F)', expenseCents / 100), "Expense transactions only; savings contributions excluded"],
    ["Total Savings Contributions", formulaCell('SUMIF(Transactions!$C:$C,"Savings Contribution",Transactions!$F:$F)', savingsCents / 100), "Money allocated to savings, not an ordinary expense"],
    ["Remaining Balance", formulaCell("B6-B7-B8", (incomeCents - expenseCents - savingsCents) / 100), "Income minus expenses minus savings contributions"],
    ["Total Needs", formulaCell('SUMIFS(Transactions!$F:$F,Transactions!$H:$H,"Needs",Transactions!$C:$C,"Expense")', needsCents / 100), "Needs-group expenses in the selected period"],
    ["Total Wants", formulaCell('SUMIFS(Transactions!$F:$F,Transactions!$H:$H,"Wants",Transactions!$C:$C,"Expense")', wantsCents / 100), "Wants-group expenses in the selected period"],
    ["Current Saved Across Goals", goalSavedCents / 100, "Current lifetime saved amount; not limited to this reporting period"],
    ["Combined Goal Targets", goalTargetCents / 100, "Combined target amount for all savings goals"],
    ["Overall Goal Completion", formulaCell("IF(B13=0,0,B12/B13)", goalTargetCents > 0 ? goalSavedCents / goalTargetCents : 0), "Current saved amount divided by combined targets"],
  ];
  metrics.forEach(([metric, amount, details], index) => {
    const row = 6 + index;
    summary.getCell(row, 1).value = metric;
    summary.getCell(row, 2).value = amount;
    summary.getCell(row, 3).value = details;
    if (metric === "Overall Goal Completion") summary.getCell(row, 2).numFmt = percentFormat;
    else if (row !== 14) summary.getCell(row, 2).numFmt = pesosFormat;
  });
  summary.getCell(16, 1).value = "Category Breakdown";
  summary.getCell(16, 1).font = { bold: true, size: 12, color: { argb: "FF285A43" } };
  summary.getRow(17).values = ["Category", "Group", "Amount (PHP)", "% of Total Income"];
  summary.getRow(17).getCell(4).value = "% of Total Income";
  summary.getColumn(4).width = 20;
  const reportCategoryRows = input.categories.filter((category) => category.kind === "income" || ["needs", "wants", "savings"].includes(category.budget_group)).map((category) => ({ id: category.id, name: category.name, budget_group: category.budget_group }));
  const hasUncategorized = transactions.some((tx) => tx.transaction_type !== "transfer" && (!tx.category_id || !categoryById.has(tx.category_id)));
  if (hasUncategorized) reportCategoryRows.push({ id: "__uncategorized__", name: "Uncategorized", budget_group: "savings" });
  reportCategoryRows.forEach((category, index) => {
    const row = 18 + index;
    const unallocatedRow = category.id === "__uncategorized__";
    const amountCents = cents(transactions.filter((tx) => tx.transaction_type !== "transfer" && (unallocatedRow ? (!tx.category_id || !categoryById.has(tx.category_id)) : tx.category_id === category.id)).map((tx) => tx.amount_centavos));
    summary.getCell(row, 1).value = category.name;
    summary.getCell(row, 2).value = unallocatedRow ? "Unallocated" : category.budget_group[0].toUpperCase() + category.budget_group.slice(1);
    summary.getCell(row, 3).value = unallocatedRow ? formulaCell(`SUMIFS(Transactions!$F:$F,Transactions!$D:$D,A${row},Transactions!$H:$H,B${row},Transactions!$C:$C,"<>Transfer")`, amountCents / 100) : formulaCell(`SUMIFS(Transactions!$F:$F,Transactions!$D:$D,A${row},Transactions!$H:$H,B${row},Transactions!$C:$C,"<>Transfer")`, amountCents / 100);
    summary.getCell(row, 3).numFmt = pesosFormat;
    summary.getCell(row, 4).value = formulaCell(`IF($B$6=0,0,C${row}/$B$6)`, incomeCents > 0 ? amountCents / incomeCents : 0);
    summary.getCell(row, 4).numFmt = percentFormat;
  });
  summary.getCell(18 + reportCategoryRows.length, 1).value = "Note";
  summary.getCell(18 + reportCategoryRows.length, 2).value = "Category totals include only transactions categorized under each listed category; uncategorized transactions are not silently allocated.";
  summary.mergeCells(18 + reportCategoryRows.length, 2, 18 + reportCategoryRows.length, 4);
  styleTable(summary, 5, 3, 5, false);
  summary.getRow(17).font = { bold: true, color: { argb: "FFFFFFFF" } };
  summary.getRow(17).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF285A43" } };
  summary.getRow(17).alignment = { vertical: "middle", wrapText: true };

  const txSheet = workbook.addWorksheet("Transactions");
  txSheet.columns = [
    { header: "Transaction Date", key: "date", width: 18 },
    { header: "Description", key: "description", width: 34 },
    { header: "Type", key: "type", width: 24 },
    { header: "Category", key: "category", width: 24 },
    { header: "Payment Method", key: "payment", width: 20 },
    { header: "Amount (PHP)", key: "amount", width: 19 },
    { header: "Notes", key: "notes", width: 36 },
    { header: "Category Group", key: "group", width: 18 },
  ];
  txSheet.addRows(transactions.map((tx) => ({
    date: new Date(tx.transaction_date + "T00:00:00.000Z"),
    description: tx.description || "",
    type: tx.transaction_type === "savings_contribution" ? "Savings Contribution" : tx.transaction_type[0].toUpperCase() + tx.transaction_type.slice(1),
    category: tx.category_name || (tx.category_id ? categoryById.get(tx.category_id)?.name : "") || "Uncategorized",
    payment: tx.payment_method.replaceAll("_", " ").replace(/\b\w/g, (letter) => letter.toUpperCase()),
    amount: tx.amount_centavos / 100,
    notes: tx.notes || "",
    group: categoryFor(tx) ? categoryFor(tx)!.budget_group[0].toUpperCase() + categoryFor(tx)!.budget_group.slice(1) : "Unallocated",
  })));
  txSheet.getColumn("amount").numFmt = pesosFormat;
  txSheet.getColumn("date").numFmt = dateFormat;
  styleTable(txSheet, 1, 8, 1, true);
  txSheet.autoFilter = { from: "A1", to: { row: Math.max(1, txSheet.rowCount), column: 8 } };

  const budgetSheet = workbook.addWorksheet("Budget Analysis");
  budgetSheet.columns = [
    { key: "category", width: 25 },
    { key: "limit", width: 21 },
    { key: "actual", width: 23 },
    { key: "remaining", width: 24 },
    { key: "percent", width: 18 },
    { key: "status", width: 20 },
  ];
  addReportHeading(budgetSheet, "BUDGET ANALYSIS", input.range, exportDate, 6);
  budgetSheet.getCell(4, 1).value = "Budget limits are prorated by the number of overlapping calendar days in the selected reporting period.";
  budgetSheet.mergeCells(4, 1, 4, 6);
  budgetSheet.getCell(4, 1).font = { italic: true, color: { argb: "FF647367" }, size: 9 };
  budgetSheet.getRow(5).values = ["Category", "Budget Limit (PHP)", "Actual Spending (PHP)", "Remaining Budget (PHP)", "Percentage Used", "Status"];
  const applicableBudgets = input.budgets.filter((budget) => intersects(budget.period_start, budget.period_end, input.range.start, input.range.end));
  const budgetResults = applicableBudgets.map((budget) => {
    const category = categoryById.get(budget.category_id);
    const overlapStart = budget.period_start > input.range.start ? budget.period_start : input.range.start;
    const overlapEnd = budget.period_end < input.range.end ? budget.period_end : input.range.end;
    const totalDays = Math.max(1, dayNumber(budget.period_end) - dayNumber(budget.period_start) + 1);
    const overlapDays = Math.max(0, dayNumber(overlapEnd) - dayNumber(overlapStart) + 1);
    const limitCents = Math.round(budget.amount_centavos * overlapDays / totalDays);
    const actualCents = cents(transactions.filter((tx) => tx.transaction_type === "expense" && tx.category_id === budget.category_id && tx.transaction_date >= overlapStart && tx.transaction_date <= overlapEnd).map((tx) => tx.amount_centavos));
    return { budget, category, limitCents, actualCents, overlapStart, overlapEnd };
  });
  budgetResults.forEach(({ budget, category, limitCents, actualCents, overlapStart, overlapEnd }, index) => {
    const row = index + 6;
    budgetSheet.addRow({
      category: category?.name ?? budget.category_name,
      limit: limitCents / 100,
      actual: formulaCell(`SUMIFS(Transactions!$F:$F,Transactions!$D:$D,A${row},Transactions!$C:$C,"Expense",Transactions!$A:$A,">="&DATE(${Number(overlapStart.slice(0,4))},${Number(overlapStart.slice(5,7))},${Number(overlapStart.slice(8,10))}),Transactions!$A:$A,"<="&DATE(${Number(overlapEnd.slice(0,4))},${Number(overlapEnd.slice(5,7))},${Number(overlapEnd.slice(8,10))}))`, actualCents / 100),
      remaining: formulaCell(`B${row}-C${row}`, (limitCents - actualCents) / 100),
      percent: formulaCell(`IF(B${row}=0,0,C${row}/B${row})`, limitCents > 0 ? actualCents / limitCents : 0),
      status: formulaCell(`IF(C${row}>B${row},"Over Budget",IF(C${row}=B${row},"Limit Reached","Within Budget"))`, actualCents > limitCents ? "Over Budget" : actualCents === limitCents ? "Limit Reached" : "Within Budget"),
    });
  });
  budgetSheet.getColumn("limit").numFmt = pesosFormat;
  budgetSheet.getColumn("actual").numFmt = pesosFormat;
  budgetSheet.getColumn("remaining").numFmt = pesosFormat;
  budgetSheet.getColumn("percent").numFmt = percentFormat;
  styleTable(budgetSheet, 5, 6, 5, true);

  const goalsSheet = workbook.addWorksheet("Savings Goals");
  goalsSheet.columns = [
    { header: "Goal Name", key: "name", width: 30 },
    { header: "Target Amount (PHP)", key: "target", width: 22 },
    { header: "Current Saved Amount (PHP)", key: "saved", width: 26 },
    { header: "Remaining Amount (PHP)", key: "remaining", width: 24 },
    { header: "Target Date", key: "date", width: 18 },
    { header: "Completion Percentage", key: "progress", width: 23 },
  ];
  addReportHeading(goalsSheet, "SAVINGS GOALS", input.range, exportDate, 6);
  goalsSheet.insertRow(4, ["Goal balances are current totals; they are not restricted to contributions within this report period."]);
  goalsSheet.mergeCells(4, 1, 4, 6);
  goalsSheet.getRow(5).values = ["Goal Name", "Target Amount (PHP)", "Current Saved Amount (PHP)", "Remaining Amount (PHP)", "Target Date", "Completion Percentage"];
  input.goals.forEach((goal, index) => {
    const row = 6 + index;
    const saved = goal.current_amount_centavos / 100;
    const target = goal.target_amount_centavos / 100;
    goalsSheet.addRow({
      name: goal.name,
      target,
      saved,
      remaining: formulaCell(`MAX(0,B${row}-C${row})`, Math.max(0, target - saved)),
      date: goal.target_date || "",
      progress: formulaCell(`IF(B${row}=0,0,C${row}/B${row})`, target > 0 ? saved / target : 0),
    });
  });
  for (const key of ["target", "saved", "remaining"]) goalsSheet.getColumn(key).numFmt = pesosFormat;
  goalsSheet.getColumn("progress").numFmt = percentFormat;
  goalsSheet.getColumn("date").numFmt = dateFormat;
  styleTable(goalsSheet, 5, 6, 5, true);
  return workbook;
}
