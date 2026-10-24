import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { calculateUnallocatedIncome, getPeriodRange, percentageOfIncome, shiftPeriod, summarizeTransactions } from "../src/lib/finance.ts";

describe("Asia/Manila calendar date ranges", () => {
  it("returns one day for daily periods", () => assert.deepEqual(getPeriodRange("2026-02-28", "Daily"), { start: "2026-02-28", end: "2026-02-28" }));
  it("uses Monday through Sunday across year boundaries", () => assert.deepEqual(getPeriodRange("2026-01-01", "Weekly"), { start: "2025-12-29", end: "2026-01-04" }));
  it("handles leap-year February", () => assert.deepEqual(getPeriodRange("2024-02-15", "Monthly"), { start: "2024-02-01", end: "2024-02-29" }));
  it("handles December and January month navigation", () => {
    assert.equal(shiftPeriod("2026-12-31", "Monthly", 1), "2027-01-01");
    assert.equal(shiftPeriod("2027-01-01", "Monthly", -1), "2026-12-01");
  });
  it("rejects invalid calendar dates", () => assert.throws(() => getPeriodRange("2026-02-30", "Daily")));
});

describe("financial summary", () => {
  it("subtracts expenses and savings contributions once, without treating savings as expenses", () => {
    const summary = summarizeTransactions([
      { transaction_type: "income", amount_centavos: 100_000, transaction_date: "2026-07-01" },
      { transaction_type: "expense", amount_centavos: 20_000, transaction_date: "2026-07-02" },
      { transaction_type: "savings_contribution", amount_centavos: 15_000, transaction_date: "2026-07-03" },
      { transaction_type: "transfer", amount_centavos: 50_000, transaction_date: "2026-07-04" },
      { transaction_type: "income", amount_centavos: 99_000, transaction_date: "2026-08-01" },
    ], { start: "2026-07-01", end: "2026-07-31" });
    assert.deepEqual(summary, { income: 100_000, expenses: 20_000, savingsContributions: 15_000, remainingBalance: 65_000 });
  });
  it("returns zero totals for an empty range", () => assert.deepEqual(summarizeTransactions([], { start: "2026-01-01", end: "2026-01-31" }), { income: 0, expenses: 0, savingsContributions: 0, remainingBalance: 0 }));
});

describe("weekly allocation chart math", () => {
  it("calculates each slice as a percentage of total income", () => {
    assert.equal(percentageOfIncome(25_000, 100_000), 25);
    assert.equal(percentageOfIncome(15_000, 100_000), 15);
  });
  it("does not divide by zero when income is zero", () => {
    assert.equal(percentageOfIncome(5_000, 0), 0);
  });
  it("reports unallocated income or allocations above income explicitly", () => {
    assert.equal(calculateUnallocatedIncome(100_000, [20_000, 30_000, 10_000]), 40_000);
    assert.equal(calculateUnallocatedIncome(50_000, [20_000, 25_000, 10_000]), -5_000);
  });
});


import { buildFinancialWorkbook } from "../src/lib/financial-report.ts";

describe("Excel financial report workbook", () => {
  const categories = [
    { id: "salary", name: "Salary", kind: "income" as const, budget_group: "income" as const },
    { id: "food", name: "Food", kind: "expense" as const, budget_group: "needs" as const },
    { id: "dining", name: "Dining out", kind: "expense" as const, budget_group: "wants" as const },
    { id: "savings", name: "Emergency fund", kind: "expense" as const, budget_group: "savings" as const },
  ];
  const transactions = [
    { id: "1", transaction_type: "income" as const, amount_centavos: 100_000, transaction_date: "2026-10-19", category_id: "salary", description: "Pay", payment_method: "bank_transfer", notes: "", category_name: "Salary" },
    { id: "2", transaction_type: "expense" as const, amount_centavos: 20_000, transaction_date: "2026-10-20", category_id: "food", description: "Groceries", payment_method: "cash", notes: "Weekly shop", category_name: "Food" },
    { id: "3", transaction_type: "expense" as const, amount_centavos: 5_000, transaction_date: "2026-10-21", category_id: "dining", description: "Lunch", payment_method: "e_wallet", notes: "", category_name: "Dining out" },
    { id: "4", transaction_type: "savings_contribution" as const, amount_centavos: 15_000, transaction_date: "2026-10-22", category_id: "savings", description: "Emergency fund", payment_method: "bank_transfer", notes: "", category_name: "Emergency fund" },
    { id: "5", transaction_type: "transfer" as const, amount_centavos: 9_000, transaction_date: "2026-10-23", category_id: null, description: "Own account transfer", payment_method: "bank_transfer", notes: "", category_name: null },
    { id: "6", transaction_type: "income" as const, amount_centavos: 999_000, transaction_date: "2026-10-26", category_id: "salary", description: "Outside period", payment_method: "bank_transfer", notes: "", category_name: "Salary" },
  ];
  const input = {
    period: "Weekly" as const,
    range: { start: "2026-10-19", end: "2026-10-25" },
    transactions,
    categories,
    budgets: [{ id: "b1", category_id: "food", category_name: "Food", period_start: "2026-10-19", period_end: "2026-10-25", amount_centavos: 30_000 }],
    goals: [{ id: "g1", name: "Emergency fund", description: "Rainy day", target_amount_centavos: 50_000, current_amount_centavos: 25_000, target_date: "2027-01-01", status: "active" }],
    exportedAt: new Date("2026-10-24T05:00:00Z"),
  };

  it("creates the four required worksheets, includes only selected dates, and keeps transfers visible but out of totals", async () => {
    const workbook = buildFinancialWorkbook(input);
    assert.deepEqual(workbook.worksheets.map((sheet) => sheet.name), ["Financial Summary", "Transactions", "Budget Analysis", "Savings Goals"]);
    const txSheet = workbook.getWorksheet("Transactions")!;
    assert.equal(txSheet.rowCount, 6);
    assert.equal(txSheet.getCell("B6").value, "Own account transfer");
    assert.equal(txSheet.getCell("B7").value, null);
    assert.equal(txSheet.views[0].state, "frozen");
    assert.equal(txSheet.views[0].ySplit, 1);
    assert.ok(txSheet.autoFilter);
    assert.equal(txSheet.getColumn(9).hidden, true);
    assert.equal(txSheet.getCell("A2").numFmt, "yyyy-mm-dd");
    assert.equal(txSheet.getCell("F3").numFmt, '"₱"#,##0.00;[Red]-"₱"#,##0.00');
    const summary = workbook.getWorksheet("Financial Summary")!;
    assert.equal((summary.getCell("B6").value as { result: number }).result, 1000);
    assert.equal((summary.getCell("B7").value as { result: number }).result, 250);
    assert.equal((summary.getCell("B8").value as { result: number }).result, 150);
    assert.equal((summary.getCell("B9").value as { result: number }).result, 600);
    assert.equal((summary.getCell("B10").value as { result: number }).result, 200);
    assert.equal((summary.getCell("B11").value as { result: number }).result, 50);
    assert.match((summary.getCell("B9").value as { formula: string }).formula, /B6-B7-B8/);
    assert.match((summary.getCell("C19").value as { formula: string }).formula, /Transactions!\\$I:\\$I/);
    const budget = workbook.getWorksheet("Budget Analysis")!;
    assert.equal(budget.getCell("C6").value && (budget.getCell("C6").value as { result: number }).result, 200);
    assert.equal((budget.getCell("D6").value as { result: number }).result, 100);
    assert.equal((budget.getCell("F6").value as { result: string }).result, "Within Budget");
    assert.match((budget.getCell("C6").value as { formula: string }).formula, /Transactions!\\$I:\\$I/);
    const goal = workbook.getWorksheet("Savings Goals")!;
    assert.equal((goal.getCell("D6").value as { result: number }).result, 250);
    assert.equal((goal.getCell("F6").value as { result: number }).result, 0.5);
    assert.equal(goal.getCell("F6").numFmt, "0.0%");
  });

  it("serializes and reloads as a genuine xlsx workbook with formulas and formatting intact", async () => {
    const workbook = buildFinancialWorkbook(input);
    const bytes = await workbook.xlsx.writeBuffer();
    assert.ok(bytes.byteLength > 2_000);
    const reopened = new (await import("exceljs")).default.Workbook();
    await reopened.xlsx.load(bytes as ArrayBuffer);
    assert.deepEqual(reopened.worksheets.map((sheet) => sheet.name), ["Financial Summary", "Transactions", "Budget Analysis", "Savings Goals"]);
    assert.match((reopened.getWorksheet("Financial Summary")!.getCell("B9").value as { formula: string }).formula, /B6-B7-B8/);
    assert.equal(reopened.getWorksheet("Transactions")!.getCell("A2").value instanceof Date, true);
  });

  it("handles an empty period and rejects impossible report dates", () => {
    const workbook = buildFinancialWorkbook({ ...input, range: { start: "2026-11-01", end: "2026-11-30" } });
    assert.equal(workbook.getWorksheet("Transactions")!.rowCount, 1);
    assert.equal(((workbook.getWorksheet("Financial Summary")!.getCell("B6").value as { result?: number }).result ?? 0), 0);
    assert.match((workbook.getWorksheet("Financial Summary")!.getCell("B6").value as { formula: string }).formula, /SUMIF/);
    assert.throws(() => buildFinancialWorkbook({ ...input, range: { start: "2026-02-30", end: "2026-03-01" } }), /valid reporting date range/i);
  });
});
