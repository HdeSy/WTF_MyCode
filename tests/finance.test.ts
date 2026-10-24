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
