export type TransactionKind = "income" | "expense" | "savings_contribution" | "transfer";
export type PeriodKind = "Daily" | "Weekly" | "Monthly" | "Custom";
export type FinancialTransaction = { transaction_type: TransactionKind; amount_centavos: number; transaction_date: string };

export type DateRange = { start: string; end: string };
export type FinancialSummary = { income: number; expenses: number; savingsContributions: number; remainingBalance: number };

function parseDate(value: string): Date {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) throw new Error("Date must use YYYY-MM-DD.");
  const date = new Date(value + "T00:00:00.000Z");
  if (!Number.isFinite(date.getTime()) || date.toISOString().slice(0, 10) !== value) throw new Error("Invalid calendar date.");
  return date;
}
function isoDate(date: Date): string { return date.toISOString().slice(0, 10); }

/** Date-only values are calendar dates in Asia/Manila; UTC is used only for safe calendar arithmetic. */
export function getPeriodRange(anchor: string, period: Exclude<PeriodKind, "Custom">): DateRange {
  const date = parseDate(anchor);
  if (period === "Daily") return { start: anchor, end: anchor };
  if (period === "Weekly") {
    const mondayOffset = (date.getUTCDay() + 6) % 7;
    const start = new Date(date);
    start.setUTCDate(start.getUTCDate() - mondayOffset);
    const end = new Date(start);
    end.setUTCDate(end.getUTCDate() + 6);
    return { start: isoDate(start), end: isoDate(end) };
  }
  const start = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), 1));
  const end = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + 1, 0));
  return { start: isoDate(start), end: isoDate(end) };
}

export function shiftPeriod(anchor: string, period: Exclude<PeriodKind, "Custom">, direction: -1 | 1): string {
  const date = parseDate(anchor);
  if (period === "Daily") date.setUTCDate(date.getUTCDate() + direction);
  else if (period === "Weekly") date.setUTCDate(date.getUTCDate() + 7 * direction);
  else date.setUTCMonth(date.getUTCMonth() + direction, 1);
  return isoDate(date);
}

export function summarizeTransactions<T extends FinancialTransaction>(transactions: T[], range: DateRange): FinancialSummary {
  let income = 0, expenses = 0, savingsContributions = 0;
  for (const tx of transactions) {
    if (tx.transaction_date < range.start || tx.transaction_date > range.end) continue;
    if (tx.transaction_type === "income") income += tx.amount_centavos;
    else if (tx.transaction_type === "expense") expenses += tx.amount_centavos;
    else if (tx.transaction_type === "savings_contribution") savingsContributions += tx.amount_centavos;
  }
  return { income, expenses, savingsContributions, remainingBalance: income - expenses - savingsContributions };
}

/** Percentages for the allocation chart are always measured against total income, never against the sum of slices. */
export function percentageOfIncome(amountCentavos: number, incomeCentavos: number): number {
  if (!Number.isFinite(amountCentavos) || !Number.isFinite(incomeCentavos) || incomeCentavos <= 0) return 0;
  return amountCentavos / incomeCentavos * 100;
}

/** A negative result means the three allocations exceed income. Keep it visible instead of hiding it in a slice. */
export function unallocatedIncome(incomeCentavos: number, allocationsCentavos: number[]): number {
  const allocated = allocationsCentavos.reduce((sum, amount) => sum + (Number.isFinite(amount) ? amount : 0), 0);
  return incomeCentavos - allocated;
}
