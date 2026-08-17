/**
 * Dashboard statistikasi — sof funksiyalar (test bilan qoplangan).
 *
 * v1 da firma nomi savdo nuqtasi matni ichidan qidirilardi
 * (`pointName.includes("SMARTMIZ")`) va "e'tiborga olinmaydigan" nuqtalar
 * ro'yxati kodga qotirib yozilgan edi. Endi hammasi katalogdan keladi:
 *   - nuqta → firma bog'lanishi `Point.firmId` orqali,
 *   - daromaddan chiqarib tashlash `Point.excludeFromRevenue` bayrog'i orqali,
 *   - umumiy xarajatlarni taqsimlash usuli `AppConfig.sharedExpenseAllocation` orqali.
 */

import { allocateMinor } from './money.js';
import { shouldGroupMonthly, trendKey } from './dates.js';
import type {
  Account,
  Category,
  DashboardStats,
  DateRange,
  ExpenseAllocation,
  Firm,
  MinorUnits,
  NamedTotal,
  PeriodTotals,
  Point,
  Transaction,
  TrendPoint,
} from './types.js';

export interface StatsContext {
  firms: readonly Firm[];
  points: readonly Point[];
  categories: readonly Category[];
  accounts: readonly Account[];
  baseCurrency: string;
  sharedExpenseAllocation: ExpenseAllocation;
}

export interface StatsQuery {
  range: DateRange;
  /** `null` — barcha firmalar bo'yicha umumiy. */
  firmId: string | null;
}

interface Indexed {
  pointById: Map<string, Point>;
  firmById: Map<string, Firm>;
  categoryById: Map<string, Category>;
  accountById: Map<string, Account>;
}

function index(context: StatsContext): Indexed {
  return {
    pointById: new Map(context.points.map((p) => [p.id, p])),
    firmById: new Map(context.firms.map((f) => [f.id, f])),
    categoryById: new Map(context.categories.map((c) => [c.id, c])),
    accountById: new Map(context.accounts.map((a) => [a.id, a])),
  };
}

/** Tranzaksiya qaysi firmaga tegishli: avval nuqta orqali, keyin bevosita maydon orqali. */
export function firmOf(tx: Transaction, pointById: ReadonlyMap<string, Point>): string | null {
  if (tx.pointId) {
    const point = pointById.get(tx.pointId);
    if (point?.firmId) return point.firmId;
  }
  return tx.firmId ?? null;
}

function countsAsRevenue(tx: Transaction, pointById: ReadonlyMap<string, Point>): boolean {
  if (tx.kind !== 'income') return false;
  if (!tx.pointId) return true;
  const point = pointById.get(tx.pointId);
  return point ? !point.excludeFromRevenue : true;
}

function bump(map: Map<string, MinorUnits>, key: string, amount: MinorUnits): void {
  map.set(key, (map.get(key) ?? 0) + amount);
}

function toNamedTotals(
  totals: ReadonlyMap<string, MinorUnits>,
  nameOf: (id: string) => string,
): NamedTotal[] {
  return Array.from(totals.entries())
    .map(([id, totalMinor]) => ({ id, name: nameOf(id), totalMinor }))
    .filter((entry) => entry.totalMinor !== 0)
    .sort((a, b) => b.totalMinor - a.totalMinor);
}

/**
 * Umumiy (firmaga bevosita bog'lanmagan) xarajatlarni firmalar bo'yicha taqsimlaydi.
 * Qaytadigan qiymatlar yig'indisi har doim `totalShared` ga teng (yaxlitlash qoldig'i yo'qolmaydi).
 */
export function allocateSharedExpense(
  totalShared: MinorUnits,
  firmIds: readonly string[],
  incomeByFirm: ReadonlyMap<string, MinorUnits>,
  mode: ExpenseAllocation,
): Map<string, MinorUnits> {
  const result = new Map<string, MinorUnits>();
  if (firmIds.length === 0 || totalShared === 0 || mode === 'direct') return result;

  const weights =
    mode === 'equal' ? firmIds.map(() => 1) : firmIds.map((id) => Math.max(incomeByFirm.get(id) ?? 0, 0));

  const hasWeight = weights.some((w) => w > 0);
  const parts = allocateMinor(totalShared, hasWeight ? weights : firmIds.map(() => 1));
  firmIds.forEach((id, i) => result.set(id, parts[i] ?? 0));
  return result;
}

interface Aggregate {
  totals: PeriodTotals;
  trend: Map<string, { incomeMinor: MinorUnits; expenseMinor: MinorUnits }>;
  byFirm: Map<string, MinorUnits>;
  byPoint: Map<string, MinorUnits>;
  byCategory: Map<string, MinorUnits>;
  byAccount: Map<string, MinorUnits>;
  count: number;
}

function aggregate(
  transactions: readonly Transaction[],
  context: StatsContext,
  query: StatsQuery,
  idx: Indexed,
  monthly: boolean,
): Aggregate {
  const inRange = transactions.filter(
    (tx) => !tx.deletedAt && tx.date >= query.range.start && tx.date <= query.range.end,
  );

  // 1-bosqich: daromadni firmalar bo'yicha yig'amiz (taqsimlash ulushlari uchun kerak).
  const incomeByFirm = new Map<string, MinorUnits>();
  for (const tx of inRange) {
    if (!countsAsRevenue(tx, idx.pointById)) continue;
    const firmId = firmOf(tx, idx.pointById);
    if (firmId) bump(incomeByFirm, firmId, tx.baseAmountMinor);
  }

  // 2-bosqich: umumiy xarajatlar summasi va ularning taqsimoti.
  const activeFirmIds = context.firms.filter((f) => f.active).map((f) => f.id);
  let sharedExpenseTotal = 0;
  for (const tx of inRange) {
    if (tx.kind !== 'expense') continue;
    if (firmOf(tx, idx.pointById) === null) sharedExpenseTotal += tx.baseAmountMinor;
  }
  const sharedShare = allocateSharedExpense(
    sharedExpenseTotal,
    activeFirmIds,
    incomeByFirm,
    context.sharedExpenseAllocation,
  );
  const sharedRatio =
    sharedExpenseTotal > 0 ? (sharedShare.get(query.firmId ?? '') ?? 0) / sharedExpenseTotal : 0;

  const result: Aggregate = {
    totals: { incomeMinor: 0, expenseMinor: 0, profitMinor: 0 },
    trend: new Map(),
    byFirm: new Map(),
    byPoint: new Map(),
    byCategory: new Map(),
    byAccount: new Map(),
    count: 0,
  };

  const trendBump = (key: string, income: MinorUnits, expense: MinorUnits): void => {
    const entry = result.trend.get(key) ?? { incomeMinor: 0, expenseMinor: 0 };
    entry.incomeMinor += income;
    entry.expenseMinor += expense;
    result.trend.set(key, entry);
  };

  for (const tx of inRange) {
    if (tx.kind === 'transfer') continue; // O'tkazma na daromad, na xarajat.

    const firmId = firmOf(tx, idx.pointById);
    const key = trendKey(tx.date, monthly);

    if (tx.kind === 'income') {
      if (!countsAsRevenue(tx, idx.pointById)) continue;
      if (query.firmId !== null && firmId !== query.firmId) continue;

      result.totals.incomeMinor += tx.baseAmountMinor;
      result.count += 1;
      trendBump(key, tx.baseAmountMinor, 0);
      if (firmId) bump(result.byFirm, firmId, tx.baseAmountMinor);
      if (tx.pointId) bump(result.byPoint, tx.pointId, tx.baseAmountMinor);
      bump(result.byAccount, tx.accountId, tx.baseAmountMinor);
      continue;
    }

    // Xarajat
    let amount: MinorUnits;
    if (firmId === null) {
      if (query.firmId === null) {
        amount = tx.baseAmountMinor;
      } else {
        // Umumiy xarajatning shu firmaga to'g'ri keladigan ulushi.
        amount = Math.round(tx.baseAmountMinor * sharedRatio);
        if (amount === 0) continue;
      }
    } else {
      if (query.firmId !== null && firmId !== query.firmId) continue;
      amount = tx.baseAmountMinor;
    }

    result.totals.expenseMinor += amount;
    result.count += 1;
    trendBump(key, 0, amount);
    if (firmId) bump(result.byFirm, firmId, amount);
    if (tx.categoryId) bump(result.byCategory, tx.categoryId, amount);
  }

  result.totals.profitMinor = result.totals.incomeMinor - result.totals.expenseMinor;
  return result;
}

function trendToArray(trend: ReadonlyMap<string, { incomeMinor: MinorUnits; expenseMinor: MinorUnits }>): TrendPoint[] {
  return Array.from(trend.entries())
    .map(([key, value]) => ({ key, incomeMinor: value.incomeMinor, expenseMinor: value.expenseMinor }))
    .sort((a, b) => (a.key < b.key ? -1 : a.key > b.key ? 1 : 0));
}

export function computeDashboard(
  transactions: readonly Transaction[],
  context: StatsContext,
  query: StatsQuery,
  previousRange: DateRange,
): DashboardStats {
  const idx = index(context);
  const monthly = shouldGroupMonthly(query.range);

  const current = aggregate(transactions, context, query, idx, monthly);
  const previous = aggregate(
    transactions,
    context,
    { range: previousRange, firmId: query.firmId },
    idx,
    monthly,
  );

  const nameFrom = <T extends { name: string }>(map: ReadonlyMap<string, T>) => (id: string): string =>
    map.get(id)?.name ?? '—';

  return {
    baseCurrency: context.baseCurrency,
    current: current.totals,
    previous: previous.totals,
    trend: trendToArray(current.trend),
    byFirm: toNamedTotals(current.byFirm, nameFrom(idx.firmById)),
    byPoint: toNamedTotals(current.byPoint, nameFrom(idx.pointById)),
    byCategory: toNamedTotals(current.byCategory, nameFrom(idx.categoryById)),
    byAccount: toNamedTotals(current.byAccount, nameFrom(idx.accountById)),
    transactionCount: current.count,
  };
}

/** O'sish foizi. Oldingi davr nolga teng bo'lsa `null` (bo'lish xatosi o'rniga). */
export function growthPercent(current: MinorUnits, previous: MinorUnits): number | null {
  if (previous === 0) return current === 0 ? 0 : null;
  return ((current - previous) / Math.abs(previous)) * 100;
}
