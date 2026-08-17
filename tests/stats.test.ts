import { describe, expect, it } from 'vitest';
import { allocateSharedExpense, computeDashboard, firmOf, growthPercent } from '../src/shared/stats.js';
import type { StatsContext } from '../src/shared/stats.js';
import type { Account, Category, Firm, Point, Transaction } from '../src/shared/types.js';

const firmA: Firm = { id: 'f1', name: 'Firma A', active: true, expenseAllocation: 'proRataIncome', sortOrder: 1 };
const firmB: Firm = { id: 'f2', name: 'Firma B', active: true, expenseAllocation: 'proRataIncome', sortOrder: 2 };

const pointA: Point = { id: 'p1', name: 'Do‘kon A', firmId: 'f1', active: true, excludeFromRevenue: false, sortOrder: 1 };
const pointB: Point = { id: 'p2', name: 'Do‘kon B', firmId: 'f2', active: true, excludeFromRevenue: false, sortOrder: 2 };
const pointExcluded: Point = {
  id: 'p3',
  name: 'Direktor kassasi',
  firmId: 'f1',
  active: true,
  excludeFromRevenue: true,
  sortOrder: 3,
};

const category: Category = { id: 'c1', name: 'Ijara', kind: 'expense', active: true, sortOrder: 1 };
const account: Account = { id: 'a1', name: 'Naqd', currency: 'UZS', active: true, showInBalance: true, sortOrder: 1 };

const context: StatsContext = {
  firms: [firmA, firmB],
  points: [pointA, pointB, pointExcluded],
  categories: [category],
  accounts: [account],
  baseCurrency: 'UZS',
  sharedExpenseAllocation: 'proRataIncome',
};

let counter = 0;
function tx(partial: Partial<Transaction> & Pick<Transaction, 'kind' | 'baseAmountMinor'>): Transaction {
  counter += 1;
  return {
    id: `t${counter}`,
    batchId: `b${counter}`,
    date: '2026-08-10',
    firmId: null,
    pointId: null,
    categoryId: null,
    accountId: 'a1',
    amountMinor: partial.baseAmountMinor,
    currency: 'UZS',
    counterAccountId: null,
    counterAmountMinor: null,
    counterCurrency: null,
    fxRate: 1,
    note: '',
    createdAt: '2026-08-10T00:00:00.000Z',
    createdBy: 'u1',
    updatedAt: '2026-08-10T00:00:00.000Z',
    updatedBy: 'u1',
    deletedAt: null,
    deletedBy: null,
    ...partial,
  };
}

const range = { start: '2026-08-01', end: '2026-08-31' };
const previousRange = { start: '2026-07-01', end: '2026-07-31' };

describe('firmOf', () => {
  it('firmani savdo nuqtasi orqali aniqlaydi', () => {
    const map = new Map([[pointA.id, pointA]]);
    expect(firmOf(tx({ kind: 'income', baseAmountMinor: 100, pointId: 'p1' }), map)).toBe('f1');
  });

  it('nuqta bo‘lmasa bevosita firmani oladi', () => {
    expect(firmOf(tx({ kind: 'expense', baseAmountMinor: 100, firmId: 'f2' }), new Map())).toBe('f2');
  });

  it('ikkalasi ham bo‘lmasa null', () => {
    expect(firmOf(tx({ kind: 'expense', baseAmountMinor: 100 }), new Map())).toBeNull();
  });
});

describe('computeDashboard', () => {
  it('kirim va chiqimni yig‘adi, o‘tkazmani chiqarib tashlaydi', () => {
    const stats = computeDashboard(
      [
        tx({ kind: 'income', baseAmountMinor: 100_000, pointId: 'p1' }),
        tx({ kind: 'expense', baseAmountMinor: 30_000, pointId: 'p1', categoryId: 'c1' }),
        tx({ kind: 'transfer', baseAmountMinor: 999_000 }),
      ],
      context,
      { range, firmId: null },
      previousRange,
    );

    expect(stats.current.incomeMinor).toBe(100_000);
    expect(stats.current.expenseMinor).toBe(30_000);
    expect(stats.current.profitMinor).toBe(70_000);
  });

  it('excludeFromRevenue belgilangan nuqtani daromaddan chiqaradi', () => {
    // v1 da bu ro'yxat kodga qotirib yozilgan edi ("DONIYOR AKA", "KASSA"...).
    const stats = computeDashboard(
      [
        tx({ kind: 'income', baseAmountMinor: 100_000, pointId: 'p1' }),
        tx({ kind: 'income', baseAmountMinor: 500_000, pointId: 'p3' }),
      ],
      context,
      { range, firmId: null },
      previousRange,
    );
    expect(stats.current.incomeMinor).toBe(100_000);
  });

  it('firma bo‘yicha filtrlaydi', () => {
    const stats = computeDashboard(
      [
        tx({ kind: 'income', baseAmountMinor: 100_000, pointId: 'p1' }),
        tx({ kind: 'income', baseAmountMinor: 400_000, pointId: 'p2' }),
      ],
      context,
      { range, firmId: 'f2' },
      previousRange,
    );
    expect(stats.current.incomeMinor).toBe(400_000);
  });

  it('umumiy xarajatni daromadga proporsional taqsimlaydi', () => {
    const transactions = [
      tx({ kind: 'income', baseAmountMinor: 300_000, pointId: 'p1' }), // Firma A
      tx({ kind: 'income', baseAmountMinor: 100_000, pointId: 'p2' }), // Firma B
      tx({ kind: 'expense', baseAmountMinor: 40_000, categoryId: 'c1' }), // umumiy
    ];

    const forA = computeDashboard(transactions, context, { range, firmId: 'f1' }, previousRange);
    const forB = computeDashboard(transactions, context, { range, firmId: 'f2' }, previousRange);

    // 3:1 nisbatda → 30 000 va 10 000
    expect(forA.current.expenseMinor).toBe(30_000);
    expect(forB.current.expenseMinor).toBe(10_000);
  });

  it('taqsimlash o‘chirilganda umumiy xarajat firmaga tushmaydi', () => {
    const directContext: StatsContext = { ...context, sharedExpenseAllocation: 'direct' };
    const stats = computeDashboard(
      [
        tx({ kind: 'income', baseAmountMinor: 100_000, pointId: 'p1' }),
        tx({ kind: 'expense', baseAmountMinor: 40_000, categoryId: 'c1' }),
      ],
      directContext,
      { range, firmId: 'f1' },
      previousRange,
    );
    expect(stats.current.expenseMinor).toBe(0);
  });

  it('davrdan tashqaridagi yozuvlarni hisobga olmaydi', () => {
    const stats = computeDashboard(
      [
        tx({ kind: 'income', baseAmountMinor: 100_000, pointId: 'p1', date: '2026-08-10' }),
        tx({ kind: 'income', baseAmountMinor: 700_000, pointId: 'p1', date: '2026-09-10' }),
      ],
      context,
      { range, firmId: null },
      previousRange,
    );
    expect(stats.current.incomeMinor).toBe(100_000);
  });

  it('oldingi davrni alohida hisoblaydi', () => {
    const stats = computeDashboard(
      [
        tx({ kind: 'income', baseAmountMinor: 100_000, pointId: 'p1', date: '2026-08-10' }),
        tx({ kind: 'income', baseAmountMinor: 80_000, pointId: 'p1', date: '2026-07-10' }),
      ],
      context,
      { range, firmId: null },
      previousRange,
    );
    expect(stats.current.incomeMinor).toBe(100_000);
    expect(stats.previous.incomeMinor).toBe(80_000);
  });

  it('kesimlarni tayyorlaydi', () => {
    const stats = computeDashboard(
      [
        tx({ kind: 'income', baseAmountMinor: 100_000, pointId: 'p1' }),
        tx({ kind: 'expense', baseAmountMinor: 20_000, pointId: 'p1', categoryId: 'c1' }),
      ],
      context,
      { range, firmId: null },
      previousRange,
    );
    expect(stats.byPoint[0]).toMatchObject({ name: 'Do‘kon A', totalMinor: 100_000 });
    expect(stats.byCategory[0]).toMatchObject({ name: 'Ijara', totalMinor: 20_000 });
    expect(stats.byAccount[0]).toMatchObject({ name: 'Naqd', totalMinor: 100_000 });
  });

  it('o‘chirilgan yozuvni hisobga olmaydi', () => {
    const stats = computeDashboard(
      [tx({ kind: 'income', baseAmountMinor: 100_000, pointId: 'p1', deletedAt: '2026-08-12T00:00:00Z' })],
      context,
      { range, firmId: null },
      previousRange,
    );
    expect(stats.current.incomeMinor).toBe(0);
  });
});

describe('allocateSharedExpense', () => {
  it('yaxlitlash qoldig‘ini yo‘qotmaydi', () => {
    const income = new Map([
      ['f1', 1],
      ['f2', 1],
      ['f3', 1],
    ]);
    const parts = allocateSharedExpense(100, ['f1', 'f2', 'f3'], income, 'proRataIncome');
    const total = Array.from(parts.values()).reduce((sum, value) => sum + value, 0);
    expect(total).toBe(100);
  });

  it('teng taqsimlash rejimida ulushlarni tenglashtiradi', () => {
    const parts = allocateSharedExpense(100, ['f1', 'f2'], new Map([['f1', 900]]), 'equal');
    expect(parts.get('f1')).toBe(50);
    expect(parts.get('f2')).toBe(50);
  });

  it('`direct` rejimida hech narsa taqsimlanmaydi', () => {
    const parts = allocateSharedExpense(100, ['f1'], new Map(), 'direct');
    expect(parts.size).toBe(0);
  });
});

describe('growthPercent', () => {
  it('o‘sishni foizda beradi', () => {
    expect(growthPercent(150, 100)).toBe(50);
    expect(growthPercent(50, 100)).toBe(-50);
  });

  it('oldingi davr nol bo‘lganda nolga bo‘lmaydi', () => {
    expect(growthPercent(100, 0)).toBeNull();
    expect(growthPercent(0, 0)).toBe(0);
  });
});
