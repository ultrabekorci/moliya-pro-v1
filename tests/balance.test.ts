import { describe, expect, it } from 'vitest';
import { computeBalances, totalInBase, validateTransferShape } from '../src/shared/balance.js';
import type { Account, Transaction } from '../src/shared/types.js';

function account(id: string, name: string, currency = 'UZS', sortOrder = 0): Account {
  return { id, name, currency, active: true, showInBalance: true, sortOrder };
}

let counter = 0;
function tx(partial: Partial<Transaction> & Pick<Transaction, 'kind' | 'accountId' | 'amountMinor'>): Transaction {
  counter += 1;
  return {
    id: `tx${counter}`,
    batchId: `b${counter}`,
    date: '2026-08-10',
    firmId: null,
    pointId: null,
    categoryId: null,
    currency: 'UZS',
    counterAccountId: null,
    counterAmountMinor: null,
    counterCurrency: null,
    fxRate: 1,
    baseAmountMinor: partial.amountMinor,
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

const options = { baseCurrency: 'UZS', rates: { UZS: 1, USD: 12_500 } };

describe('computeBalances', () => {
  it('kirimni qo‘shadi, chiqimni ayiradi', () => {
    const accounts = [account('cash', 'Naqd')];
    const balances = computeBalances(
      accounts,
      [
        tx({ kind: 'income', accountId: 'cash', amountMinor: 100_000 }),
        tx({ kind: 'expense', accountId: 'cash', amountMinor: 30_000 }),
      ],
      options,
    );
    expect(balances[0]?.balanceMinor).toBe(70_000);
  });

  it('IXTIYORIY to‘lov turini hisobga oladi', () => {
    // v1 da balans faqat {Naqd, P2P, Dollar, Bank} ni bilardi va Uzcard/Humo
    // tushumlari jimgina yo'qolardi. Bu test aynan o'sha xatoni qo'riqlaydi.
    const accounts = [account('uzcard', 'Uzcard'), account('humo', 'Humo')];
    const balances = computeBalances(
      accounts,
      [
        tx({ kind: 'income', accountId: 'uzcard', amountMinor: 50_000 }),
        tx({ kind: 'income', accountId: 'humo', amountMinor: 25_000 }),
      ],
      options,
    );
    expect(balances.find((b) => b.accountId === 'uzcard')?.balanceMinor).toBe(50_000);
    expect(balances.find((b) => b.accountId === 'humo')?.balanceMinor).toBe(25_000);
  });

  it('o‘tkazmada ikkala tomonni ham to‘g‘ri o‘zgartiradi', () => {
    const accounts = [account('cash', 'Naqd'), account('bank', 'Bank')];
    const balances = computeBalances(
      accounts,
      [
        tx({ kind: 'income', accountId: 'cash', amountMinor: 100_000 }),
        tx({
          kind: 'transfer',
          accountId: 'cash',
          amountMinor: 40_000,
          counterAccountId: 'bank',
          counterAmountMinor: 40_000,
          counterCurrency: 'UZS',
        }),
      ],
      options,
    );
    expect(balances.find((b) => b.accountId === 'cash')?.balanceMinor).toBe(60_000);
    expect(balances.find((b) => b.accountId === 'bank')?.balanceMinor).toBe(40_000);
  });

  it('turli valyutadagi o‘tkazmada har bir tomon o‘z valyutasida o‘zgaradi', () => {
    const accounts = [account('cash', 'Naqd', 'UZS'), account('usd', 'Dollar kassa', 'USD')];
    const balances = computeBalances(
      accounts,
      [
        tx({ kind: 'income', accountId: 'cash', amountMinor: 125_000_000 }), // 1 250 000 UZS
        tx({
          kind: 'transfer',
          accountId: 'cash',
          amountMinor: 125_000_000,
          counterAccountId: 'usd',
          counterAmountMinor: 10_000, // 100 USD
          counterCurrency: 'USD',
        }),
      ],
      options,
    );
    expect(balances.find((b) => b.accountId === 'cash')?.balanceMinor).toBe(0);
    expect(balances.find((b) => b.accountId === 'usd')?.balanceMinor).toBe(10_000);
  });

  it('dollar chiqimi USD da ayiriladi, izohdan qidirilmaydi', () => {
    // v1 da USD chiqimi UZS ga o'girib saqlanar va asl summa izohdan
    // regexp bilan «ajratib olinardi». Endi summa hisobning o'z valyutasida.
    const accounts = [account('usd', 'Dollar kassa', 'USD')];
    const balances = computeBalances(
      accounts,
      [
        tx({ kind: 'income', accountId: 'usd', amountMinor: 100_000, currency: 'USD' }),
        tx({ kind: 'expense', accountId: 'usd', amountMinor: 25_000, currency: 'USD', note: '' }),
      ],
      options,
    );
    expect(balances[0]?.balanceMinor).toBe(75_000);
    expect(balances[0]?.currency).toBe('USD');
  });

  it('o‘chirilgan yozuvlarni hisobga olmaydi', () => {
    const accounts = [account('cash', 'Naqd')];
    const balances = computeBalances(
      accounts,
      [
        tx({ kind: 'income', accountId: 'cash', amountMinor: 100_000 }),
        tx({ kind: 'income', accountId: 'cash', amountMinor: 999_000, deletedAt: '2026-08-11T00:00:00Z' }),
      ],
      options,
    );
    expect(balances[0]?.balanceMinor).toBe(100_000);
  });

  it('asOf sanasidan keyingi yozuvlarni kesib tashlaydi', () => {
    const accounts = [account('cash', 'Naqd')];
    const balances = computeBalances(
      accounts,
      [
        tx({ kind: 'income', accountId: 'cash', amountMinor: 100_000, date: '2026-08-01' }),
        tx({ kind: 'income', accountId: 'cash', amountMinor: 50_000, date: '2026-08-20' }),
      ],
      { ...options, asOf: '2026-08-10' },
    );
    expect(balances[0]?.balanceMinor).toBe(100_000);
  });

  it('mavjud bo‘lmagan hisobga tegishli yozuvni e’tiborsiz qoldiradi', () => {
    const balances = computeBalances(
      [account('cash', 'Naqd')],
      [tx({ kind: 'income', accountId: 'yoq', amountMinor: 1000 })],
      options,
    );
    expect(balances[0]?.balanceMinor).toBe(0);
  });

  it('hisobot valyutasidagi ekvivalentni hisoblaydi', () => {
    const accounts = [account('usd', 'Dollar kassa', 'USD')];
    const balances = computeBalances(
      accounts,
      [tx({ kind: 'income', accountId: 'usd', amountMinor: 10_000, currency: 'USD' })],
      options,
    );
    // 100 USD × 12 500 = 1 250 000 UZS → minor: 125 000 000
    expect(balances[0]?.baseBalanceMinor).toBe(125_000_000);
    expect(totalInBase(balances)).toBe(125_000_000);
  });
});

describe('validateTransferShape', () => {
  it('bir xil hisoblar orasidagi o‘tkazmani rad etadi', () => {
    expect(validateTransferShape({ kind: 'transfer', accountId: 'a', counterAccountId: 'a' })).toMatch(/bir xil/i);
  });

  it('ikkinchi hisob ko‘rsatilmasa xato beradi', () => {
    expect(validateTransferShape({ kind: 'transfer', accountId: 'a', counterAccountId: null })).toBeTruthy();
  });

  it('to‘g‘ri o‘tkazmada null qaytaradi', () => {
    expect(validateTransferShape({ kind: 'transfer', accountId: 'a', counterAccountId: 'b' })).toBeNull();
  });
});
