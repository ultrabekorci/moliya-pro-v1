/**
 * Balans hisoblash — sof (pure) funksiyalar. Server ham, klient ham shu bitta
 * implementatsiyadan foydalanadi, shuning uchun ikkalasi hech qachon farq qilmaydi.
 *
 * v1 dagi xatolar shu yerda tuzatilgan:
 *  - `{Naqd, P2P, Dollar, Bank}` qotirib yozilgan ro'yxat o'rniga IXTIYORIY hisoblar.
 *    Uzcard/Humo/Perechislenie tushumlari endi balansdan tushib qolmaydi.
 *  - Dollar chiqimi izohdan regexp bilan "ajratib olinmaydi" — summa hisobning
 *    o'z valyutasida saqlangan.
 *  - O'chirilgan yozuvlar `deletedAt` bo'yicha chetlab o'tiladi.
 */

import { convertMinor } from './money.js';
import type {
  Account,
  AccountBalance,
  CurrencyCode,
  IsoDate,
  MinorUnits,
  Transaction,
} from './types.js';

export interface BalanceOptions {
  /** Shu sanagacha (shu kun ham kiradi) bo'lgan holat. Berilmasa — hammasi. */
  asOf?: IsoDate;
  /** Hisobot valyutasi. */
  baseCurrency: CurrencyCode;
  /** Valyuta → base kursi. `USD: 12500` → 1 USD = 12500 base birlik. */
  rates: Readonly<Record<string, number>>;
}

function rateFor(currency: CurrencyCode, baseCurrency: CurrencyCode, rates: Readonly<Record<string, number>>): number {
  if (currency.toUpperCase() === baseCurrency.toUpperCase()) return 1;
  const rate = rates[currency.toUpperCase()];
  return typeof rate === 'number' && Number.isFinite(rate) && rate > 0 ? rate : 0;
}

/**
 * Har bir hisob bo'yicha qoldiq.
 *
 * Belgilar qoidasi:
 *  - `income`   → `accountId` hisobiga `+amountMinor`
 *  - `expense`  → `accountId` hisobidan `-amountMinor`
 *  - `transfer` → `accountId` dan `-amountMinor`, `counterAccountId` ga `+counterAmountMinor`
 */
export function computeBalances(
  accounts: readonly Account[],
  transactions: readonly Transaction[],
  options: BalanceOptions,
): AccountBalance[] {
  const totals = new Map<string, MinorUnits>();
  for (const account of accounts) totals.set(account.id, 0);

  const add = (accountId: string | null, delta: MinorUnits): void => {
    if (!accountId) return;
    const current = totals.get(accountId);
    if (current === undefined) return; // O'chirilgan hisobga tegishli yozuv — e'tiborsiz.
    totals.set(accountId, current + delta);
  };

  for (const tx of transactions) {
    if (tx.deletedAt) continue;
    if (options.asOf && tx.date > options.asOf) continue;

    switch (tx.kind) {
      case 'income':
        add(tx.accountId, tx.amountMinor);
        break;
      case 'expense':
        add(tx.accountId, -tx.amountMinor);
        break;
      case 'transfer':
        add(tx.accountId, -tx.amountMinor);
        add(tx.counterAccountId, tx.counterAmountMinor ?? 0);
        break;
      default:
        break;
    }
  }

  return accounts
    .map((account) => {
      const balanceMinor = totals.get(account.id) ?? 0;
      const rate = rateFor(account.currency, options.baseCurrency, options.rates);
      const baseBalanceMinor =
        rate > 0 ? convertMinor(balanceMinor, account.currency, options.baseCurrency, rate) : 0;
      return {
        accountId: account.id,
        name: account.name,
        currency: account.currency,
        balanceMinor,
        baseBalanceMinor,
        showInBalance: account.showInBalance,
        sortOrder: account.sortOrder,
      };
    })
    .sort((a, b) => a.sortOrder - b.sortOrder || a.name.localeCompare(b.name));
}

/** Barcha hisoblarning hisobot valyutasidagi umumiy qiymati. */
export function totalInBase(balances: readonly AccountBalance[]): MinorUnits {
  let total = 0;
  for (const balance of balances) total += balance.baseBalanceMinor;
  return total;
}

/**
 * O'tkazma yozuvining to'g'riligini tekshiradi — bir hisobdan o'ziga o'tkazma
 * yoki yo'q hisobga o'tkazma bo'lmasligi kerak.
 */
export function validateTransferShape(tx: Pick<Transaction, 'kind' | 'accountId' | 'counterAccountId'>): string | null {
  if (tx.kind !== 'transfer') return null;
  if (!tx.counterAccountId) return "O'tkazma uchun ikkinchi hisob ko'rsatilmagan";
  if (tx.accountId === tx.counterAccountId) return 'Bir xil hisoblar orasida o‘tkazma qilib bo‘lmaydi';
  return null;
}
