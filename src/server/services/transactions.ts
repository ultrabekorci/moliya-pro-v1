/**
 * Tranzaksiyalar xizmati — tizimning yuragi.
 *
 * v1 dagi eng og'ir xatolar shu yerda yopilgan:
 *  - Summa hisobning O'Z valyutasida saqlanadi. Dollar chiqimini tahrirlaganda
 *    u qayta kursga ko'paytirilmaydi va izohga «(Aslida: ...)» qo'shilmaydi.
 *  - `batchId` — `Date.now()` emas, UUID. Ikki foydalanuvchi bir vaqtda saqlasa
 *    ham identifikatorlar to'qnashmaydi va begona yozuv o'chib ketmaydi.
 *  - Barcha yozish amallari `LockService` ostida va bitta guruh atomik yangilanadi.
 *  - O'chirish — yumshoq (`deletedAt`), tarix yo'qolmaydi.
 *  - Ro'yxat serverda filtrlanadi va sahifalanadi (butun jadval klientga tashlanmaydi).
 */

import { AppError, validationError } from '../../shared/api.js';
import type {
  BalanceQuery,
  DashboardQuery,
  ListTransactionsQuery,
  SaveTransactionInput,
} from '../../shared/api.js';
import type {
  AccountBalance,
  DashboardStats,
  Page,
  Transaction,
  TransactionBatch,
  TransactionKind,
} from '../../shared/types.js';
import { computeBalances, validateTransferShape } from '../../shared/balance.js';
import { computeDashboard } from '../../shared/stats.js';
import { convertMinor, parseAmountToMinor } from '../../shared/money.js';
import { isIsoDate, todayIso } from '../../shared/dates.js';
import {
  normalizePaging,
  optionalId,
  requireIsoDate,
  requireTransactionKind,
} from '../../shared/validation.js';
import { accountsTable, categoriesTable, pointsTable, transactionsTable } from '../db.js';
import type { StoredUser } from '../db.js';
import { newId, scriptTimeZone, withLock } from '../sheets.js';
import { getCatalog } from './catalog.js';
import { getConfig } from './config.js';
import { currentRates, getRate } from './fx.js';
import * as audit from './audit.js';

function liveTransactions(): Transaction[] {
  return transactionsTable.filter((tx) => !tx.deletedAt);
}

function requireAccount(accountId: unknown, field: string) {
  const id = optionalId(accountId, field);
  if (!id) throw validationError('Hisobni tanlang', field);
  const account = accountsTable.findById(id);
  if (!account) throw validationError('Bunday hisob yo‘q', field);
  return account;
}

function assertAmountLimit(minor: number, currency: string): void {
  const config = getConfig();
  const maxMinor = parseAmountToMinor(config.maxTransactionAmount, currency);
  if (minor > maxMinor) {
    throw validationError(`Summa ${config.maxTransactionAmount.toLocaleString('en-US')} dan oshmasin`, 'amount');
  }
}

function parsePositiveAmount(raw: unknown, currency: string, field: string): number {
  let minor: number;
  try {
    minor = parseAmountToMinor(typeof raw === 'number' ? raw : String(raw ?? ''), currency);
  } catch (error) {
    throw validationError(error instanceof Error ? error.message : 'Summa noto‘g‘ri', field);
  }
  if (minor <= 0) throw validationError('Summa noldan katta bo‘lishi kerak', field);
  assertAmountLimit(minor, currency);
  return minor;
}

interface PreparedRow {
  accountId: string;
  amountMinor: number;
  currency: string;
  counterAccountId: string | null;
  counterAmountMinor: number | null;
  counterCurrency: string | null;
  fxRate: number;
  baseAmountMinor: number;
}

function toBase(amountMinor: number, currency: string, date: string): { rate: number; baseMinor: number } {
  const config = getConfig();
  if (currency.toUpperCase() === config.baseCurrency.toUpperCase()) {
    return { rate: 1, baseMinor: amountMinor };
  }
  const rate = getRate(date, currency).rate;
  return {
    rate,
    baseMinor: convertMinor(amountMinor, currency, config.baseCurrency, rate),
  };
}

function prepareRows(input: SaveTransactionInput, kind: TransactionKind, date: string): PreparedRow[] {
  if (kind === 'income') {
    const entries = Array.isArray(input.entries) ? input.entries : [];
    const rows: PreparedRow[] = [];
    const seen = new Set<string>();

    for (const entry of entries) {
      const account = requireAccount(entry?.accountId, 'accountId');
      const raw = String(entry?.amount ?? '').trim();
      if (raw === '' || raw === '0') continue;

      if (seen.has(account.id)) {
        throw validationError(`«${account.name}» hisobi ikki marta ko‘rsatilgan`, 'entries');
      }
      seen.add(account.id);

      const amountMinor = parsePositiveAmount(raw, account.currency, 'entries');
      const { rate, baseMinor } = toBase(amountMinor, account.currency, date);
      rows.push({
        accountId: account.id,
        amountMinor,
        currency: account.currency,
        counterAccountId: null,
        counterAmountMinor: null,
        counterCurrency: null,
        fxRate: rate,
        baseAmountMinor: baseMinor,
      });
    }

    if (rows.length === 0) throw validationError('Kamida bitta to‘lov summasini kiriting', 'entries');
    return rows;
  }

  const account = requireAccount(input.accountId, 'accountId');
  const amountMinor = parsePositiveAmount(input.amount, account.currency, 'amount');

  if (kind === 'expense') {
    const { rate, baseMinor } = toBase(amountMinor, account.currency, date);
    return [
      {
        accountId: account.id,
        amountMinor,
        currency: account.currency,
        counterAccountId: null,
        counterAmountMinor: null,
        counterCurrency: null,
        fxRate: rate,
        baseAmountMinor: baseMinor,
      },
    ];
  }

  // transfer
  const counter = requireAccount(input.counterAccountId, 'counterAccountId');
  const shapeError = validateTransferShape({
    kind: 'transfer',
    accountId: account.id,
    counterAccountId: counter.id,
  });
  if (shapeError) throw validationError(shapeError, 'counterAccountId');

  const sameCurrency = account.currency.toUpperCase() === counter.currency.toUpperCase();
  const rawCounter = input.counterAmount === undefined || input.counterAmount === null || input.counterAmount === ''
    ? null
    : input.counterAmount;

  if (!sameCurrency && rawCounter === null) {
    throw validationError('Turli valyutadagi o‘tkazma uchun ikkinchi summani kiriting', 'counterAmount');
  }

  const counterAmountMinor =
    rawCounter === null ? amountMinor : parsePositiveAmount(rawCounter, counter.currency, 'counterAmount');

  const { rate, baseMinor } = toBase(amountMinor, account.currency, date);
  return [
    {
      accountId: account.id,
      amountMinor,
      currency: account.currency,
      counterAccountId: counter.id,
      counterAmountMinor,
      counterCurrency: counter.currency,
      fxRate: rate,
      baseAmountMinor: baseMinor,
    },
  ];
}

export function save(input: SaveTransactionInput, user: StoredUser): { batchId: string } {
  const config = getConfig();
  const kind = requireTransactionKind(input.kind);
  const date = requireIsoDate(input.date, 'sana');

  const note = String(input.note ?? '').trim();
  if (note.length > config.maxNoteLength) {
    throw validationError(`Izoh ${config.maxNoteLength} ta belgidan oshmasin`, 'note');
  }

  const pointId = optionalId(input.pointId, 'pointId');
  if (pointId && !pointsTable.findById(pointId)) throw validationError('Bunday savdo nuqtasi yo‘q', 'pointId');

  const categoryId = optionalId(input.categoryId, 'categoryId');
  if (categoryId && !categoriesTable.findById(categoryId)) {
    throw validationError('Bunday kategoriya yo‘q', 'categoryId');
  }
  if (kind === 'expense' && !categoryId) throw validationError('Kategoriyani tanlang', 'categoryId');

  // Firma: nuqta orqali aniqlanadi, aks holda bevosita ko'rsatiladi.
  let firmId = optionalId(input.firmId, 'firmId');
  if (pointId) {
    const point = pointsTable.findById(pointId);
    if (point) firmId = point.firmId;
  }

  const rows = prepareRows(input, kind, date);
  const now = new Date().toISOString();

  return withLock(() => {
    const existingBatchId = optionalId(input.batchId, 'batchId');
    let batchId = existingBatchId;

    if (existingBatchId) {
      const existing = transactionsTable.filter(
        (tx) => tx.batchId === existingBatchId && !tx.deletedAt,
      );
      if (existing.length === 0) throw new AppError('NOT_FOUND', 'Tahrirlanayotgan yozuv topilmadi');

      // Eski qatorlar yumshoq o'chiriladi — moliyaviy tarix saqlanadi.
      transactionsTable.updateMany(
        existing.map((tx) => ({
          id: tx.id,
          patch: { deletedAt: now, deletedBy: user.id } as Partial<Transaction>,
        })),
      );
      audit.record(user, 'transaction.edit', 'transaction', existingBatchId, {
        before: existing.map((tx) => ({ id: tx.id, amountMinor: tx.amountMinor, accountId: tx.accountId })),
      });
    } else {
      batchId = newId();
    }

    const finalBatchId = batchId ?? newId();
    const records: Transaction[] = rows.map((row) => ({
      id: newId(),
      batchId: finalBatchId,
      date,
      kind,
      firmId,
      pointId,
      categoryId,
      accountId: row.accountId,
      amountMinor: row.amountMinor,
      currency: row.currency,
      counterAccountId: row.counterAccountId,
      counterAmountMinor: row.counterAmountMinor,
      counterCurrency: row.counterCurrency,
      fxRate: row.fxRate,
      baseAmountMinor: row.baseAmountMinor,
      note,
      createdAt: now,
      createdBy: user.id,
      updatedAt: now,
      updatedBy: user.id,
      deletedAt: null,
      deletedBy: null,
    }));

    transactionsTable.insertMany(records);
    if (!existingBatchId) {
      audit.record(user, 'transaction.create', 'transaction', finalBatchId, {
        kind,
        date,
        rows: records.length,
      });
    }
    return { batchId: finalBatchId };
  });
}

/** Guruhning turini aniqlaydi — o'chirishdan oldin ruxsatni tekshirish uchun. */
export function batchKind(batchId: string): TransactionKind | null {
  const row = transactionsTable.find((tx) => tx.batchId === batchId && !tx.deletedAt);
  return row ? row.kind : null;
}

export function remove(batchId: string, user: StoredUser): void {
  withLock(() => {
    const rows = transactionsTable.filter((tx) => tx.batchId === batchId && !tx.deletedAt);
    if (rows.length === 0) throw new AppError('NOT_FOUND', 'Yozuv topilmadi');

    const now = new Date().toISOString();
    transactionsTable.updateMany(
      rows.map((tx) => ({
        id: tx.id,
        patch: { deletedAt: now, deletedBy: user.id } as Partial<Transaction>,
      })),
    );
    audit.record(user, 'transaction.delete', 'transaction', batchId, { rows: rows.length });
  });
}

// ---------------------------------------------------------------------------
// O'qish
// ---------------------------------------------------------------------------

function applyFilter(query: ListTransactionsQuery): Transaction[] {
  const filter = query.filter ?? {};
  const includeDeleted = filter.includeDeleted === true;
  const search = (filter.search ?? '').trim().toLowerCase();

  const catalog = getCatalog();
  const nameById = new Map<string, string>();
  for (const item of [...catalog.points, ...catalog.categories, ...catalog.accounts, ...catalog.firms]) {
    nameById.set(item.id, item.name.toLowerCase());
  }

  return transactionsTable
    .filter((tx) => {
      if (!includeDeleted && tx.deletedAt) return false;
      if (filter.range) {
        if (!isIsoDate(filter.range.start) || !isIsoDate(filter.range.end)) return false;
        if (tx.date < filter.range.start || tx.date > filter.range.end) return false;
      }
      if (filter.kinds && filter.kinds.length > 0 && !filter.kinds.includes(tx.kind)) return false;
      if (filter.firmId && tx.firmId !== filter.firmId) return false;
      if (filter.pointId && tx.pointId !== filter.pointId) return false;
      if (filter.categoryId && tx.categoryId !== filter.categoryId) return false;
      if (filter.accountId && tx.accountId !== filter.accountId && tx.counterAccountId !== filter.accountId) {
        return false;
      }
      if (search !== '') {
        const haystack = [
          tx.note,
          nameById.get(tx.pointId ?? '') ?? '',
          nameById.get(tx.categoryId ?? '') ?? '',
          nameById.get(tx.accountId) ?? '',
          nameById.get(tx.counterAccountId ?? '') ?? '',
        ]
          .join(' ')
          .toLowerCase();
        if (!haystack.includes(search)) return false;
      }
      return true;
    })
    .sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : b.createdAt.localeCompare(a.createdAt)));
}

export function list(query: ListTransactionsQuery): Page<Transaction> {
  const paging = normalizePaging(query.offset, query.limit);
  const filtered = applyFilter(query);
  return {
    items: filtered.slice(paging.offset, paging.offset + paging.limit),
    total: filtered.length,
    offset: paging.offset,
    limit: paging.limit,
  };
}

export function listGrouped(query: ListTransactionsQuery): Page<TransactionBatch> {
  const paging = normalizePaging(query.offset, query.limit);
  const filtered = applyFilter(query);

  const batches = new Map<string, TransactionBatch>();
  for (const tx of filtered) {
    const existing = batches.get(tx.batchId);
    if (existing) {
      existing.entries.push(tx);
      continue;
    }
    batches.set(tx.batchId, {
      batchId: tx.batchId,
      date: tx.date,
      kind: tx.kind,
      firmId: tx.firmId,
      pointId: tx.pointId,
      categoryId: tx.categoryId,
      note: tx.note,
      entries: [tx],
      updatedAt: tx.updatedAt,
      updatedBy: tx.updatedBy,
    });
  }

  const all = Array.from(batches.values());
  return {
    items: all.slice(paging.offset, paging.offset + paging.limit),
    total: all.length,
    offset: paging.offset,
    limit: paging.limit,
  };
}

export function balances(query: BalanceQuery): AccountBalance[] {
  const config = getConfig();
  const catalog = getCatalog();
  const today = todayIso(scriptTimeZone());
  const asOf = query.asOf && isIsoDate(query.asOf) ? query.asOf : undefined;

  const rates = currentRates(
    catalog.accounts.map((account) => account.currency),
    today,
  );

  return computeBalances(catalog.accounts, liveTransactions(), {
    baseCurrency: config.baseCurrency,
    rates,
    ...(asOf ? { asOf } : {}),
  });
}

export function dashboard(query: DashboardQuery): DashboardStats {
  const config = getConfig();
  const catalog = getCatalog();

  if (!isIsoDate(query.range?.start) || !isIsoDate(query.range?.end)) {
    throw validationError('Davr sanalari noto‘g‘ri', 'range');
  }

  return computeDashboard(
    liveTransactions(),
    {
      firms: catalog.firms,
      points: catalog.points,
      categories: catalog.categories,
      accounts: catalog.accounts,
      baseCurrency: config.baseCurrency,
      sharedExpenseAllocation: config.sharedExpenseAllocation,
    },
    { range: query.range, firmId: optionalId(query.firmId, 'firmId') },
    query.previousRange && isIsoDate(query.previousRange.start) ? query.previousRange : query.range,
  );
}
