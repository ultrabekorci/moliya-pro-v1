/**
 * Ma'lumotnomalar: firmalar, hisoblar, savdo nuqtalari, kategoriyalar.
 *
 * v1 dan farqlar:
 *  - Firmalar endi katalog yozuvi (kodda "Greenpen"/"Smartmiz" degan narsa yo'q).
 *  - Hisoblar (to'lov turlari) ixtiyoriy valyutada — Uzcard/Humo/Perechislenie
 *    balansdan tushib qolmaydi.
 *  - Yozuv o'chirilganda undan foydalanayotgan tranzaksiyalar tekshiriladi:
 *    ishlatilayotgan bo'lsa o'chirish o'rniga "nofaol" qilish taklif etiladi.
 *    Bu v1 dagi "yetim ma'lumot" muammosini yopadi.
 */

import { AppError, validationError } from '../../shared/api.js';
import type {
  AccountInput,
  CategoryInput,
  FirmInput,
  PointInput,
} from '../../shared/api.js';
import type { Account, Catalog, Category, Firm, Point } from '../../shared/types.js';
import { CATEGORY_KINDS, EXPENSE_ALLOCATIONS } from '../../shared/types.js';
import {
  LIMITS,
  requireBoolean,
  requireCurrency,
  requireEnum,
  requireId,
  requireString,
} from '../../shared/validation.js';
import { accountsTable, categoriesTable, firmsTable, pointsTable, transactionsTable } from '../db.js';
import { newId, withLock } from '../sheets.js';

function bySort<T extends { sortOrder: number; name: string }>(a: T, b: T): number {
  return a.sortOrder - b.sortOrder || a.name.localeCompare(b.name);
}

export function getCatalog(): Catalog {
  return {
    firms: firmsTable.all().slice().sort(bySort),
    accounts: accountsTable.all().slice().sort(bySort),
    points: pointsTable.all().slice().sort(bySort),
    categories: categoriesTable.all().slice().sort(bySort),
  };
}

function nextSortOrder(items: ReadonlyArray<{ sortOrder: number }>): number {
  return items.reduce((max, item) => Math.max(max, item.sortOrder), 0) + 10;
}

function assertUniqueName(
  items: ReadonlyArray<{ id: string; name: string }>,
  name: string,
  currentId: string | null,
  label: string,
): void {
  const normalized = name.trim().toLowerCase();
  const clash = items.find((item) => item.name.trim().toLowerCase() === normalized && item.id !== currentId);
  if (clash) throw validationError(`Bunday ${label} allaqachon mavjud: «${name}»`, 'name');
}

// ---------------------------------------------------------------------------
// Firmalar
// ---------------------------------------------------------------------------

export function saveFirm(input: FirmInput): { id: string } {
  const name = requireString(input.name, 'nom', { min: LIMITS.nameMin, max: LIMITS.nameMax });
  const allocation = input.expenseAllocation
    ? requireEnum(input.expenseAllocation, EXPENSE_ALLOCATIONS, 'expenseAllocation')
    : 'proRataIncome';

  return withLock(() => {
    const existing = getCatalog().firms;
    const id = input.id ? requireId(input.id, 'id') : null;
    assertUniqueName(existing, name, id, 'firma');

    if (id) {
      const current = firmsTable.findById(id);
      if (!current) throw new AppError('NOT_FOUND', 'Firma topilmadi');
      firmsTable.update(id, {
        name,
        active: requireBoolean(input.active, current.active),
        expenseAllocation: allocation,
        sortOrder: typeof input.sortOrder === 'number' ? input.sortOrder : current.sortOrder,
      });
      return { id };
    }

    const firm: Firm = {
      id: newId(),
      name,
      active: requireBoolean(input.active, true),
      expenseAllocation: allocation,
      sortOrder: typeof input.sortOrder === 'number' ? input.sortOrder : nextSortOrder(existing),
    };
    firmsTable.insert(firm);
    return { id: firm.id };
  });
}

// ---------------------------------------------------------------------------
// Hisoblar
// ---------------------------------------------------------------------------

export function saveAccount(input: AccountInput): { id: string } {
  const name = requireString(input.name, 'nom', { min: LIMITS.nameMin, max: LIMITS.nameMax });
  const currency = requireCurrency(input.currency);

  return withLock(() => {
    const existing = getCatalog().accounts;
    const id = input.id ? requireId(input.id, 'id') : null;
    assertUniqueName(existing, name, id, 'hisob');

    if (id) {
      const current = accountsTable.findById(id);
      if (!current) throw new AppError('NOT_FOUND', 'Hisob topilmadi');
      // Valyutani o'zgartirish mavjud qoldiqni ma'nosiz qilib qo'yadi.
      if (current.currency !== currency && isAccountUsed(id)) {
        throw validationError(
          'Bu hisobda yozuvlar bor — valyutasini o‘zgartirib bo‘lmaydi. Yangi hisob oching.',
          'currency',
        );
      }
      accountsTable.update(id, {
        name,
        currency,
        active: requireBoolean(input.active, current.active),
        showInBalance: requireBoolean(input.showInBalance, current.showInBalance),
        sortOrder: typeof input.sortOrder === 'number' ? input.sortOrder : current.sortOrder,
      });
      return { id };
    }

    const account: Account = {
      id: newId(),
      name,
      currency,
      active: requireBoolean(input.active, true),
      showInBalance: requireBoolean(input.showInBalance, true),
      sortOrder: typeof input.sortOrder === 'number' ? input.sortOrder : nextSortOrder(existing),
    };
    accountsTable.insert(account);
    return { id: account.id };
  });
}

// ---------------------------------------------------------------------------
// Savdo nuqtalari
// ---------------------------------------------------------------------------

export function savePoint(input: PointInput): { id: string } {
  const name = requireString(input.name, 'nom', { min: LIMITS.nameMin, max: LIMITS.nameMax });
  const firmId = requireId(input.firmId, 'firma');
  if (!firmsTable.findById(firmId)) throw validationError('Bunday firma yo‘q', 'firmId');

  return withLock(() => {
    const existing = getCatalog().points;
    const id = input.id ? requireId(input.id, 'id') : null;
    assertUniqueName(existing, name, id, 'savdo nuqtasi');

    if (id) {
      const current = pointsTable.findById(id);
      if (!current) throw new AppError('NOT_FOUND', 'Savdo nuqtasi topilmadi');
      pointsTable.update(id, {
        name,
        firmId,
        active: requireBoolean(input.active, current.active),
        excludeFromRevenue: requireBoolean(input.excludeFromRevenue, current.excludeFromRevenue),
        sortOrder: typeof input.sortOrder === 'number' ? input.sortOrder : current.sortOrder,
      });
      return { id };
    }

    const point: Point = {
      id: newId(),
      name,
      firmId,
      active: requireBoolean(input.active, true),
      excludeFromRevenue: requireBoolean(input.excludeFromRevenue, false),
      sortOrder: typeof input.sortOrder === 'number' ? input.sortOrder : nextSortOrder(existing),
    };
    pointsTable.insert(point);
    return { id: point.id };
  });
}

// ---------------------------------------------------------------------------
// Kategoriyalar
// ---------------------------------------------------------------------------

export function saveCategory(input: CategoryInput): { id: string } {
  const name = requireString(input.name, 'nom', { min: LIMITS.nameMin, max: LIMITS.nameMax });
  const kind = requireEnum(input.kind, CATEGORY_KINDS, 'kind');

  return withLock(() => {
    const existing = getCatalog().categories;
    const id = input.id ? requireId(input.id, 'id') : null;
    assertUniqueName(existing, name, id, 'kategoriya');

    if (id) {
      const current = categoriesTable.findById(id);
      if (!current) throw new AppError('NOT_FOUND', 'Kategoriya topilmadi');
      categoriesTable.update(id, {
        name,
        kind,
        active: requireBoolean(input.active, current.active),
        sortOrder: typeof input.sortOrder === 'number' ? input.sortOrder : current.sortOrder,
      });
      return { id };
    }

    const category: Category = {
      id: newId(),
      name,
      kind,
      active: requireBoolean(input.active, true),
      sortOrder: typeof input.sortOrder === 'number' ? input.sortOrder : nextSortOrder(existing),
    };
    categoriesTable.insert(category);
    return { id: category.id };
  });
}

// ---------------------------------------------------------------------------
// O'chirish (ishlatilayotgan yozuvlar himoyalangan)
// ---------------------------------------------------------------------------

function isAccountUsed(accountId: string): boolean {
  return transactionsTable.all().some((tx) => tx.accountId === accountId || tx.counterAccountId === accountId);
}

function usageError(label: string): AppError {
  return new AppError(
    'CONFLICT',
    `Bu ${label} yozuvlarda ishlatilgan, shuning uchun o‘chirib bo‘lmaydi. ` +
      'Uni "nofaol" qilib qo‘ying — eski hisobotlar buzilmaydi, yangi yozuvlarda esa ko‘rinmaydi.',
  );
}

export function deleteFirm(id: string): void {
  withLock(() => {
    if (!firmsTable.findById(id)) throw new AppError('NOT_FOUND', 'Firma topilmadi');
    if (pointsTable.all().some((point) => point.firmId === id)) {
      throw new AppError('CONFLICT', 'Avval bu firmaga tegishli savdo nuqtalarini o‘chiring yoki boshqa firmaga o‘tkazing');
    }
    if (transactionsTable.all().some((tx) => tx.firmId === id)) throw usageError('firma');
    firmsTable.deleteById(id);
  });
}

export function deleteAccount(id: string): void {
  withLock(() => {
    if (!accountsTable.findById(id)) throw new AppError('NOT_FOUND', 'Hisob topilmadi');
    if (isAccountUsed(id)) throw usageError('hisob');
    accountsTable.deleteById(id);
  });
}

export function deletePoint(id: string): void {
  withLock(() => {
    if (!pointsTable.findById(id)) throw new AppError('NOT_FOUND', 'Savdo nuqtasi topilmadi');
    if (transactionsTable.all().some((tx) => tx.pointId === id)) throw usageError('savdo nuqtasi');
    pointsTable.deleteById(id);
  });
}

export function deleteCategory(id: string): void {
  withLock(() => {
    if (!categoriesTable.findById(id)) throw new AppError('NOT_FOUND', 'Kategoriya topilmadi');
    if (transactionsTable.all().some((tx) => tx.categoryId === id)) throw usageError('kategoriya');
    categoriesTable.deleteById(id);
  });
}

// ---------------------------------------------------------------------------
// Birinchi ishga tushirish uchun boshlang'ich to'plam
// ---------------------------------------------------------------------------

/** Bo'sh bazani ishlashga yaroqli minimal katalog bilan to'ldiradi. */
export function seedDefaults(baseCurrency: string): void {
  withLock(() => {
    if (firmsTable.count() === 0) {
      firmsTable.insert({
        id: newId(),
        name: 'Asosiy firma',
        active: true,
        expenseAllocation: 'proRataIncome',
        sortOrder: 10,
      });
    }

    if (accountsTable.count() === 0) {
      const defaults: Array<[string, string]> = [
        ['Naqd', baseCurrency],
        ['Plastik', baseCurrency],
        ['Bank hisobi', baseCurrency],
        ['Dollar kassa', 'USD'],
      ];
      defaults.forEach(([name, currency], index) => {
        accountsTable.insert({
          id: newId(),
          name,
          currency,
          active: true,
          showInBalance: true,
          sortOrder: (index + 1) * 10,
        });
      });
    }

    if (categoriesTable.count() === 0) {
      const defaults: Array<[string, 'income' | 'expense' | 'both']> = [
        ['Savdo tushumi', 'income'],
        ['Ish haqi', 'expense'],
        ['Ijara', 'expense'],
        ['Tovar xaridi', 'expense'],
        ['Kommunal', 'expense'],
        ['Boshqa', 'both'],
      ];
      defaults.forEach(([name, kind], index) => {
        categoriesTable.insert({
          id: newId(),
          name,
          kind,
          active: true,
          sortOrder: (index + 1) * 10,
        });
      });
    }

    if (pointsTable.count() === 0) {
      const firm = firmsTable.all()[0];
      if (firm) {
        pointsTable.insert({
          id: newId(),
          name: 'Asosiy do‘kon',
          firmId: firm.id,
          active: true,
          excludeFromRevenue: false,
          sortOrder: 10,
        });
      }
    }
  });
}
