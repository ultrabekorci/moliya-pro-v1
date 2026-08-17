/**
 * v1 → v2 ma'lumot ko'chirish.
 *
 * Bir martalik amal: eski varaqlarni (`Kirim Chiqim`, `Data`, `Users`) o'qib,
 * yangi strukturaga yozadi. Eski varaqlar O'ZGARTIRILMAYDI — xohlagan paytda
 * qaytib tekshirsa bo'ladi.
 *
 * Apps Script muharriridan ishga tushiring:
 *   migrateDryRun()  — hech narsa yozmaydi, faqat nima bo'lishini hisoblab beradi
 *   migrateFromV1()  — haqiqiy ko'chirish
 */

import { AppError } from '../shared/api.js';
import { isIsoDate, toIsoDate } from '../shared/dates.js';
import { convertMinor, parseAmountToMinor } from '../shared/money.js';
import { sanitizePermissions } from '../shared/permissions.js';
import type { Account, Category, Firm, Point, Role, Transaction } from '../shared/types.js';
import {
  accountsTable,
  categoriesTable,
  ensureSchema,
  firmsTable,
  pointsTable,
  transactionsTable,
  usersTable,
} from './db.js';
import type { StoredUser } from './db.js';
import { newId, scriptTimeZone, withLock } from './sheets.js';
import { buildPasswordFields } from './services/auth.js';
import { getConfig } from './services/config.js';

export interface MigrationReport {
  dryRun: boolean;
  firms: number;
  accounts: number;
  points: number;
  categories: number;
  users: number;
  transactions: number;
  skipped: string[];
  warnings: string[];
}

const OLD_TRANSACTIONS_SHEET = 'Kirim Chiqim';
const OLD_CATALOG_SHEET = 'Data';
const OLD_USERS_SHEET = 'Users';
const MIGRATED_MARKER = 'v1-migration';

function sheet(name: string): GoogleAppsScript.Spreadsheet.Sheet | null {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function readRows(name: string): unknown[][] {
  const target = sheet(name);
  if (!target) return [];
  const lastRow = target.getLastRow();
  const lastColumn = target.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return [];
  return target.getRange(2, 1, lastRow - 1, lastColumn).getValues();
}

function text(value: unknown): string {
  return value === null || value === undefined ? '' : String(value).trim();
}

function num(value: unknown): number {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const parsed = Number(text(value).replace(/[\s,]/g, ''));
  return Number.isFinite(parsed) ? parsed : 0;
}

/** Eski izohdagi «(Aslida: $500 @ 12500)» yozuvidan asl USD summasini ajratadi. */
export function extractOriginalUsd(note: string): number {
  const match = /Aslida:\s*\$?\s*([\d.,\s]+)/i.exec(note);
  if (!match?.[1]) return 0;
  const value = Number(match[1].replace(/[\s,]/g, ''));
  return Number.isFinite(value) && value > 0 ? value : 0;
}

/** Eski to'lov turi nomiga qarab valyutani taxmin qiladi. */
function guessCurrency(name: string, baseCurrency: string): string {
  const upper = name.toUpperCase();
  if (upper.includes('DOLLAR') || upper.includes('USD') || upper.includes('$')) return 'USD';
  if (upper.includes('EVRO') || upper.includes('EUR')) return 'EUR';
  return baseCurrency;
}

interface Lookups {
  accountByName: Map<string, Account>;
  pointByName: Map<string, Point>;
  categoryByName: Map<string, Category>;
  defaultFirm: Firm;
}

function migrateCatalog(report: MigrationReport, baseCurrency: string): Lookups {
  const rows = readRows(OLD_CATALOG_SHEET);

  const firmNames = new Set<string>();
  const pointNames = new Set<string>();
  const paymentNames = new Set<string>();
  const categoryNames = new Set<string>();

  for (const row of rows) {
    const firm = text(row[0]);
    const point = text(row[1]);
    const payment = text(row[2]);
    const category = text(row[3]);
    if (firm) firmNames.add(firm);
    if (point) pointNames.add(point);
    if (payment) paymentNames.add(payment);
    if (category) categoryNames.add(category);
  }

  // v1 da to'lov turlarining bir qismi kodga qotirib yozilgan edi.
  for (const fallback of ['Naqd', 'P2P', 'Bank', 'Dollar']) paymentNames.add(fallback);
  categoryNames.add('Savdo tushumi');

  // Firma: eskisida bu ustun amalda ishlatilmagan, shuning uchun bittasi yetarli.
  const firmName = Array.from(firmNames)[0] ?? 'Asosiy firma';
  let defaultFirm = firmsTable.find((firm) => firm.name === firmName);
  if (!defaultFirm) {
    defaultFirm = {
      id: newId(),
      name: firmName,
      active: true,
      expenseAllocation: 'proRataIncome',
      sortOrder: 10,
    };
    if (!report.dryRun) firmsTable.insert(defaultFirm);
    report.firms += 1;
  }

  const accountByName = new Map<string, Account>();
  let order = 0;
  for (const name of paymentNames) {
    order += 10;
    const existing = accountsTable.find((account) => account.name === name);
    if (existing) {
      accountByName.set(name.toUpperCase(), existing);
      continue;
    }
    const account: Account = {
      id: newId(),
      name,
      currency: guessCurrency(name, baseCurrency),
      active: true,
      showInBalance: true,
      sortOrder: order,
    };
    if (!report.dryRun) accountsTable.insert(account);
    accountByName.set(name.toUpperCase(), account);
    report.accounts += 1;
  }

  const pointByName = new Map<string, Point>();
  order = 0;
  // v1 da bu nomlar kodga qotirib yozilgan "e'tiborga olinmaydigan" ro'yxat edi.
  const legacyExcluded = ['DONIYOR AKA', 'BOSHQA', 'DIREKTOR', 'KASSA'];
  for (const name of pointNames) {
    order += 10;
    const existing = pointsTable.find((point) => point.name === name);
    if (existing) {
      pointByName.set(name.toUpperCase(), existing);
      continue;
    }
    const point: Point = {
      id: newId(),
      name,
      firmId: defaultFirm.id,
      active: true,
      excludeFromRevenue: legacyExcluded.some((bad) => name.toUpperCase().includes(bad)),
      sortOrder: order,
    };
    if (!report.dryRun) pointsTable.insert(point);
    pointByName.set(name.toUpperCase(), point);
    report.points += 1;
  }
  if (pointNames.size > 0) {
    report.warnings.push(
      'Barcha savdo nuqtalari bitta firmaga bog‘landi. Sozlamalar → Savdo nuqtalari bo‘limidan ' +
        'ularni kerakli firmalarga taqsimlang.',
    );
  }

  const categoryByName = new Map<string, Category>();
  order = 0;
  for (const name of categoryNames) {
    order += 10;
    const existing = categoriesTable.find((category) => category.name === name);
    if (existing) {
      categoryByName.set(name.toUpperCase(), existing);
      continue;
    }
    const category: Category = {
      id: newId(),
      name,
      kind: name.toUpperCase() === 'SAVDO TUSHUMI' ? 'income' : 'expense',
      active: true,
      sortOrder: order,
    };
    if (!report.dryRun) categoriesTable.insert(category);
    categoryByName.set(name.toUpperCase(), category);
    report.categories += 1;
  }

  return { accountByName, pointByName, categoryByName, defaultFirm };
}

function migrateUsers(report: MigrationReport): void {
  const rows = readRows(OLD_USERS_SHEET);
  const now = new Date().toISOString();

  for (const row of rows) {
    const login = text(row[0]).toLowerCase();
    const password = text(row[1]);
    if (login === '') continue;
    if (usersTable.find((user) => user.login === login)) continue;

    const roleText = text(row[2]).toLowerCase();
    const role: Role = roleText === 'admin' ? 'admin' : 'operator';
    const status = text(row[3]).toLowerCase() === 'inactive' ? 'blocked' : 'active';

    let permissions = {};
    const rawPermissions = text(row[4]);
    if (rawPermissions !== '') {
      try {
        permissions = sanitizePermissions(JSON.parse(rawPermissions));
      } catch {
        report.warnings.push(`«${login}» huquqlari o‘qilmadi — rol standartiga qaytarildi.`);
      }
    }

    if (password === '') {
      report.skipped.push(`Foydalanuvchi «${login}»: paroli bo‘sh`);
      continue;
    }

    if (!report.dryRun) {
      const user: StoredUser = {
        id: newId(),
        login,
        displayName: login,
        ...buildPasswordFields(password),
        role,
        status,
        permissions,
        createdAt: now,
        updatedAt: now,
        lastLoginAt: null,
      };
      usersTable.insert(user);
    }
    report.users += 1;
  }

  if (report.users > 0) {
    report.warnings.push(
      'Eski parollar ochiq matnda saqlangan edi. Ko‘chirishdan so‘ng ularni xesh ko‘rinishida ' +
        'saqlaymiz, lekin barcha xodimlarga parolni almashtirishni tavsiya qiling.',
    );
  }
}

function migrateTransactions(report: MigrationReport, lookups: Lookups, baseCurrency: string): void {
  const rows = readRows(OLD_TRANSACTIONS_SHEET);
  const timeZone = scriptTimeZone();
  const now = new Date().toISOString();
  const records: Transaction[] = [];

  const accountFor = (name: string): Account | null => lookups.accountByName.get(name.toUpperCase()) ?? null;

  rows.forEach((row, index) => {
    const rowNumber = index + 2;
    const date = toIsoDate(row[0], timeZone);
    if (!date || !isIsoDate(date)) {
      report.skipped.push(`${rowNumber}-qator: sana o‘qilmadi`);
      return;
    }

    const pointName = text(row[2]);
    const categoryName = text(row[3]);
    const paymentName = text(row[4]);
    const kirim = num(row[5]);
    const chiqim = num(row[6]);
    const note = text(row[7]);
    const batchId = text(row[8]) || newId();

    const isTransfer = categoryName === "O'tkazma" || categoryName === 'Transfer';
    const base = {
      id: newId(),
      batchId,
      date,
      note,
      createdAt: now,
      createdBy: MIGRATED_MARKER,
      updatedAt: now,
      updatedBy: MIGRATED_MARKER,
      deletedAt: null,
      deletedBy: null,
    };

    if (isTransfer) {
      // Eski format: to'lov turi «Naqd -> Dollar», yoki 10/11-ustunlarda manba/manzil.
      let sourceName = text(row[9]);
      let destName = text(row[10]);
      if ((sourceName === '' || destName === '') && paymentName.includes('->')) {
        const parts = paymentName.split('->');
        sourceName = text(parts[0]);
        destName = text(parts[1]);
      }

      const source = accountFor(sourceName);
      const dest = accountFor(destName);
      if (!source || !dest) {
        report.skipped.push(`${rowNumber}-qator: o‘tkazma hisoblari topilmadi (${sourceName} → ${destName})`);
        return;
      }

      const outMinor = parseAmountToMinor(chiqim, source.currency);
      const inMinor = parseAmountToMinor(kirim, dest.currency);
      if (outMinor <= 0 || inMinor <= 0) {
        report.skipped.push(`${rowNumber}-qator: o‘tkazma summasi noto‘g‘ri`);
        return;
      }

      records.push({
        ...base,
        kind: 'transfer',
        firmId: null,
        pointId: null,
        categoryId: null,
        accountId: source.id,
        amountMinor: outMinor,
        currency: source.currency,
        counterAccountId: dest.id,
        counterAmountMinor: inMinor,
        counterCurrency: dest.currency,
        fxRate: 1,
        baseAmountMinor: source.currency === baseCurrency ? outMinor : 0,
      });
      return;
    }

    const account = accountFor(paymentName);
    if (!account) {
      report.skipped.push(`${rowNumber}-qator: «${paymentName}» hisobi topilmadi`);
      return;
    }

    const point = lookups.pointByName.get(pointName.toUpperCase()) ?? null;
    const category = lookups.categoryByName.get(categoryName.toUpperCase()) ?? null;

    if (kirim > 0) {
      const amountMinor = parseAmountToMinor(kirim, account.currency);
      records.push({
        ...base,
        kind: 'income',
        firmId: point ? point.firmId : lookups.defaultFirm.id,
        pointId: point?.id ?? null,
        categoryId: category?.id ?? null,
        accountId: account.id,
        amountMinor,
        currency: account.currency,
        counterAccountId: null,
        counterAmountMinor: null,
        counterCurrency: null,
        fxRate: 1,
        baseAmountMinor: account.currency === baseCurrency ? amountMinor : 0,
      });
      return;
    }

    if (chiqim > 0) {
      // Eski dollar chiqimlari UZS ga o'girib saqlangan; asl USD summasi
      // 10-ustunda yoki izohda «Aslida: $X» ko'rinishida bo'lishi mumkin.
      let amountMinor: number;
      let fxRate = 1;
      let baseAmountMinor: number;

      if (account.currency !== baseCurrency) {
        const originalUsd = num(row[9]) > 0 ? num(row[9]) : extractOriginalUsd(note);
        if (originalUsd > 0) {
          amountMinor = parseAmountToMinor(originalUsd, account.currency);
          baseAmountMinor = parseAmountToMinor(chiqim, baseCurrency);
          fxRate = originalUsd > 0 ? chiqim / originalUsd : 1;
        } else {
          report.skipped.push(
            `${rowNumber}-qator: valyutadagi chiqimning asl summasi topilmadi — qo‘lda kiriting`,
          );
          return;
        }
      } else {
        amountMinor = parseAmountToMinor(chiqim, account.currency);
        baseAmountMinor = amountMinor;
      }

      records.push({
        ...base,
        kind: 'expense',
        firmId: point ? point.firmId : null,
        pointId: point?.id ?? null,
        categoryId: category?.id ?? null,
        accountId: account.id,
        amountMinor,
        currency: account.currency,
        counterAccountId: null,
        counterAmountMinor: null,
        counterCurrency: null,
        fxRate,
        baseAmountMinor,
      });
      return;
    }

    report.skipped.push(`${rowNumber}-qator: summa nol`);
  });

  // Hisobot valyutasidagi ekvivalenti hisoblanmaganlarni to'ldiramiz.
  for (const record of records) {
    if (record.baseAmountMinor === 0 && record.currency !== baseCurrency && record.fxRate > 0) {
      record.baseAmountMinor = convertMinor(record.amountMinor, record.currency, baseCurrency, record.fxRate);
    }
  }

  report.transactions = records.length;
  if (!report.dryRun && records.length > 0) transactionsTable.insertMany(records);
}

function run(dryRun: boolean): MigrationReport {
  if (!sheet(OLD_TRANSACTIONS_SHEET)) {
    throw new AppError('NOT_FOUND', `«${OLD_TRANSACTIONS_SHEET}» varag‘i topilmadi — ko‘chirish uchun v1 bazasi kerak`);
  }

  const report: MigrationReport = {
    dryRun,
    firms: 0,
    accounts: 0,
    points: 0,
    categories: 0,
    users: 0,
    transactions: 0,
    skipped: [],
    warnings: [],
  };

  return withLock(() => {
    ensureSchema();
    const baseCurrency = getConfig().baseCurrency;

    if (!dryRun && transactionsTable.count() > 0) {
      throw new AppError(
        'CONFLICT',
        'Yangi `Transactions` varag‘i bo‘sh emas. Takroriy ko‘chirishning oldini olish uchun to‘xtatildi.',
      );
    }

    const lookups = migrateCatalog(report, baseCurrency);
    migrateUsers(report);
    migrateTransactions(report, lookups, baseCurrency);

    if (report.skipped.length > 20) {
      const extra = report.skipped.length - 20;
      report.skipped = report.skipped.slice(0, 20);
      report.skipped.push(`…va yana ${extra} ta qator`);
    }
    return report;
  });
}

/** Sinov rejimi: hech narsa yozilmaydi, faqat natija hisoblanadi. */
export function migrateDryRun(): string {
  return JSON.stringify(run(true), null, 2);
}

/** Haqiqiy ko'chirish. Eski varaqlar o'zgartirilmaydi. */
export function migrateFromV1(): string {
  return JSON.stringify(run(false), null, 2);
}
