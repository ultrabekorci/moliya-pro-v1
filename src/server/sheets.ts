/**
 * Google Sheets ustidagi tipizatsiyalangan repozitoriy.
 *
 * v1 dagi muammolar va ularning yechimi:
 *  - `row[8]`, `getRange(row, 10)` kabi sehrli indekslar → ustunlar SARLAVHA
 *    nomi bo'yicha topiladi, ustun tartibini o'zgartirish kodni buzmaydi.
 *  - Klientdan kelgan qator raqami bilan yozish/o'chirish → yozuvlar barqaror
 *    `id` (UUID) bo'yicha topiladi.
 *  - Bir vaqtda yozish (race) → barcha yozuv amallari `LockService` ostida.
 *  - Yetishmayotgan ustun/varaq → `ensure()` avtomatik yaratadi va to'ldiradi.
 */

import { AppError } from '../shared/api.js';

export type ColumnType = 'string' | 'number' | 'boolean' | 'json' | 'date' | 'timestamp';

export interface Column<T> {
  key: Extract<keyof T, string>;
  header: string;
  type: ColumnType;
  /** `true` bo'lsa bo'sh katak `null` sifatida o'qiladi. */
  nullable?: boolean;
  defaultValue?: unknown;
}

export interface TableSchema<T> {
  sheetName: string;
  columns: ReadonlyArray<Column<T>>;
  /** Yozuvlarni ajratuvchi ustun (odatda `id`). */
  idKey: Extract<keyof T, string>;
}

const LOCK_TIMEOUT_MS = 20_000;

/**
 * Yozuv amallarini butun skript bo'ylab ketma-ketlashtiradi.
 * Ichma-ich chaqirilganda qayta qulflamaydi.
 */
let lockDepth = 0;

export function withLock<T>(operation: () => T): T {
  if (lockDepth > 0) return operation();

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_TIMEOUT_MS)) {
    throw new AppError('CONFLICT', 'Tizim band, iltimos bir necha soniyadan so‘ng qayta urining');
  }
  lockDepth += 1;
  try {
    return operation();
  } finally {
    lockDepth -= 1;
    lock.releaseLock();
  }
}

function spreadsheet(): GoogleAppsScript.Spreadsheet.Spreadsheet {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;

  // Standalone deploy uchun: jadval ID `Script Properties` da saqlanadi.
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) {
    throw new AppError(
      'INTERNAL',
      'Ma’lumotlar jadvali topilmadi. Script Properties ichida SPREADSHEET_ID ni ko‘rsating.',
    );
  }
  return SpreadsheetApp.openById(id);
}

export function scriptTimeZone(): string {
  try {
    return spreadsheet().getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'UTC';
  } catch {
    return 'UTC';
  }
}

function serialize(value: unknown, column: Column<unknown>): string | number | boolean {
  if (value === null || value === undefined) return '';
  switch (column.type) {
    case 'json':
      return JSON.stringify(value);
    case 'boolean':
      return value === true;
    case 'number':
      return typeof value === 'number' ? value : Number(value);
    default:
      return String(value);
  }
}

function deserialize(raw: unknown, column: Column<unknown>): unknown {
  const isEmpty = raw === '' || raw === null || raw === undefined;

  switch (column.type) {
    case 'number': {
      if (isEmpty) return column.nullable ? null : (column.defaultValue ?? 0);
      const num = typeof raw === 'number' ? raw : Number(String(raw).replace(/\s/g, ''));
      return Number.isFinite(num) ? num : (column.defaultValue ?? 0);
    }
    case 'boolean': {
      if (isEmpty) return column.defaultValue ?? false;
      if (typeof raw === 'boolean') return raw;
      const text = String(raw).trim().toLowerCase();
      return text === 'true' || text === 'ha' || text === '1' || text === 'yes';
    }
    case 'json': {
      if (isEmpty) return column.defaultValue ?? null;
      try {
        return JSON.parse(String(raw));
      } catch {
        return column.defaultValue ?? null;
      }
    }
    case 'date':
    case 'timestamp':
    case 'string':
    default: {
      if (isEmpty) return column.nullable ? null : (column.defaultValue ?? '');
      if (raw instanceof Date) {
        return column.type === 'date'
          ? Utilities.formatDate(raw, 'UTC', 'yyyy-MM-dd')
          : raw.toISOString();
      }
      return String(raw);
    }
  }
}

interface LoadedTable<T> {
  rows: T[];
  /** id → jadvaldagi qator raqami (1-asosli). */
  rowNumberById: Map<string, number>;
  headers: string[];
}

/** Bitta bajarilish (execution) davomida takroriy o'qishning oldini oladi. */
const memo = new Map<string, unknown>();

export function clearMemo(sheetName?: string): void {
  if (sheetName) memo.delete(sheetName);
  else memo.clear();
}

export class Table<T extends object> {
  constructor(private readonly schema: TableSchema<T>) {}

  get sheetName(): string {
    return this.schema.sheetName;
  }

  /** Varaq va sarlavhalarni tekshiradi, yetishmayotganini qo'shadi. */
  ensure(): GoogleAppsScript.Spreadsheet.Sheet {
    const book = spreadsheet();
    let sheet = book.getSheetByName(this.schema.sheetName);
    const headers = this.schema.columns.map((column) => column.header);

    if (!sheet) {
      sheet = book.insertSheet(this.schema.sheetName);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      return sheet;
    }

    const lastColumn = Math.max(sheet.getLastColumn(), 1);
    const existing = sheet
      .getRange(1, 1, 1, lastColumn)
      .getValues()[0]!
      .map((value) => String(value).trim());

    const missing = headers.filter((header) => !existing.includes(header));
    if (missing.length > 0) {
      const startColumn = existing.filter((header) => header !== '').length + 1;
      sheet.getRange(1, startColumn, 1, missing.length).setValues([missing]);
      sheet.getRange(1, 1, 1, startColumn + missing.length - 1).setFontWeight('bold');
      clearMemo(this.schema.sheetName);
    }
    return sheet;
  }

  private load(): LoadedTable<T> {
    const cached = memo.get(this.schema.sheetName) as LoadedTable<T> | undefined;
    if (cached) return cached;

    const sheet = this.ensure();
    const lastRow = sheet.getLastRow();
    const lastColumn = Math.max(sheet.getLastColumn(), 1);
    const headerRow = sheet
      .getRange(1, 1, 1, lastColumn)
      .getValues()[0]!
      .map((value) => String(value).trim());

    const rows: T[] = [];
    const rowNumberById = new Map<string, number>();

    if (lastRow >= 2) {
      const values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
      for (let i = 0; i < values.length; i += 1) {
        const raw = values[i]!;
        const record: Record<string, unknown> = {};
        let hasAnyValue = false;

        for (const column of this.schema.columns) {
          const columnIndex = headerRow.indexOf(column.header);
          const cell = columnIndex === -1 ? '' : raw[columnIndex];
          if (cell !== '' && cell !== null && cell !== undefined) hasAnyValue = true;
          record[column.key] = deserialize(cell, column as Column<unknown>);
        }

        if (!hasAnyValue) continue; // Bo'sh qator — o'tkazib yuboramiz.
        const id = String(record[this.schema.idKey] ?? '');
        if (id === '') continue;

        rows.push(record as T);
        rowNumberById.set(id, i + 2);
      }
    }

    const loaded: LoadedTable<T> = { rows, rowNumberById, headers: headerRow };
    memo.set(this.schema.sheetName, loaded);
    return loaded;
  }

  all(): T[] {
    return this.load().rows;
  }

  findById(id: string): T | null {
    return this.load().rows.find((row) => String(row[this.schema.idKey]) === id) ?? null;
  }

  find(predicate: (row: T) => boolean): T | null {
    return this.load().rows.find(predicate) ?? null;
  }

  filter(predicate: (row: T) => boolean): T[] {
    return this.load().rows.filter(predicate);
  }

  count(): number {
    return this.load().rows.length;
  }

  private toRow(record: T, headers: string[]): unknown[] {
    const row = new Array<unknown>(headers.length).fill('');
    for (const column of this.schema.columns) {
      const index = headers.indexOf(column.header);
      if (index === -1) continue;
      row[index] = serialize(record[column.key], column as Column<unknown>);
    }
    return row;
  }

  insert(record: T): T {
    return withLock(() => {
      const sheet = this.ensure();
      const { headers } = this.load();
      sheet.appendRow(this.toRow(record, headers) as (string | number | boolean)[]);
      clearMemo(this.schema.sheetName);
      return record;
    });
  }

  insertMany(records: readonly T[]): void {
    if (records.length === 0) return;
    withLock(() => {
      const sheet = this.ensure();
      const { headers } = this.load();
      const rows = records.map((record) => this.toRow(record, headers));
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows as unknown[][]);
      clearMemo(this.schema.sheetName);
    });
  }

  update(id: string, patch: Partial<T>): T {
    return withLock(() => {
      const sheet = this.ensure();
      const loaded = this.load();
      const rowNumber = loaded.rowNumberById.get(id);
      const existing = loaded.rows.find((row) => String(row[this.schema.idKey]) === id);
      if (!rowNumber || !existing) throw new AppError('NOT_FOUND', 'Yozuv topilmadi');

      const merged = { ...existing, ...patch } as T;
      sheet
        .getRange(rowNumber, 1, 1, loaded.headers.length)
        .setValues([this.toRow(merged, loaded.headers) as unknown[]]);
      clearMemo(this.schema.sheetName);
      return merged;
    });
  }

  /** Bir nechta yozuvni bitta amalda yangilaydi (har biri uchun alohida `setValues` emas). */
  updateMany(patches: ReadonlyArray<{ id: string; patch: Partial<T> }>): void {
    if (patches.length === 0) return;
    withLock(() => {
      const sheet = this.ensure();
      const loaded = this.load();
      for (const { id, patch } of patches) {
        const rowNumber = loaded.rowNumberById.get(id);
        const existing = loaded.rows.find((row) => String(row[this.schema.idKey]) === id);
        if (!rowNumber || !existing) continue;
        const merged = { ...existing, ...patch } as T;
        sheet
          .getRange(rowNumber, 1, 1, loaded.headers.length)
          .setValues([this.toRow(merged, loaded.headers) as unknown[]]);
      }
      clearMemo(this.schema.sheetName);
    });
  }

  /** Qatorni butunlay o'chiradi. Tranzaksiyalar uchun ishlatilmaydi — u yerda yumshoq o'chirish. */
  deleteById(id: string): void {
    withLock(() => {
      const sheet = this.ensure();
      const rowNumber = this.load().rowNumberById.get(id);
      if (!rowNumber) throw new AppError('NOT_FOUND', 'Yozuv topilmadi');
      sheet.deleteRow(rowNumber);
      clearMemo(this.schema.sheetName);
    });
  }

  /** Pastdan yuqoriga o'chiradi — qator raqamlari surilib ketmaydi. */
  deleteWhere(predicate: (row: T) => boolean): number {
    return withLock(() => {
      const sheet = this.ensure();
      const loaded = this.load();
      const rowNumbers = loaded.rows
        .filter(predicate)
        .map((row) => loaded.rowNumberById.get(String(row[this.schema.idKey])))
        .filter((value): value is number => typeof value === 'number')
        .sort((a, b) => b - a);

      for (const rowNumber of rowNumbers) sheet.deleteRow(rowNumber);
      clearMemo(this.schema.sheetName);
      return rowNumbers.length;
    });
  }
}

export function newId(): string {
  return Utilities.getUuid();
}
