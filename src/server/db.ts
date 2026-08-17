/**
 * Jadval sxemalari. Bu fayl — ma'lumotlar bazasi strukturasining yagona ta'rifi
 * (v1 da sxema hech qayerda hujjatlashtirilmagan edi).
 */

import { Table } from './sheets.js';
import type { Column, TableSchema } from './sheets.js';
import type {
  Account,
  Category,
  Firm,
  Permissions,
  Point,
  Role,
  Transaction,
  UserStatus,
} from '../shared/types.js';

/** Parol maydonlari bilan birga saqlanadigan foydalanuvchi. Klientga hech qachon chiqmaydi. */
export interface StoredUser {
  id: string;
  login: string;
  displayName: string;
  passwordHash: string;
  passwordSalt: string;
  passwordIterations: number;
  role: Role;
  status: UserStatus;
  permissions: Permissions | null;
  createdAt: string;
  updatedAt: string;
  lastLoginAt: string | null;
}

export interface StoredFxRate {
  id: string;
  date: string;
  currency: string;
  rate: number;
  source: string;
  fetchedAt: string;
}

export interface StoredAudit {
  id: string;
  at: string;
  userId: string;
  userLogin: string;
  action: string;
  entity: string;
  entityId: string;
  details: string;
}

export interface StoredConfig {
  id: string;
  value: string;
}

function col<T>(
  key: Extract<keyof T, string>,
  header: string,
  type: Column<T>['type'],
  extra: Partial<Column<T>> = {},
): Column<T> {
  return { key, header, type, ...extra };
}

const firmSchema: TableSchema<Firm> = {
  sheetName: 'Firms',
  idKey: 'id',
  columns: [
    col<Firm>('id', 'id', 'string'),
    col<Firm>('name', 'name', 'string'),
    col<Firm>('active', 'active', 'boolean', { defaultValue: true }),
    col<Firm>('expenseAllocation', 'expenseAllocation', 'string', { defaultValue: 'proRataIncome' }),
    col<Firm>('sortOrder', 'sortOrder', 'number', { defaultValue: 0 }),
  ],
};

const accountSchema: TableSchema<Account> = {
  sheetName: 'Accounts',
  idKey: 'id',
  columns: [
    col<Account>('id', 'id', 'string'),
    col<Account>('name', 'name', 'string'),
    col<Account>('currency', 'currency', 'string', { defaultValue: 'UZS' }),
    col<Account>('active', 'active', 'boolean', { defaultValue: true }),
    col<Account>('showInBalance', 'showInBalance', 'boolean', { defaultValue: true }),
    col<Account>('sortOrder', 'sortOrder', 'number', { defaultValue: 0 }),
  ],
};

const pointSchema: TableSchema<Point> = {
  sheetName: 'Points',
  idKey: 'id',
  columns: [
    col<Point>('id', 'id', 'string'),
    col<Point>('name', 'name', 'string'),
    col<Point>('firmId', 'firmId', 'string'),
    col<Point>('active', 'active', 'boolean', { defaultValue: true }),
    col<Point>('excludeFromRevenue', 'excludeFromRevenue', 'boolean', { defaultValue: false }),
    col<Point>('sortOrder', 'sortOrder', 'number', { defaultValue: 0 }),
  ],
};

const categorySchema: TableSchema<Category> = {
  sheetName: 'Categories',
  idKey: 'id',
  columns: [
    col<Category>('id', 'id', 'string'),
    col<Category>('name', 'name', 'string'),
    col<Category>('kind', 'kind', 'string', { defaultValue: 'expense' }),
    col<Category>('active', 'active', 'boolean', { defaultValue: true }),
    col<Category>('sortOrder', 'sortOrder', 'number', { defaultValue: 0 }),
  ],
};

const userSchema: TableSchema<StoredUser> = {
  sheetName: 'Users',
  idKey: 'id',
  columns: [
    col<StoredUser>('id', 'id', 'string'),
    col<StoredUser>('login', 'login', 'string'),
    col<StoredUser>('displayName', 'displayName', 'string'),
    col<StoredUser>('passwordHash', 'passwordHash', 'string'),
    col<StoredUser>('passwordSalt', 'passwordSalt', 'string'),
    col<StoredUser>('passwordIterations', 'passwordIterations', 'number', { defaultValue: 0 }),
    col<StoredUser>('role', 'role', 'string', { defaultValue: 'viewer' }),
    col<StoredUser>('status', 'status', 'string', { defaultValue: 'active' }),
    col<StoredUser>('permissions', 'permissions', 'json', { nullable: true }),
    col<StoredUser>('createdAt', 'createdAt', 'timestamp'),
    col<StoredUser>('updatedAt', 'updatedAt', 'timestamp'),
    col<StoredUser>('lastLoginAt', 'lastLoginAt', 'timestamp', { nullable: true }),
  ],
};

const transactionSchema: TableSchema<Transaction> = {
  sheetName: 'Transactions',
  idKey: 'id',
  columns: [
    col<Transaction>('id', 'id', 'string'),
    col<Transaction>('batchId', 'batchId', 'string'),
    col<Transaction>('date', 'date', 'date'),
    col<Transaction>('kind', 'kind', 'string'),
    col<Transaction>('firmId', 'firmId', 'string', { nullable: true }),
    col<Transaction>('pointId', 'pointId', 'string', { nullable: true }),
    col<Transaction>('categoryId', 'categoryId', 'string', { nullable: true }),
    col<Transaction>('accountId', 'accountId', 'string'),
    col<Transaction>('amountMinor', 'amountMinor', 'number', { defaultValue: 0 }),
    col<Transaction>('currency', 'currency', 'string'),
    col<Transaction>('counterAccountId', 'counterAccountId', 'string', { nullable: true }),
    col<Transaction>('counterAmountMinor', 'counterAmountMinor', 'number', { nullable: true }),
    col<Transaction>('counterCurrency', 'counterCurrency', 'string', { nullable: true }),
    col<Transaction>('fxRate', 'fxRate', 'number', { defaultValue: 1 }),
    col<Transaction>('baseAmountMinor', 'baseAmountMinor', 'number', { defaultValue: 0 }),
    col<Transaction>('note', 'note', 'string', { defaultValue: '' }),
    col<Transaction>('createdAt', 'createdAt', 'timestamp'),
    col<Transaction>('createdBy', 'createdBy', 'string'),
    col<Transaction>('updatedAt', 'updatedAt', 'timestamp'),
    col<Transaction>('updatedBy', 'updatedBy', 'string'),
    col<Transaction>('deletedAt', 'deletedAt', 'timestamp', { nullable: true }),
    col<Transaction>('deletedBy', 'deletedBy', 'string', { nullable: true }),
  ],
};

const fxSchema: TableSchema<StoredFxRate> = {
  sheetName: 'FxRates',
  idKey: 'id',
  columns: [
    col<StoredFxRate>('id', 'id', 'string'),
    col<StoredFxRate>('date', 'date', 'date'),
    col<StoredFxRate>('currency', 'currency', 'string'),
    col<StoredFxRate>('rate', 'rate', 'number', { defaultValue: 0 }),
    col<StoredFxRate>('source', 'source', 'string'),
    col<StoredFxRate>('fetchedAt', 'fetchedAt', 'timestamp'),
  ],
};

const auditSchema: TableSchema<StoredAudit> = {
  sheetName: 'AuditLog',
  idKey: 'id',
  columns: [
    col<StoredAudit>('id', 'id', 'string'),
    col<StoredAudit>('at', 'at', 'timestamp'),
    col<StoredAudit>('userId', 'userId', 'string'),
    col<StoredAudit>('userLogin', 'userLogin', 'string'),
    col<StoredAudit>('action', 'action', 'string'),
    col<StoredAudit>('entity', 'entity', 'string'),
    col<StoredAudit>('entityId', 'entityId', 'string'),
    col<StoredAudit>('details', 'details', 'string'),
  ],
};

const configSchema: TableSchema<StoredConfig> = {
  sheetName: 'Config',
  idKey: 'id',
  columns: [
    col<StoredConfig>('id', 'key', 'string'),
    col<StoredConfig>('value', 'value', 'string'),
  ],
};

export const firmsTable = new Table<Firm>(firmSchema);
export const accountsTable = new Table<Account>(accountSchema);
export const pointsTable = new Table<Point>(pointSchema);
export const categoriesTable = new Table<Category>(categorySchema);
export const usersTable = new Table<StoredUser>(userSchema);
export const transactionsTable = new Table<Transaction>(transactionSchema);
export const fxTable = new Table<StoredFxRate>(fxSchema);
export const auditTable = new Table<StoredAudit>(auditSchema);
export const configTable = new Table<StoredConfig>(configSchema);

export const ALL_TABLES = [
  configTable,
  firmsTable,
  accountsTable,
  pointsTable,
  categoriesTable,
  usersTable,
  transactionsTable,
  fxTable,
  auditTable,
];

/** Barcha varaqlar mavjudligini ta'minlaydi (birinchi ishga tushirish / migratsiya). */
export function ensureSchema(): void {
  for (const table of ALL_TABLES) table.ensure();
}
