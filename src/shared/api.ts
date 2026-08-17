/**
 * Klient ↔ server shartnomasi.
 *
 * v1 da har bir server funksiyasi alohida global funksiya edi va istalgan
 * odam ularni `google.script.run.deleteUser(3)` deb chaqira olardi. Endi
 * yagona kirish nuqtasi bor: `rpc(request)`. Har bir metod uchun kerakli
 * ruxsat server tomonda e'lon qilingan (`src/server/router.ts`) va token
 * tekshiruvidan o'tmagan chaqiruv umuman bajarilmaydi.
 */

import type {
  AccountBalance,
  AppConfig,
  AuditEntry,
  Catalog,
  CurrencyCode,
  DashboardStats,
  DateRange,
  IsoDate,
  MinorUnits,
  Page,
  Permissions,
  PublicUser,
  Role,
  Transaction,
  TransactionBatch,
  TransactionFilter,
  TransactionKind,
  UserStatus,
} from './types.js';

export interface SessionInfo {
  token: string;
  expiresAt: string;
  user: PublicUser;
  permissions: Permissions;
}

/** Ilova ishga tushganda kerak bo'ladigan hamma narsa — bitta so'rovda. */
export interface Bootstrap {
  config: AppConfig;
  catalog: Catalog;
  balances: AccountBalance[];
  rates: Record<string, number>;
  today: IsoDate;
  serverTimeZone: string;
}

export interface AmountInput {
  accountId: string;
  /** Foydalanuvchi kiritgan matn — server o'zi minor birlikka aylantiradi. */
  amount: string;
}

export interface SaveTransactionInput {
  /** Mavjud guruhni tahrirlash uchun. Bo'sh bo'lsa — yangi yozuv. */
  batchId?: string | null;
  kind: TransactionKind;
  date: IsoDate;
  pointId?: string | null;
  firmId?: string | null;
  categoryId?: string | null;
  note?: string;
  /** `income` uchun: bir nechta hisobga bir vaqtda tushum. */
  entries?: AmountInput[];
  /** `expense` va `transfer` uchun: pul chiqadigan hisob va summa. */
  accountId?: string;
  amount?: string;
  /** Faqat `transfer`: pul tushadigan hisob va u yerdagi summa. */
  counterAccountId?: string | null;
  counterAmount?: string | null;
}

export interface CatalogItemInput {
  id?: string | null;
  name: string;
  active?: boolean;
  sortOrder?: number;
}

export interface FirmInput extends CatalogItemInput {
  expenseAllocation?: AppConfig['sharedExpenseAllocation'];
}

export interface AccountInput extends CatalogItemInput {
  currency: CurrencyCode;
  showInBalance?: boolean;
}

export interface PointInput extends CatalogItemInput {
  firmId: string;
  excludeFromRevenue?: boolean;
}

export interface CategoryInput extends CatalogItemInput {
  kind: 'income' | 'expense' | 'both';
}

export interface UserInput {
  id?: string | null;
  login: string;
  displayName?: string;
  /** Bo'sh qoldirilsa mavjud parol o'zgarmaydi (v1 da bu mumkin emas edi). */
  password?: string | null;
  role: Role;
  status?: UserStatus;
  permissions?: Permissions;
}

export interface DashboardQuery {
  range: DateRange;
  previousRange: DateRange;
  firmId?: string | null;
}

export interface BalanceQuery {
  asOf?: IsoDate | null;
}

export interface FxQuery {
  date: IsoDate;
  currency: CurrencyCode;
}

export interface FxRateResult {
  currency: CurrencyCode;
  rate: number;
  /** Kurs aslida qaysi kunga tegishli (dam olish kunlarida oldingi kun bo'lishi mumkin). */
  effectiveDate: IsoDate;
  source: string;
}

export interface ManualFxInput {
  date: IsoDate;
  currency: CurrencyCode;
  rate: number;
}

export interface ListTransactionsQuery {
  filter?: TransactionFilter;
  offset?: number;
  limit?: number;
  /** Guruhlangan ko'rinish (Kirim jadvali uchun) yoki tekis ro'yxat. */
  grouped?: boolean;
}

export interface DeleteInput {
  id: string;
}

export interface ChangePasswordInput {
  currentPassword: string;
  newPassword: string;
}

/** Metod nomi → (parametr, natija) juftligi. */
/** Login ekranida kerak bo'ladigan, maxfiy bo'lmagan holat. */
export interface AppStatus {
  /** `false` — hali birorta foydalanuvchi yo'q, birinchi administrator yaratiladi. */
  configured: boolean;
  organizationName: string;
}

export interface ApiContract {
  'app.status': { params: Record<string, never>; result: AppStatus };
  'auth.setup': { params: { login: string; password: string }; result: SessionInfo };
  'auth.login': { params: { login: string; password: string }; result: SessionInfo };
  'auth.logout': { params: Record<string, never>; result: { ok: true } };
  'auth.session': { params: Record<string, never>; result: SessionInfo };
  'auth.changePassword': { params: ChangePasswordInput; result: { ok: true } };

  'app.bootstrap': { params: Record<string, never>; result: Bootstrap };

  'tx.list': { params: ListTransactionsQuery; result: Page<Transaction> };
  'tx.listGrouped': { params: ListTransactionsQuery; result: Page<TransactionBatch> };
  'tx.save': { params: SaveTransactionInput; result: { batchId: string } };
  'tx.delete': { params: DeleteInput; result: { ok: true } };
  'tx.balances': { params: BalanceQuery; result: AccountBalance[] };
  'tx.dashboard': { params: DashboardQuery; result: DashboardStats };

  'catalog.get': { params: Record<string, never>; result: Catalog };
  'catalog.saveFirm': { params: FirmInput; result: { id: string } };
  'catalog.saveAccount': { params: AccountInput; result: { id: string } };
  'catalog.savePoint': { params: PointInput; result: { id: string } };
  'catalog.saveCategory': { params: CategoryInput; result: { id: string } };
  'catalog.deleteFirm': { params: DeleteInput; result: { ok: true } };
  'catalog.deleteAccount': { params: DeleteInput; result: { ok: true } };
  'catalog.deletePoint': { params: DeleteInput; result: { ok: true } };
  'catalog.deleteCategory': { params: DeleteInput; result: { ok: true } };

  'users.list': { params: Record<string, never>; result: PublicUser[] };
  'users.save': { params: UserInput; result: { id: string } };
  'users.delete': { params: DeleteInput; result: { ok: true } };

  'config.get': { params: Record<string, never>; result: AppConfig };
  'config.save': { params: Partial<AppConfig>; result: AppConfig };

  'fx.rate': { params: FxQuery; result: FxRateResult };
  'fx.setManual': { params: ManualFxInput; result: FxRateResult };

  'audit.list': { params: { offset?: number; limit?: number }; result: Page<AuditEntry> };
}

export type ApiMethod = keyof ApiContract;
export type ApiParams<M extends ApiMethod> = ApiContract[M]['params'];
export type ApiResult<M extends ApiMethod> = ApiContract[M]['result'];

export interface RpcRequest<M extends ApiMethod = ApiMethod> {
  method: M;
  params: ApiParams<M>;
  token: string | null;
}

export type RpcResponse<M extends ApiMethod = ApiMethod> =
  | { ok: true; data: ApiResult<M> }
  | { ok: false; error: RpcError };

export const ERROR_CODES = [
  'UNAUTHENTICATED',
  'FORBIDDEN',
  'VALIDATION',
  'NOT_FOUND',
  'CONFLICT',
  'RATE_LIMITED',
  'INTERNAL',
] as const;

export type ErrorCode = (typeof ERROR_CODES)[number];

export interface RpcError {
  code: ErrorCode;
  message: string;
  /** Qaysi maydon xato — forma validatsiyasi uchun. */
  field?: string;
}

/** Klientda ham, serverda ham ishlatiladigan xato sinfi. */
export class AppError extends Error {
  readonly code: ErrorCode;
  readonly field?: string;

  constructor(code: ErrorCode, message: string, field?: string) {
    super(message);
    this.name = 'AppError';
    this.code = code;
    if (field !== undefined) this.field = field;
  }

  toRpcError(): RpcError {
    return this.field === undefined
      ? { code: this.code, message: this.message }
      : { code: this.code, message: this.message, field: this.field };
  }
}

export function validationError(message: string, field?: string): AppError {
  return new AppError('VALIDATION', message, field);
}

/** Hisobot valyutasidagi summani ko'rsatish uchun yordamchi tur. */
export interface BaseAmount {
  minor: MinorUnits;
  currency: CurrencyCode;
}
