/**
 * Domen modeli — server va klient uchun yagona manba.
 *
 * Muhim qoidalar:
 *  - Har bir yozuv barqaror UUID `id` ga ega. Jadval qator raqami HECH QAChON
 *    identifikator sifatida ishlatilmaydi (v1 dagi "qator surilib ketdi" xatosi).
 *  - Pul har doim butun son — minor birlikda (tiyin/cent). Float yig'indi xatosi yo'q.
 *  - Sana har doim `YYYY-MM-DD` matn. `Date` obyekti saqlanmaydi → vaqt mintaqasi
 *    muammosi tug'ilmaydi.
 */

/** ISO 4217 kodi, masalan "UZS", "USD". */
export type CurrencyCode = string;

/** `YYYY-MM-DD` ko'rinishidagi kalendar sana. */
export type IsoDate = string;

/** ISO 8601 timestamp (UTC), masalan `2026-08-17T09:12:33.000Z`. */
export type IsoTimestamp = string;

/** Pul miqdori — minor birlikda (masalan 12.34 USD → 1234). */
export type MinorUnits = number;

// ---------------------------------------------------------------------------
// Ruxsatlar
// ---------------------------------------------------------------------------

export const RESOURCES = [
  'income',
  'expense',
  'transfer',
  'dashboard',
  'catalog',
  'users',
  'config',
  'audit',
] as const;

export type Resource = (typeof RESOURCES)[number];

export const ACTIONS = ['view', 'create', 'edit', 'delete'] as const;

export type Action = (typeof ACTIONS)[number];

export type Permissions = {
  [R in Resource]?: Partial<Record<Action, boolean>>;
};

export const ROLES = ['admin', 'manager', 'operator', 'viewer'] as const;

export type Role = (typeof ROLES)[number];

// ---------------------------------------------------------------------------
// Katalog (sozlanuvchi ma'lumotnomalar)
// ---------------------------------------------------------------------------

/** Xarajatlarni firmalar o'rtasida taqsimlash usuli. */
export const EXPENSE_ALLOCATIONS = ['direct', 'proRataIncome', 'equal'] as const;
export type ExpenseAllocation = (typeof EXPENSE_ALLOCATIONS)[number];

/**
 * Firma / biznes yo'nalishi. v1 da "Greenpen" va "Smartmiz" kodga qotirib
 * yozilgan edi; endi bu oddiy katalog yozuvi — dastur ichida qo'shiladi.
 */
export interface Firm {
  id: string;
  name: string;
  active: boolean;
  /** Firmaga bevosita bog'lanmagan (umumiy) xarajatlar qanday taqsimlanadi. */
  expenseAllocation: ExpenseAllocation;
  sortOrder: number;
}

/** Hisob / to'lov turi (Naqd, P2P, Uzcard, Bank, USD kassa, ...). */
export interface Account {
  id: string;
  name: string;
  currency: CurrencyCode;
  active: boolean;
  /** `false` bo'lsa hisob balans panelida ko'rsatilmaydi (lekin hisoblanadi). */
  showInBalance: boolean;
  sortOrder: number;
}

/** Savdo nuqtasi. Firma bilan bog'lanish nom ichidagi matn orqali emas, `firmId` orqali. */
export interface Point {
  id: string;
  name: string;
  firmId: string;
  active: boolean;
  /**
   * `true` bo'lsa bu nuqta tushumi daromad statistikasiga kirmaydi
   * (v1 dagi qotirib yozilgan "DONIYOR AKA / KASSA / DIREKTOR" ro'yxati o'rniga).
   */
  excludeFromRevenue: boolean;
  sortOrder: number;
}

export const CATEGORY_KINDS = ['income', 'expense', 'both'] as const;
export type CategoryKind = (typeof CATEGORY_KINDS)[number];

/** Kirim/chiqim kategoriyasi. */
export interface Category {
  id: string;
  name: string;
  kind: CategoryKind;
  active: boolean;
  sortOrder: number;
}

export type CatalogEntity = Firm | Account | Point | Category;

export interface Catalog {
  firms: Firm[];
  accounts: Account[];
  points: Point[];
  categories: Category[];
}

// ---------------------------------------------------------------------------
// Foydalanuvchilar
// ---------------------------------------------------------------------------

export const USER_STATUSES = ['active', 'blocked'] as const;
export type UserStatus = (typeof USER_STATUSES)[number];

/** Klientga yuboriladigan foydalanuvchi — parol maydonlari YO'Q. */
export interface PublicUser {
  id: string;
  login: string;
  displayName: string;
  role: Role;
  status: UserStatus;
  permissions: Permissions;
  createdAt: IsoTimestamp;
  lastLoginAt: IsoTimestamp | null;
}

// ---------------------------------------------------------------------------
// Tranzaksiyalar
// ---------------------------------------------------------------------------

export const TRANSACTION_KINDS = ['income', 'expense', 'transfer'] as const;
export type TransactionKind = (typeof TRANSACTION_KINDS)[number];

/**
 * Bitta pul harakati.
 *
 * `amountMinor` HAR DOIM `accountId` hisobining o'z valyutasida saqlanadi.
 * Bu v1 dagi eng og'riqli xatoni butunlay yo'q qiladi: dollar chiqimini
 * tahrirlaganda summa qayta kursga ko'paytirilmaydi, chunki UZS ekvivalenti
 * alohida (`baseAmountMinor`) va faqat hisobot uchun saqlanadi.
 */
export interface Transaction {
  id: string;
  /** Bitta amalda kiritilgan qatorlarni bog'laydi (masalan: Naqd + P2P tushumi). */
  batchId: string;
  date: IsoDate;
  kind: TransactionKind;

  firmId: string | null;
  pointId: string | null;
  categoryId: string | null;

  /** income: pul kirgan hisob. expense/transfer: pul chiqqan hisob. */
  accountId: string;
  amountMinor: MinorUnits;
  currency: CurrencyCode;

  /** Faqat `transfer` uchun: pul tushgan hisob. */
  counterAccountId: string | null;
  counterAmountMinor: MinorUnits | null;
  counterCurrency: CurrencyCode | null;

  /** Hisobot valyutasiga (base) o'tkazish uchun ishlatilgan kurs. */
  fxRate: number;
  /** Hisobot valyutasidagi ekvivalent — faqat statistika uchun. */
  baseAmountMinor: MinorUnits;

  note: string;

  createdAt: IsoTimestamp;
  createdBy: string;
  updatedAt: IsoTimestamp;
  updatedBy: string;
  /** Yumshoq o'chirish — moliyaviy tarix hech qachon yo'qolmaydi. */
  deletedAt: IsoTimestamp | null;
  deletedBy: string | null;
}

/** Bir amalda kiritilgan tranzaksiyalar guruhi (UI uchun). */
export interface TransactionBatch {
  batchId: string;
  date: IsoDate;
  kind: TransactionKind;
  firmId: string | null;
  pointId: string | null;
  categoryId: string | null;
  note: string;
  entries: Transaction[];
  updatedAt: IsoTimestamp;
  updatedBy: string;
}

// ---------------------------------------------------------------------------
// Konfiguratsiya
// ---------------------------------------------------------------------------

export interface AppConfig {
  /** Hisobot valyutasi — barcha umumlashtirilgan ko'rsatkichlar shunda. */
  baseCurrency: CurrencyCode;
  organizationName: string;
  /** Sessiya amal qilish muddati (daqiqa). */
  sessionTtlMinutes: number;
  /** Ketma-ket muvaffaqiyatsiz urinishlar soni — undan keyin login vaqtincha bloklanadi. */
  maxLoginAttempts: number;
  loginLockoutMinutes: number;
  /** Bitta tranzaksiya uchun maksimal summa (major birlikda). */
  maxTransactionAmount: number;
  maxNoteLength: number;
  /** Umumiy (firmasiz) xarajatlarni taqsimlash usuli. */
  sharedExpenseAllocation: ExpenseAllocation;
  /** Valyuta kursini avtomatik olish uchun manba. `none` — faqat qo'lda. */
  fxProvider: 'cbu.uz' | 'none';
}

export interface CurrencyMeta {
  code: CurrencyCode;
  decimals: number;
  symbol: string;
}

// ---------------------------------------------------------------------------
// Hisobotlar
// ---------------------------------------------------------------------------

export interface AccountBalance {
  accountId: string;
  name: string;
  currency: CurrencyCode;
  balanceMinor: MinorUnits;
  /** Hisobot valyutasidagi taxminiy ekvivalent (joriy kurs bo'yicha). */
  baseBalanceMinor: MinorUnits;
  showInBalance: boolean;
  sortOrder: number;
}

export interface PeriodTotals {
  incomeMinor: MinorUnits;
  expenseMinor: MinorUnits;
  profitMinor: MinorUnits;
}

export interface TrendPoint {
  key: string;
  incomeMinor: MinorUnits;
  expenseMinor: MinorUnits;
}

export interface NamedTotal {
  id: string;
  name: string;
  totalMinor: MinorUnits;
}

export interface DashboardStats {
  baseCurrency: CurrencyCode;
  current: PeriodTotals;
  previous: PeriodTotals;
  trend: TrendPoint[];
  byFirm: NamedTotal[];
  byPoint: NamedTotal[];
  byCategory: NamedTotal[];
  byAccount: NamedTotal[];
  transactionCount: number;
}

export interface DateRange {
  start: IsoDate;
  end: IsoDate;
}

export interface TransactionFilter {
  range?: DateRange;
  kinds?: TransactionKind[];
  firmId?: string | null;
  pointId?: string | null;
  accountId?: string | null;
  categoryId?: string | null;
  search?: string;
  includeDeleted?: boolean;
}

export interface Page<T> {
  items: T[];
  total: number;
  offset: number;
  limit: number;
}

// ---------------------------------------------------------------------------
// Audit
// ---------------------------------------------------------------------------

export interface AuditEntry {
  id: string;
  at: IsoTimestamp;
  userId: string;
  userLogin: string;
  action: string;
  entity: string;
  entityId: string;
  details: string;
}
