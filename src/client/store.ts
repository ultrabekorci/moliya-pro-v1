/**
 * Klient holati.
 *
 * MUHIM: `sessionStorage` da FAQAT token saqlanadi. Rol va ruxsatlar hech qachon
 * brauzerda saqlanmaydi — ular har safar serverdan keladi. v1 da rol
 * `sessionStorage.moliya_role` da turardi va uni DevTools orqali "Admin" qilib
 * qo'yish kifoya edi.
 */

import type {
  Account,
  AccountBalance,
  AppConfig,
  Catalog,
  Category,
  Firm,
  IsoDate,
  Permissions,
  Point,
  PublicUser,
} from '../shared/types.js';
import type { Bootstrap, SessionInfo } from '../shared/api.js';
import { can } from '../shared/permissions.js';
import type { Action, Resource } from '../shared/types.js';
import { setToken } from './api.js';

const TOKEN_KEY = 'moliya.token';

export interface AppState {
  user: PublicUser | null;
  permissions: Permissions | null;
  config: AppConfig | null;
  catalog: Catalog;
  balances: AccountBalance[];
  rates: Record<string, number>;
  today: IsoDate;
  timeZone: string;
  expiresAt: number;
}

const emptyCatalog: Catalog = { firms: [], accounts: [], points: [], categories: [] };

export const state: AppState = {
  user: null,
  permissions: null,
  config: null,
  catalog: emptyCatalog,
  balances: [],
  rates: {},
  today: new Date().toISOString().slice(0, 10),
  timeZone: 'UTC',
  expiresAt: 0,
};

type Listener = () => void;
const listeners = new Set<Listener>();

export function subscribe(listener: Listener): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

export function notify(): void {
  for (const listener of listeners) listener();
}

export function loadStoredToken(): string | null {
  try {
    return window.sessionStorage.getItem(TOKEN_KEY);
  } catch {
    return null;
  }
}

function storeToken(token: string | null): void {
  try {
    if (token) window.sessionStorage.setItem(TOKEN_KEY, token);
    else window.sessionStorage.removeItem(TOKEN_KEY);
  } catch {
    // Maxfiylik rejimida sessionStorage ishlamasligi mumkin — bu jiddiy emas.
  }
}

export function applySession(session: SessionInfo): void {
  setToken(session.token);
  storeToken(session.token);
  state.user = session.user;
  state.permissions = session.permissions;
  state.expiresAt = Date.parse(session.expiresAt);
  notify();
}

export function applyBootstrap(bootstrap: Bootstrap): void {
  state.config = bootstrap.config;
  state.catalog = bootstrap.catalog;
  state.balances = bootstrap.balances;
  state.rates = bootstrap.rates;
  state.today = bootstrap.today;
  state.timeZone = bootstrap.serverTimeZone;
  notify();
}

export function clearSession(): void {
  setToken(null);
  storeToken(null);
  state.user = null;
  state.permissions = null;
  state.balances = [];
  state.expiresAt = 0;
  notify();
}

export function allowed(resource: Resource, action: Action): boolean {
  return can(state.permissions, resource, action);
}

export function baseCurrency(): string {
  return state.config?.baseCurrency ?? 'UZS';
}

// --- Katalogdan qidirish -----------------------------------------------------

export function accountById(id: string | null | undefined): Account | null {
  if (!id) return null;
  return state.catalog.accounts.find((account) => account.id === id) ?? null;
}

export function pointById(id: string | null | undefined): Point | null {
  if (!id) return null;
  return state.catalog.points.find((point) => point.id === id) ?? null;
}

export function categoryById(id: string | null | undefined): Category | null {
  if (!id) return null;
  return state.catalog.categories.find((category) => category.id === id) ?? null;
}

export function firmById(id: string | null | undefined): Firm | null {
  if (!id) return null;
  return state.catalog.firms.find((firm) => firm.id === id) ?? null;
}

export function nameOf(id: string | null | undefined): string {
  if (!id) return '—';
  return (
    accountById(id)?.name ??
    pointById(id)?.name ??
    categoryById(id)?.name ??
    firmById(id)?.name ??
    '—'
  );
}

export function activeAccounts(): Account[] {
  return state.catalog.accounts.filter((account) => account.active);
}

export function activePoints(): Point[] {
  return state.catalog.points.filter((point) => point.active);
}

export function activeFirms(): Firm[] {
  return state.catalog.firms.filter((firm) => firm.active);
}

export function activeCategories(kind: 'income' | 'expense'): Category[] {
  return state.catalog.categories.filter(
    (category) => category.active && (category.kind === kind || category.kind === 'both'),
  );
}

/** Hisob valyutasi — summani to'g'ri formatlash uchun. */
export function currencyOfAccount(id: string | null | undefined): string {
  return accountById(id)?.currency ?? baseCurrency();
}
