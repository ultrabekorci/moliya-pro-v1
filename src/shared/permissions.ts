/**
 * Ruxsatlar modeli.
 *
 * Asosiy tamoyil — FAIL-CLOSED: ruxsat aniq berilmagan bo'lsa, u YO'Q.
 * v1 dagi `if (!permissions || permissions.admin) { hammasi ochiq }` mantig'i
 * teskari edi va huquqlar yuklanmaganda foydalanuvchiga admin huquqini berardi.
 */

import { ACTIONS, RESOURCES } from './types.js';
import type { Action, Permissions, Resource, Role } from './types.js';

export function isResource(value: unknown): value is Resource {
  return typeof value === 'string' && (RESOURCES as readonly string[]).includes(value);
}

export function isAction(value: unknown): value is Action {
  return typeof value === 'string' && (ACTIONS as readonly string[]).includes(value);
}

function grant(resources: readonly Resource[], actions: readonly Action[]): Permissions {
  const result: Permissions = {};
  for (const resource of resources) {
    const entry: Partial<Record<Action, boolean>> = {};
    for (const action of actions) entry[action] = true;
    result[resource] = entry;
  }
  return result;
}

const ALL_ACTIONS: readonly Action[] = ACTIONS;

/** Rol bo'yicha standart ruxsatlar. Foydalanuvchi darajasidagi sozlamalar shu ustiga qo'yiladi. */
export const ROLE_DEFAULTS: Readonly<Record<Role, Permissions>> = {
  admin: grant(RESOURCES, ALL_ACTIONS),
  manager: {
    ...grant(['income', 'expense', 'transfer'], ALL_ACTIONS),
    ...grant(['dashboard', 'catalog', 'audit'], ['view']),
    ...grant(['catalog'], ['view', 'create', 'edit']),
  },
  operator: {
    ...grant(['income', 'expense', 'transfer'], ['view', 'create']),
    ...grant(['catalog'], ['view']),
  },
  viewer: {
    ...grant(['income', 'expense', 'transfer', 'dashboard', 'catalog'], ['view']),
  },
};

/** Ikki ruxsat to'plamini birlashtiradi (`override` ustun turadi). */
export function mergePermissions(base: Permissions, override: Permissions): Permissions {
  const result: Permissions = {};
  for (const resource of RESOURCES) {
    const baseEntry = base[resource];
    const overrideEntry = override[resource];
    if (!baseEntry && !overrideEntry) continue;
    result[resource] = { ...(baseEntry ?? {}), ...(overrideEntry ?? {}) };
  }
  return result;
}

/**
 * Foydalanuvchining haqiqiy ruxsatlari: rol standarti + shaxsiy sozlamalar.
 * `admin` roli har doim to'liq huquqqa ega (o'zini qulflab qo'yishning oldi olinadi).
 */
export function effectivePermissions(role: Role, overrides: Permissions | null | undefined): Permissions {
  if (role === 'admin') return ROLE_DEFAULTS.admin;
  const base = ROLE_DEFAULTS[role] ?? {};
  return mergePermissions(base, overrides ?? {});
}

/** Yagona tekshiruv nuqtasi. Server ham, klient ham shu funksiyani chaqiradi. */
export function can(permissions: Permissions | null | undefined, resource: Resource, action: Action): boolean {
  if (!permissions) return false;
  const entry = permissions[resource];
  if (!entry) return false;
  return entry[action] === true;
}

/** Berilgan resurslardan kamida bittasida `action` ruxsati bormi. */
export function canAny(
  permissions: Permissions | null | undefined,
  resources: readonly Resource[],
  action: Action,
): boolean {
  return resources.some((resource) => can(permissions, resource, action));
}

/** Tranzaksiya turi uchun mos resurs nomi. */
export function resourceForKind(kind: 'income' | 'expense' | 'transfer'): Resource {
  return kind;
}

/** Noma'lum kalitlarni tashlab, ruxsat obyektini xavfsiz normallashtiradi. */
export function sanitizePermissions(input: unknown): Permissions {
  const result: Permissions = {};
  if (typeof input !== 'object' || input === null) return result;

  for (const [resourceKey, actionsValue] of Object.entries(input as Record<string, unknown>)) {
    if (!isResource(resourceKey)) continue;
    if (typeof actionsValue !== 'object' || actionsValue === null) continue;

    const entry: Partial<Record<Action, boolean>> = {};
    for (const [actionKey, allowed] of Object.entries(actionsValue as Record<string, unknown>)) {
      if (!isAction(actionKey)) continue;
      if (allowed === true) entry[actionKey] = true;
    }
    if (Object.keys(entry).length > 0) result[resourceKey] = entry;
  }
  return result;
}

export const RESOURCE_LABELS: Readonly<Record<Resource, string>> = {
  income: 'Kirim',
  expense: 'Chiqim',
  transfer: "O'tkazma",
  dashboard: 'Dashboard',
  catalog: 'Ma’lumotnomalar',
  users: 'Xodimlar',
  config: 'Sozlamalar',
  audit: 'Amallar tarixi',
};

export const ACTION_LABELS: Readonly<Record<Action, string>> = {
  view: "Ko'rish",
  create: "Qo'shish",
  edit: 'Tahrirlash',
  delete: "O'chirish",
};

export const ROLE_LABELS: Readonly<Record<Role, string>> = {
  admin: 'Administrator',
  manager: 'Menejer',
  operator: 'Operator',
  viewer: 'Kuzatuvchi',
};
