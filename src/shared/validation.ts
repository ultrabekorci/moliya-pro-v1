/**
 * Validatsiya — klient ham, server ham AYNI SHU qoidalarni ishlatadi.
 *
 * v1 da tekshiruvlar faqat brauzerda edi, ya'ni ular umuman himoya emas edi.
 * Bu yerdagi funksiyalar serverda majburiy chaqiriladi; klient esa faqat
 * foydalanuvchiga tezroq xabar berish uchun ishlatadi.
 */

import { isIsoDate } from './dates.js';
import { validationError } from './api.js';
import type { CurrencyCode, Role, TransactionKind, UserStatus } from './types.js';
import { ROLES, TRANSACTION_KINDS, USER_STATUSES } from './types.js';

export const LIMITS = {
  loginMin: 3,
  loginMax: 32,
  passwordMin: 8,
  passwordMax: 128,
  nameMin: 1,
  nameMax: 60,
  noteMax: 500,
  currencyLength: 3,
  pageMax: 500,
} as const;

export function requireString(value: unknown, field: string, options: { min?: number; max?: number } = {}): string {
  if (typeof value !== 'string') throw validationError(`«${field}» matn bo‘lishi kerak`, field);
  const trimmed = value.trim();
  const min = options.min ?? 0;
  const max = options.max ?? Number.MAX_SAFE_INTEGER;
  if (trimmed.length < min) throw validationError(`«${field}» kamida ${min} ta belgidan iborat bo‘lsin`, field);
  if (trimmed.length > max) throw validationError(`«${field}» ${max} ta belgidan oshmasin`, field);
  return trimmed;
}

export function optionalString(value: unknown, field: string, max: number): string {
  if (value === undefined || value === null || value === '') return '';
  return requireString(value, field, { max });
}

export function requireIsoDate(value: unknown, field: string): string {
  if (typeof value !== 'string' || !isIsoDate(value)) {
    throw validationError(`«${field}» sanasi noto‘g‘ri (YYYY-MM-DD kutiladi)`, field);
  }
  return value;
}

export function requireBoolean(value: unknown, fallback: boolean): boolean {
  return typeof value === 'boolean' ? value : fallback;
}

export function requireEnum<T extends string>(
  value: unknown,
  allowed: readonly T[],
  field: string,
): T {
  if (typeof value === 'string' && (allowed as readonly string[]).includes(value)) return value as T;
  throw validationError(`«${field}» qiymati noto‘g‘ri`, field);
}

export function optionalId(value: unknown, field: string): string | null {
  if (value === undefined || value === null || value === '') return null;
  return requireString(value, field, { max: 64 });
}

export function requireId(value: unknown, field: string): string {
  return requireString(value, field, { min: 1, max: 64 });
}

export function requireInt(value: unknown, field: string, min: number, max: number): number {
  const num = typeof value === 'number' ? value : Number(value);
  if (!Number.isFinite(num) || !Number.isInteger(num)) {
    throw validationError(`«${field}» butun son bo‘lishi kerak`, field);
  }
  if (num < min || num > max) throw validationError(`«${field}» ${min}…${max} oralig‘ida bo‘lsin`, field);
  return num;
}

export function requirePositiveNumber(value: unknown, field: string): number {
  const num = typeof value === 'number' ? value : Number(String(value).replace(/\s/g, '').replace(',', '.'));
  if (!Number.isFinite(num) || num <= 0) throw validationError(`«${field}» musbat son bo‘lishi kerak`, field);
  return num;
}

export function requireCurrency(value: unknown, field = 'valyuta'): CurrencyCode {
  const text = requireString(value, field, { min: LIMITS.currencyLength, max: LIMITS.currencyLength });
  if (!/^[A-Za-z]{3}$/.test(text)) throw validationError('Valyuta kodi 3 ta harfdan iborat bo‘lsin', field);
  return text.toUpperCase();
}

export function requireRole(value: unknown): Role {
  return requireEnum<Role>(value, ROLES, 'rol');
}

export function requireUserStatus(value: unknown): UserStatus {
  return requireEnum<UserStatus>(value, USER_STATUSES, 'holat');
}

export function requireTransactionKind(value: unknown): TransactionKind {
  return requireEnum<TransactionKind>(value, TRANSACTION_KINDS, 'tur');
}

export function normalizeLogin(value: unknown): string {
  const login = requireString(value, 'login', { min: LIMITS.loginMin, max: LIMITS.loginMax });
  if (!/^[a-z0-9._-]+$/i.test(login)) {
    throw validationError('Login faqat harf, raqam va . _ - belgilaridan iborat bo‘lsin', 'login');
  }
  // Login registrga sezgir emas — v1 da kiritilgan qiymat lowercase qilinardi,
  // jadvaldagi qiymat esa yo'q, natijada tizimga kirib bo'lmasdi.
  return login.toLowerCase();
}

export interface PasswordCheck {
  ok: boolean;
  message?: string;
}

export function checkPasswordStrength(password: string): PasswordCheck {
  if (password.length < LIMITS.passwordMin) {
    return { ok: false, message: `Parol kamida ${LIMITS.passwordMin} ta belgidan iborat bo‘lsin` };
  }
  if (password.length > LIMITS.passwordMax) {
    return { ok: false, message: 'Parol juda uzun' };
  }
  const hasLetter = /[a-z]/i.test(password);
  const hasDigit = /\d/.test(password);
  if (!hasLetter || !hasDigit) {
    return { ok: false, message: 'Parolda kamida bitta harf va bitta raqam bo‘lsin' };
  }
  if (/^(.)\1+$/.test(password)) {
    return { ok: false, message: 'Parol juda oddiy' };
  }
  return { ok: true };
}

export function requirePassword(value: unknown, field = 'parol'): string {
  if (typeof value !== 'string') throw validationError('Parol kiritilmagan', field);
  const check = checkPasswordStrength(value);
  if (!check.ok) throw validationError(check.message ?? 'Parol talabga javob bermaydi', field);
  return value;
}

/** Sahifalash parametrlarini xavfsiz chegaraga soladi. */
export function normalizePaging(offset: unknown, limit: unknown): { offset: number; limit: number } {
  const safeOffset = Number.isFinite(Number(offset)) ? Math.max(0, Math.trunc(Number(offset))) : 0;
  const rawLimit = Number.isFinite(Number(limit)) ? Math.trunc(Number(limit)) : 100;
  const safeLimit = Math.min(Math.max(rawLimit, 1), LIMITS.pageMax);
  return { offset: safeOffset, limit: safeLimit };
}
