/**
 * Autentifikatsiya va sessiyalar — BUTUNLAY server tomonda.
 *
 * v1 dagi holat: barcha loginlar va OCHIQ MATNDAGI parollar HTML sahifaga
 * joylashtirilardi, tekshiruv esa brauzerda `find(u => u.p === p)` bilan
 * bajarilardi. Endi:
 *   - parol faqat xesh + tuz (salt) ko'rinishida saqlanadi (iterativ HMAC-SHA256);
 *   - klient hech qachon boshqa foydalanuvchining ma'lumotini ko'rmaydi;
 *   - har bir so'rov sessiya tokeni bilan tekshiriladi;
 *   - ketma-ket noto'g'ri urinishlar vaqtincha bloklanadi.
 */

import { AppError } from '../../shared/api.js';
import type { PublicUser } from '../../shared/types.js';
import { effectivePermissions } from '../../shared/permissions.js';
import { usersTable } from '../db.js';
import type { StoredUser } from '../db.js';
import { getConfig } from './config.js';

const DEFAULT_ITERATIONS = 12_000;
const SESSION_PREFIX = 'sess:';
const LOGIN_FAIL_PREFIX = 'lf:';
const MAX_CACHE_SECONDS = 21_600; // CacheService chegarasi: 6 soat.

function bytes(text: string): number[] {
  return Utilities.newBlob(text).getBytes();
}

function base64(input: number[]): string {
  return Utilities.base64Encode(input);
}

export function randomSalt(): string {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

/**
 * Iterativ HMAC-SHA256 (PBKDF2 ruhida). Apps Script'da tayyor PBKDF2 yo'q,
 * shuning uchun qo'lda takrorlanadi — bu shunchaki bitta SHA'dan ancha qimmatroq
 * va lug'at hujumini sezilarli sekinlashtiradi.
 */
export function hashPassword(password: string, salt: string, iterations: number): string {
  const keyBytes = bytes(salt);
  let digest = Utilities.computeHmacSha256Signature(bytes(`${salt}:${password}`), keyBytes);
  for (let i = 1; i < iterations; i += 1) {
    digest = Utilities.computeHmacSha256Signature(digest, keyBytes);
  }
  return base64(digest);
}

/** Vaqt bo'yicha doimiy solishtirish — javob vaqtidan parolni "tuslash" mumkin emas. */
export function constantTimeEquals(a: string, b: string): boolean {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i += 1) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

export function toPublicUser(user: StoredUser): PublicUser {
  return {
    id: user.id,
    login: user.login,
    displayName: user.displayName || user.login,
    role: user.role,
    status: user.status,
    permissions: effectivePermissions(user.role, user.permissions),
    createdAt: user.createdAt,
    lastLoginAt: user.lastLoginAt,
  };
}

// ---------------------------------------------------------------------------
// Sessiyalar
// ---------------------------------------------------------------------------

interface SessionRecord {
  userId: string;
  expiresAt: number;
}

function tokenKey(token: string): string {
  // Tokenning o'zi emas, xeshi saqlanadi: Properties tarkibi sizib chiqsa ham
  // tayyor token qo'lga tushmaydi.
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token);
  return SESSION_PREFIX + base64(digest);
}

function properties(): GoogleAppsScript.Properties.Properties {
  return PropertiesService.getScriptProperties();
}

function cache(): GoogleAppsScript.Cache.Cache {
  return CacheService.getScriptCache();
}

export function createSession(userId: string): { token: string; expiresAt: Date } {
  const config = getConfig();
  const token = `${Utilities.getUuid()}${Utilities.getUuid()}`.replace(/-/g, '');
  const expiresAt = new Date(Date.now() + config.sessionTtlMinutes * 60_000);
  const record: SessionRecord = { userId, expiresAt: expiresAt.getTime() };
  const key = tokenKey(token);
  const payload = JSON.stringify(record);

  properties().setProperty(key, payload);
  cache().put(key, payload, Math.min(config.sessionTtlMinutes * 60, MAX_CACHE_SECONDS));
  return { token, expiresAt };
}

export function resolveSession(token: string | null | undefined): StoredUser {
  if (!token || typeof token !== 'string' || token.length < 32) {
    throw new AppError('UNAUTHENTICATED', 'Tizimga kiring');
  }

  const key = tokenKey(token);
  let payload = cache().get(key);
  if (!payload) {
    payload = properties().getProperty(key);
    if (payload) cache().put(key, payload, 600);
  }
  if (!payload) throw new AppError('UNAUTHENTICATED', 'Sessiya topilmadi, qaytadan kiring');

  let record: SessionRecord;
  try {
    record = JSON.parse(payload) as SessionRecord;
  } catch {
    destroySession(token);
    throw new AppError('UNAUTHENTICATED', 'Sessiya buzilgan, qaytadan kiring');
  }

  if (!record.expiresAt || record.expiresAt < Date.now()) {
    destroySession(token);
    throw new AppError('UNAUTHENTICATED', 'Sessiya muddati tugadi, qaytadan kiring');
  }

  const user = usersTable.findById(record.userId);
  if (!user) {
    destroySession(token);
    throw new AppError('UNAUTHENTICATED', 'Foydalanuvchi topilmadi');
  }
  if (user.status !== 'active') {
    destroySession(token);
    throw new AppError('FORBIDDEN', 'Profilingiz bloklangan. Administrator bilan bog‘laning.');
  }
  return user;
}

export function destroySession(token: string | null | undefined): void {
  if (!token) return;
  const key = tokenKey(token);
  properties().deleteProperty(key);
  cache().remove(key);
}

/** Foydalanuvchining barcha sessiyalarini bekor qiladi (parol o'zgarganda, bloklanganda). */
export function destroyAllSessionsFor(userId: string): void {
  const store = properties();
  const all = store.getProperties();
  for (const [key, value] of Object.entries(all)) {
    if (!key.startsWith(SESSION_PREFIX)) continue;
    try {
      const record = JSON.parse(value) as SessionRecord;
      if (record.userId === userId) {
        store.deleteProperty(key);
        cache().remove(key);
      }
    } catch {
      store.deleteProperty(key);
    }
  }
}

/** Muddati o'tgan sessiyalarni tozalaydi (vaqtli trigger uchun). */
export function purgeExpiredSessions(): number {
  const store = properties();
  const all = store.getProperties();
  const now = Date.now();
  let removed = 0;

  for (const [key, value] of Object.entries(all)) {
    if (!key.startsWith(SESSION_PREFIX)) continue;
    let expired = true;
    try {
      const record = JSON.parse(value) as SessionRecord;
      expired = !record.expiresAt || record.expiresAt < now;
    } catch {
      expired = true;
    }
    if (expired) {
      store.deleteProperty(key);
      cache().remove(key);
      removed += 1;
    }
  }
  return removed;
}

// ---------------------------------------------------------------------------
// Login urinishlarini cheklash
// ---------------------------------------------------------------------------

function failKey(login: string): string {
  return LOGIN_FAIL_PREFIX + login;
}

export function assertNotLockedOut(login: string): void {
  const config = getConfig();
  const raw = cache().get(failKey(login));
  const attempts = raw ? Number(raw) : 0;
  if (attempts >= config.maxLoginAttempts) {
    throw new AppError(
      'RATE_LIMITED',
      `Juda ko‘p noto‘g‘ri urinish. ${config.loginLockoutMinutes} daqiqadan so‘ng qayta urining.`,
    );
  }
}

export function registerFailedLogin(login: string): void {
  const config = getConfig();
  const key = failKey(login);
  const raw = cache().get(key);
  const attempts = (raw ? Number(raw) : 0) + 1;
  cache().put(key, String(attempts), Math.min(config.loginLockoutMinutes * 60, MAX_CACHE_SECONDS));
}

export function clearFailedLogins(login: string): void {
  cache().remove(failKey(login));
}

// ---------------------------------------------------------------------------
// Login / parol
// ---------------------------------------------------------------------------

export interface NewPasswordFields {
  passwordHash: string;
  passwordSalt: string;
  passwordIterations: number;
}

export function buildPasswordFields(password: string): NewPasswordFields {
  const salt = randomSalt();
  return {
    passwordHash: hashPassword(password, salt, DEFAULT_ITERATIONS),
    passwordSalt: salt,
    passwordIterations: DEFAULT_ITERATIONS,
  };
}

export function verifyPassword(user: StoredUser, password: string): boolean {
  if (!user.passwordHash || !user.passwordSalt) return false;
  const iterations = user.passwordIterations > 0 ? user.passwordIterations : DEFAULT_ITERATIONS;
  const candidate = hashPassword(password, user.passwordSalt, iterations);
  return constantTimeEquals(candidate, user.passwordHash);
}

export function findUserByLogin(login: string): StoredUser | null {
  const normalized = login.trim().toLowerCase();
  return usersTable.find((user) => user.login.trim().toLowerCase() === normalized);
}

export function hasAnyUser(): boolean {
  return usersTable.count() > 0;
}
