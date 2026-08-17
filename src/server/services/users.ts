/**
 * Foydalanuvchilarni boshqarish.
 *
 * v1 dan farqlar:
 *  - yozuvlar qator raqami emas, UUID bo'yicha topiladi (o'chirishdan keyin
 *    "boshqa xodimga tegib ketish" muammosi yo'q);
 *  - parol tahrirlashda ixtiyoriy — bo'sh qoldirilsa eskisi saqlanadi;
 *  - login takrorlanishi taqiqlangan;
 *  - oxirgi administratorni o'chirish yoki bloklash mumkin emas (tizimdan
 *    butunlay chiqib qolishning oldi olinadi);
 *  - bloklangan yoki paroli o'zgargan foydalanuvchining barcha sessiyalari bekor qilinadi.
 */

import { AppError, validationError } from '../../shared/api.js';
import type { ChangePasswordInput, UserInput } from '../../shared/api.js';
import type { PublicUser } from '../../shared/types.js';
import { sanitizePermissions } from '../../shared/permissions.js';
import {
  LIMITS,
  normalizeLogin,
  requirePassword,
  requireRole,
  requireString,
  requireUserStatus,
} from '../../shared/validation.js';
import { usersTable } from '../db.js';
import type { StoredUser } from '../db.js';
import { newId, withLock } from '../sheets.js';
import {
  buildPasswordFields,
  destroyAllSessionsFor,
  toPublicUser,
  verifyPassword,
} from './auth.js';
import * as audit from './audit.js';

export function list(): PublicUser[] {
  return usersTable
    .all()
    .slice()
    .sort((a, b) => a.login.localeCompare(b.login))
    .map(toPublicUser);
}

function activeAdminCount(excludeId?: string): number {
  return usersTable
    .all()
    .filter((user) => user.role === 'admin' && user.status === 'active' && user.id !== excludeId).length;
}

function assertLoginAvailable(login: string, currentId: string | null): void {
  const clash = usersTable.find(
    (user) => user.login.trim().toLowerCase() === login && user.id !== currentId,
  );
  if (clash) throw validationError('Bunday login allaqachon band', 'login');
}

export function save(input: UserInput, actor: StoredUser): { id: string } {
  const login = normalizeLogin(input.login);
  const role = requireRole(input.role);
  const status = input.status === undefined ? 'active' : requireUserStatus(input.status);
  const displayName =
    input.displayName === undefined || input.displayName === null || input.displayName === ''
      ? login
      : requireString(input.displayName, 'ism', { min: 1, max: LIMITS.nameMax });
  const permissions = sanitizePermissions(input.permissions ?? {});

  return withLock(() => {
    const id = input.id ? String(input.id) : null;
    assertLoginAvailable(login, id);

    const now = new Date().toISOString();

    if (id) {
      const current = usersTable.findById(id);
      if (!current) throw new AppError('NOT_FOUND', 'Xodim topilmadi');

      const losingAdmin = current.role === 'admin' && (role !== 'admin' || status !== 'active');
      if (losingAdmin && activeAdminCount(current.id) === 0) {
        throw new AppError('CONFLICT', 'Tizimda kamida bitta faol administrator qolishi shart');
      }

      const patch: Partial<StoredUser> = {
        login,
        displayName,
        role,
        status,
        permissions,
        updatedAt: now,
      };

      // Parol faqat kiritilgan bo'lsa o'zgaradi.
      const wantsPasswordChange =
        typeof input.password === 'string' && input.password.trim() !== '';
      if (wantsPasswordChange) {
        Object.assign(patch, buildPasswordFields(requirePassword(input.password)));
      }

      usersTable.update(id, patch);
      if (wantsPasswordChange || status !== 'active' || current.login !== login) {
        destroyAllSessionsFor(id);
      }
      audit.record(actor, 'user.update', 'user', id, { login, role, status });
      return { id };
    }

    const password = requirePassword(input.password, 'parol');
    const user: StoredUser = {
      id: newId(),
      login,
      displayName,
      ...buildPasswordFields(password),
      role,
      status,
      permissions,
      createdAt: now,
      updatedAt: now,
      lastLoginAt: null,
    };
    usersTable.insert(user);
    audit.record(actor, 'user.create', 'user', user.id, { login, role });
    return { id: user.id };
  });
}

export function remove(id: string, actor: StoredUser): void {
  withLock(() => {
    const user = usersTable.findById(id);
    if (!user) throw new AppError('NOT_FOUND', 'Xodim topilmadi');
    if (user.id === actor.id) throw new AppError('CONFLICT', 'O‘zingizni o‘chira olmaysiz');
    if (user.role === 'admin' && activeAdminCount(user.id) === 0) {
      throw new AppError('CONFLICT', 'Tizimda kamida bitta faol administrator qolishi shart');
    }

    destroyAllSessionsFor(id);
    usersTable.deleteById(id);
    audit.record(actor, 'user.delete', 'user', id, { login: user.login });
  });
}

export function changeOwnPassword(input: ChangePasswordInput, actor: StoredUser): void {
  if (!verifyPassword(actor, String(input.currentPassword ?? ''))) {
    throw validationError('Joriy parol noto‘g‘ri', 'currentPassword');
  }
  const next = requirePassword(input.newPassword, 'newPassword');
  if (verifyPassword(actor, next)) {
    throw validationError('Yangi parol eskisidan farq qilsin', 'newPassword');
  }

  withLock(() => {
    usersTable.update(actor.id, {
      ...buildPasswordFields(next),
      updatedAt: new Date().toISOString(),
    });
  });
  destroyAllSessionsFor(actor.id);
  audit.record(actor, 'user.password', 'user', actor.id, null);
}

/** Birinchi ishga tushirish: faqat foydalanuvchilar umuman bo'lmaganda ishlaydi. */
export function createFirstAdmin(login: string, password: string): PublicUser {
  return withLock(() => {
    if (usersTable.count() > 0) {
      throw new AppError('CONFLICT', 'Tizim allaqachon sozlangan');
    }
    const normalized = normalizeLogin(login);
    const checked = requirePassword(password);
    const now = new Date().toISOString();

    const user: StoredUser = {
      id: newId(),
      login: normalized,
      displayName: normalized,
      ...buildPasswordFields(checked),
      role: 'admin',
      status: 'active',
      permissions: {},
      createdAt: now,
      updatedAt: now,
      lastLoginAt: null,
    };
    usersTable.insert(user);
    audit.record({ id: user.id, login: user.login }, 'user.bootstrap', 'user', user.id, null);
    return toPublicUser(user);
  });
}

export function touchLastLogin(userId: string): void {
  try {
    usersTable.update(userId, { lastLoginAt: new Date().toISOString() });
  } catch {
    // Login jarayonini bu xato to'xtatmasligi kerak.
  }
}
