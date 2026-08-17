import { describe, expect, it } from 'vitest';
import {
  can,
  canAny,
  effectivePermissions,
  mergePermissions,
  sanitizePermissions,
} from '../src/shared/permissions.js';

describe('can — fail-closed', () => {
  it('ruxsatlar yo‘q bo‘lsa hech narsaga ruxsat bermaydi', () => {
    // v1 da bu holat teskari edi: `!permissions` → to'liq admin huquqi.
    expect(can(null, 'income', 'view')).toBe(false);
    expect(can(undefined, 'users', 'delete')).toBe(false);
    expect(can({}, 'income', 'view')).toBe(false);
  });

  it('faqat aniq berilgan ruxsatni tan oladi', () => {
    const permissions = { income: { view: true } };
    expect(can(permissions, 'income', 'view')).toBe(true);
    expect(can(permissions, 'income', 'create')).toBe(false);
    expect(can(permissions, 'expense', 'view')).toBe(false);
  });

  it('`false` qiymatini ruxsat deb hisoblamaydi', () => {
    expect(can({ income: { view: false } }, 'income', 'view')).toBe(false);
  });
});

describe('effectivePermissions', () => {
  it('admin har doim to‘liq huquqqa ega', () => {
    const permissions = effectivePermissions('admin', {});
    expect(can(permissions, 'users', 'delete')).toBe(true);
    expect(can(permissions, 'config', 'edit')).toBe(true);
  });

  it('adminni shaxsiy sozlama bilan cheklab bo‘lmaydi', () => {
    // O'zini tizimdan qulflab qo'yishning oldi olinadi.
    const permissions = effectivePermissions('admin', { users: { delete: false } });
    expect(can(permissions, 'users', 'delete')).toBe(true);
  });

  it('operator standarti faqat ko‘rish va qo‘shishni beradi', () => {
    const permissions = effectivePermissions('operator', {});
    expect(can(permissions, 'income', 'create')).toBe(true);
    expect(can(permissions, 'income', 'delete')).toBe(false);
    expect(can(permissions, 'users', 'view')).toBe(false);
  });

  it('shaxsiy sozlamalar rol standartiga qo‘shiladi', () => {
    const permissions = effectivePermissions('operator', { income: { delete: true } });
    expect(can(permissions, 'income', 'delete')).toBe(true);
    expect(can(permissions, 'income', 'create')).toBe(true);
  });

  it('kuzatuvchi hech narsani o‘zgartira olmaydi', () => {
    const permissions = effectivePermissions('viewer', {});
    expect(can(permissions, 'income', 'view')).toBe(true);
    expect(can(permissions, 'income', 'create')).toBe(false);
    expect(can(permissions, 'expense', 'edit')).toBe(false);
  });
});

describe('mergePermissions', () => {
  it('ustki qiymat ustun turadi', () => {
    const merged = mergePermissions({ income: { view: true, edit: true } }, { income: { edit: false } });
    expect(can(merged, 'income', 'view')).toBe(true);
    expect(can(merged, 'income', 'edit')).toBe(false);
  });
});

describe('sanitizePermissions', () => {
  it('noma‘lum kalitlarni tashlab yuboradi', () => {
    const result = sanitizePermissions({
      income: { view: true, hack: true },
      __proto__: { view: true },
      nonsense: { view: true },
    });
    expect(result).toEqual({ income: { view: true } });
  });

  it('obyekt bo‘lmagan kiritmada bo‘sh natija', () => {
    expect(sanitizePermissions('admin')).toEqual({});
    expect(sanitizePermissions(null)).toEqual({});
  });

  it('faqat `true` qiymatini saqlaydi', () => {
    expect(sanitizePermissions({ income: { view: 'yes', create: true } })).toEqual({
      income: { create: true },
    });
  });
});

describe('canAny', () => {
  it('kamida bitta resursda ruxsat bo‘lsa true', () => {
    const permissions = { expense: { view: true } };
    expect(canAny(permissions, ['income', 'expense', 'transfer'], 'view')).toBe(true);
    expect(canAny(permissions, ['income', 'transfer'], 'view')).toBe(false);
  });
});
