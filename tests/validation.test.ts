import { describe, expect, it } from 'vitest';
import {
  checkPasswordStrength,
  normalizeLogin,
  normalizePaging,
  requireCurrency,
  requireIsoDate,
  requirePassword,
  requireString,
} from '../src/shared/validation.js';
import { AppError } from '../src/shared/api.js';

describe('normalizeLogin', () => {
  it('registrni tenglashtiradi', () => {
    // v1 da kiritilgan login lowercase qilinar, jadvaldagi qiymat esa yo'q edi —
    // "Admin" deb yozilgan foydalanuvchi tizimga umuman kira olmasdi.
    expect(normalizeLogin('Admin')).toBe('admin');
    expect(normalizeLogin('  DILSHOD  ')).toBe('dilshod');
  });

  it('ruxsat etilmagan belgilarni rad etadi', () => {
    expect(() => normalizeLogin('ali baba')).toThrow(AppError);
    expect(() => normalizeLogin('ali@example.com')).toThrow(AppError);
  });

  it('juda qisqa loginni rad etadi', () => {
    expect(() => normalizeLogin('ab')).toThrow(AppError);
  });
});

describe('checkPasswordStrength', () => {
  it('kuchli parolni qabul qiladi', () => {
    expect(checkPasswordStrength('moliya2026').ok).toBe(true);
  });

  it('qisqa parolni rad etadi', () => {
    expect(checkPasswordStrength('abc12').ok).toBe(false);
  });

  it('faqat harfdan iborat parolni rad etadi', () => {
    expect(checkPasswordStrength('parolparol').ok).toBe(false);
  });

  it('takrorlanuvchi belgilardan iborat parolni rad etadi', () => {
    expect(checkPasswordStrength('11111111').ok).toBe(false);
  });

  it('requirePassword xatoni AppError sifatida beradi', () => {
    expect(() => requirePassword('123')).toThrow(AppError);
    expect(requirePassword('moliya2026')).toBe('moliya2026');
  });
});

describe('requireString', () => {
  it('bo‘sh joyni kesadi', () => {
    expect(requireString('  salom  ', 'nom')).toBe('salom');
  });

  it('uzunlik chegarasini tekshiradi', () => {
    expect(() => requireString('a', 'nom', { min: 2 })).toThrow(AppError);
    expect(() => requireString('abcdef', 'nom', { max: 3 })).toThrow(AppError);
  });

  it('matn bo‘lmagan qiymatni rad etadi', () => {
    expect(() => requireString(42, 'nom')).toThrow(AppError);
  });
});

describe('requireIsoDate', () => {
  it('to‘g‘ri sanani qabul qiladi', () => {
    expect(requireIsoDate('2026-08-17', 'sana')).toBe('2026-08-17');
  });

  it('noto‘g‘ri sanani rad etadi', () => {
    expect(() => requireIsoDate('17.08.2026', 'sana')).toThrow(AppError);
    expect(() => requireIsoDate('2026-02-31', 'sana')).toThrow(AppError);
  });
});

describe('requireCurrency', () => {
  it('kodni katta harfga keltiradi', () => {
    expect(requireCurrency('usd')).toBe('USD');
  });

  it('noto‘g‘ri uzunlikni rad etadi', () => {
    expect(() => requireCurrency('DOLLAR')).toThrow(AppError);
    expect(() => requireCurrency('12')).toThrow(AppError);
  });
});

describe('normalizePaging', () => {
  it('chegaradan chiqqan qiymatlarni to‘g‘rilaydi', () => {
    expect(normalizePaging(-5, 10_000)).toEqual({ offset: 0, limit: 500 });
    expect(normalizePaging('abc', undefined)).toEqual({ offset: 0, limit: 100 });
    expect(normalizePaging(20, 25)).toEqual({ offset: 20, limit: 25 });
  });
});
