import { describe, expect, it } from 'vitest';
import {
  allocateMinor,
  convertMinor,
  formatMinor,
  minorToMajor,
  parseAmountToMinor,
  sumMinor,
} from '../src/shared/money.js';

describe('parseAmountToMinor', () => {
  it('oddiy sonlarni minor birlikka aylantiradi', () => {
    expect(parseAmountToMinor('100', 'UZS')).toBe(10_000);
    expect(parseAmountToMinor('12.34', 'USD')).toBe(1234);
    expect(parseAmountToMinor(0.05, 'USD')).toBe(5);
  });

  it('bo‘sh joy va guruh ajratgichlarini tushunadi', () => {
    expect(parseAmountToMinor('1 234 567', 'UZS')).toBe(123_456_700);
    expect(parseAmountToMinor('1,234.56', 'USD')).toBe(123_456);
    expect(parseAmountToMinor('12,5', 'USD')).toBe(1250);
  });

  it('ortiqcha kasr xonalarini yaxlitlaydi', () => {
    expect(parseAmountToMinor('1.005', 'USD')).toBe(101);
    expect(parseAmountToMinor('1.004', 'USD')).toBe(100);
  });

  it('kasr xonasi yo‘q valyutani to‘g‘ri o‘qiydi', () => {
    expect(parseAmountToMinor('150', 'JPY')).toBe(150);
  });

  it('noto‘g‘ri qiymatda xato beradi', () => {
    expect(() => parseAmountToMinor('abc', 'UZS')).toThrow(RangeError);
    expect(() => parseAmountToMinor('', 'UZS')).toThrow(RangeError);
  });

  it('float yig‘indi xatosini keltirib chiqarmaydi', () => {
    // 0.1 + 0.2 !== 0.3 muammosi butun sonlarda umuman yuzaga kelmaydi.
    const total = sumMinor([parseAmountToMinor('0.1', 'USD'), parseAmountToMinor('0.2', 'USD')]);
    expect(total).toBe(parseAmountToMinor('0.3', 'USD'));
  });
});

describe('formatMinor', () => {
  // Guruh ajratgichi — uzilmas probel (U+00A0), summa qatorlarga bo‘linib ketmasligi uchun.
  const NBSP = '\u00a0';

  it('guruh ajratgichi bilan ko‘rsatadi', () => {
    expect(formatMinor(123_456_700, 'UZS', { maxDecimals: 0 })).toBe(`1${NBSP}234${NBSP}567`);
    expect(formatMinor(1234, 'USD')).toBe('12,34');
  });

  it('manfiy qiymatni belgilaydi', () => {
    expect(formatMinor(-10_000, 'UZS', { maxDecimals: 0 })).toBe('−100');
  });

  it('valyuta belgisini qo‘shadi', () => {
    expect(formatMinor(10_000, 'USD', { withSymbol: true })).toBe('$100');
    expect(formatMinor(10_000, 'UZS', { withSymbol: true })).toBe("100 so'm");
  });
});

describe('convertMinor', () => {
  it('valyutani kurs bo‘yicha o‘giradi', () => {
    // 100 USD × 12 500 = 1 250 000 UZS
    expect(convertMinor(10_000, 'USD', 'UZS', 12_500)).toBe(125_000_000);
  });

  it('bir xil valyutada qiymatni o‘zgartirmaydi', () => {
    expect(convertMinor(555, 'UZS', 'uzs', 999)).toBe(555);
  });

  it('noto‘g‘ri kursda xato beradi', () => {
    expect(() => convertMinor(100, 'USD', 'UZS', 0)).toThrow(RangeError);
  });
});

describe('allocateMinor', () => {
  it('yig‘indi har doim saqlanadi', () => {
    const parts = allocateMinor(100, [1, 1, 1]);
    expect(sumMinor(parts)).toBe(100);
  });

  it('ulushlarga proporsional taqsimlaydi', () => {
    expect(allocateMinor(1000, [3, 1])).toEqual([750, 250]);
  });

  it('ulushlar bo‘lmasa teng bo‘ladi', () => {
    const parts = allocateMinor(10, [0, 0, 0]);
    expect(sumMinor(parts)).toBe(10);
    expect(parts).toHaveLength(3);
  });

  it('bo‘sh ro‘yxatda bo‘sh natija', () => {
    expect(allocateMinor(100, [])).toEqual([]);
  });
});

describe('minorToMajor', () => {
  it('ko‘rsatish uchun teskari aylantiradi', () => {
    expect(minorToMajor(123_456, 'USD')).toBe(1234.56);
  });
});
