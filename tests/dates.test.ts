import { describe, expect, it } from 'vitest';
import {
  addDays,
  addMonths,
  daysBetween,
  endOfMonth,
  formatDateInZone,
  isIsoDate,
  isWithin,
  previousPeriod,
  resolvePeriod,
  shouldGroupMonthly,
  startOfMonth,
  toIsoDate,
} from '../src/shared/dates.js';

describe('isIsoDate', () => {
  it('haqiqiy sanalarni qabul qiladi', () => {
    expect(isIsoDate('2026-08-17')).toBe(true);
    expect(isIsoDate('2024-02-29')).toBe(true);
  });

  it('mavjud bo‘lmagan sanani rad etadi', () => {
    expect(isIsoDate('2025-02-30')).toBe(false);
    expect(isIsoDate('2026-13-01')).toBe(false);
    expect(isIsoDate('17.08.2026')).toBe(false);
    expect(isIsoDate('')).toBe(false);
  });
});

describe('toIsoDate', () => {
  it('kun.oy.yil formatini o‘giradi', () => {
    expect(toIsoDate('17.08.2026')).toBe('2026-08-17');
  });

  it('Date obyektini vaqt mintaqasi bilan o‘giradi', () => {
    // UTC bo'yicha 2026-08-16T21:00Z — Toshkentda allaqachon 17-avgust.
    const date = new Date('2026-08-16T21:00:00.000Z');
    expect(formatDateInZone(date, 'Asia/Tashkent')).toBe('2026-08-17');
    expect(formatDateInZone(date, 'UTC')).toBe('2026-08-16');
  });

  it('tushunarsiz qiymatda null qaytaradi', () => {
    expect(toIsoDate('salom')).toBeNull();
    expect(toIsoDate(null)).toBeNull();
  });
});

describe('sana arifmetikasi', () => {
  it('kun qo‘shadi va oy chegarasidan o‘tadi', () => {
    expect(addDays('2026-01-31', 1)).toBe('2026-02-01');
    expect(addDays('2026-03-01', -1)).toBe('2026-02-28');
  });

  it('oy qo‘shganda mavjud bo‘lmagan kunni kesadi', () => {
    expect(addMonths('2026-01-31', 1)).toBe('2026-02-28');
    expect(addMonths('2026-12-15', 1)).toBe('2027-01-15');
  });

  it('oy chegaralarini topadi', () => {
    expect(startOfMonth('2026-08-17')).toBe('2026-08-01');
    expect(endOfMonth('2026-02-10')).toBe('2026-02-28');
    expect(endOfMonth('2024-02-10')).toBe('2024-02-29');
  });

  it('kunlar sonini ikkala chekka bilan sanaydi', () => {
    expect(daysBetween('2026-08-01', '2026-08-01')).toBe(1);
    expect(daysBetween('2026-08-01', '2026-08-31')).toBe(31);
  });
});

describe('resolvePeriod', () => {
  const today = '2026-08-17';

  it('joriy oyni beradi', () => {
    expect(resolvePeriod('this_month', today)).toEqual({ start: '2026-08-01', end: '2026-08-31' });
  });

  it('o‘tgan oyni beradi', () => {
    expect(resolvePeriod('last_month', today)).toEqual({ start: '2026-07-01', end: '2026-07-31' });
  });

  it('chorakni to‘g‘ri hisoblaydi', () => {
    expect(resolvePeriod('this_quarter', today)).toEqual({ start: '2026-07-01', end: '2026-09-30' });
  });

  it('teskari kiritilgan maxsus davrni tartiblaydi', () => {
    const range = resolvePeriod('custom', today, { start: '2026-09-01', end: '2026-08-01' });
    expect(range).toEqual({ start: '2026-08-01', end: '2026-09-01' });
  });
});

describe('previousPeriod', () => {
  it('joriy oy uchun o‘tgan oyni SHU KUNGACHA kesadi', () => {
    const range = resolvePeriod('this_month', '2026-08-17');
    const previous = previousPeriod('this_month', range, '2026-08-17');
    expect(previous).toEqual({ start: '2026-07-01', end: '2026-07-17' });
  });

  it('to‘liq o‘tgan oy uchun to‘liq oldingi oyni beradi', () => {
    const range = resolvePeriod('last_month', '2026-08-17');
    const previous = previousPeriod('last_month', range, '2026-08-17');
    expect(previous).toEqual({ start: '2026-06-01', end: '2026-06-30' });
  });

  it('maxsus davr uchun shuncha uzunlikdagi oldingi oraliqni beradi', () => {
    const range = { start: '2026-08-10', end: '2026-08-19' };
    expect(previousPeriod('custom', range, '2026-08-20')).toEqual({
      start: '2026-07-31',
      end: '2026-08-09',
    });
  });
});

describe('isWithin', () => {
  it('chegara kunlarini ham qamrab oladi', () => {
    const range = { start: '2026-08-01', end: '2026-08-31' };
    expect(isWithin('2026-08-01', range)).toBe(true);
    expect(isWithin('2026-08-31', range)).toBe(true);
    expect(isWithin('2026-09-01', range)).toBe(false);
  });
});

describe('shouldGroupMonthly', () => {
  it('uzun davr uchun oylik guruhlashni tanlaydi', () => {
    expect(shouldGroupMonthly({ start: '2026-01-01', end: '2026-12-31' })).toBe(true);
    expect(shouldGroupMonthly({ start: '2026-08-01', end: '2026-08-31' })).toBe(false);
  });
});
