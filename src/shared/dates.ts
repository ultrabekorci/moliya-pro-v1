/**
 * Sana bilan ishlash.
 *
 * Butun tizim bo'ylab sana `YYYY-MM-DD` MATN sifatida yuriydi va faqat matn
 * sifatida solishtiriladi. `new Date("2026-08-17")` (UTC) va
 * `new Date(2026, 7, 17)` (lokal) aralashib ketishi mumkin emas — v1 dagi
 * vaqt mintaqasiga bog'liq chegara xatolari shu bilan yopiladi.
 */

import type { DateRange, IsoDate, IsoTimestamp } from './types.js';

const ISO_DATE_RE = /^\d{4}-\d{2}-\d{2}$/;

export function isIsoDate(value: unknown): value is IsoDate {
  if (typeof value !== 'string' || !ISO_DATE_RE.test(value)) return false;
  const year = Number(value.slice(0, 4));
  const month = Number(value.slice(5, 7));
  const day = Number(value.slice(8, 10));
  if (month < 1 || month > 12 || day < 1) return false;
  return day <= daysInMonth(year, month);
}

export function daysInMonth(year: number, month: number): number {
  return new Date(Date.UTC(year, month, 0)).getUTCDate();
}

/** `Date`, matn yoki Sheets qiymatini `YYYY-MM-DD` ga keltiradi. */
export function toIsoDate(value: unknown, timeZone = 'UTC'): IsoDate | null {
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (isIsoDate(trimmed)) return trimmed;
    const match = /^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/.exec(trimmed);
    if (match) {
      const [, d, m, y] = match;
      const candidate = `${y}-${(m ?? '').padStart(2, '0')}-${(d ?? '').padStart(2, '0')}`;
      return isIsoDate(candidate) ? candidate : null;
    }
    return null;
  }
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return formatDateInZone(value, timeZone);
  }
  return null;
}

/**
 * `Date` ni berilgan vaqt mintaqasidagi kalendar sanaga aylantiradi.
 * `Intl` mavjud bo'lmagan muhitda UTC ga qaytadi.
 */
export function formatDateInZone(date: Date, timeZone: string): IsoDate {
  try {
    const formatter = new Intl.DateTimeFormat('en-CA', {
      timeZone,
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
    });
    const formatted = formatter.format(date);
    return isIsoDate(formatted) ? formatted : date.toISOString().slice(0, 10);
  } catch {
    return date.toISOString().slice(0, 10);
  }
}

export function todayIso(timeZone = 'UTC'): IsoDate {
  return formatDateInZone(new Date(), timeZone);
}

export function nowIso(): IsoTimestamp {
  return new Date().toISOString();
}

/** Sanalarni matn sifatida solishtiradi: <0, 0, >0. */
export function compareIsoDate(a: IsoDate, b: IsoDate): number {
  return a < b ? -1 : a > b ? 1 : 0;
}

export function isWithin(date: IsoDate, range: DateRange): boolean {
  return date >= range.start && date <= range.end;
}

function parts(date: IsoDate): { year: number; month: number; day: number } {
  return {
    year: Number(date.slice(0, 4)),
    month: Number(date.slice(5, 7)),
    day: Number(date.slice(8, 10)),
  };
}

function build(year: number, month: number, day: number): IsoDate {
  const clampedDay = Math.min(day, daysInMonth(year, month));
  return `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(clampedDay).padStart(2, '0')}`;
}

export function addDays(date: IsoDate, days: number): IsoDate {
  const { year, month, day } = parts(date);
  const shifted = new Date(Date.UTC(year, month - 1, day + days));
  return shifted.toISOString().slice(0, 10);
}

export function addMonths(date: IsoDate, months: number): IsoDate {
  const { year, month, day } = parts(date);
  const totalMonths = year * 12 + (month - 1) + months;
  const newYear = Math.floor(totalMonths / 12);
  const newMonth = (totalMonths % 12) + 1;
  return build(newYear, newMonth, day);
}

export function startOfMonth(date: IsoDate): IsoDate {
  return `${date.slice(0, 7)}-01`;
}

export function endOfMonth(date: IsoDate): IsoDate {
  const { year, month } = parts(date);
  return build(year, month, daysInMonth(year, month));
}

export function startOfYear(date: IsoDate): IsoDate {
  return `${date.slice(0, 4)}-01-01`;
}

export function endOfYear(date: IsoDate): IsoDate {
  return `${date.slice(0, 4)}-12-31`;
}

/** Ikki sana orasidagi kunlar soni (ikkala chekka ham hisobga olinadi). */
export function daysBetween(start: IsoDate, end: IsoDate): number {
  const a = Date.UTC(...(dateTuple(start) as [number, number, number]));
  const b = Date.UTC(...(dateTuple(end) as [number, number, number]));
  return Math.round((b - a) / 86_400_000) + 1;
}

function dateTuple(date: IsoDate): [number, number, number] {
  const { year, month, day } = parts(date);
  return [year, month - 1, day];
}

export const PERIOD_PRESETS = [
  'today',
  'this_month',
  'last_month',
  'this_quarter',
  'this_year',
  'last_year',
  'all',
  'custom',
] as const;

export type PeriodPreset = (typeof PERIOD_PRESETS)[number];

export const MIN_DATE: IsoDate = '1970-01-01';
export const MAX_DATE: IsoDate = '2999-12-31';

export function resolvePeriod(
  preset: PeriodPreset,
  today: IsoDate,
  custom?: Partial<DateRange>,
): DateRange {
  switch (preset) {
    case 'today':
      return { start: today, end: today };
    case 'this_month':
      return { start: startOfMonth(today), end: endOfMonth(today) };
    case 'last_month': {
      const prev = addMonths(startOfMonth(today), -1);
      return { start: startOfMonth(prev), end: endOfMonth(prev) };
    }
    case 'this_quarter': {
      const month = Number(today.slice(5, 7));
      const firstMonth = Math.floor((month - 1) / 3) * 3 + 1;
      const start = `${today.slice(0, 4)}-${String(firstMonth).padStart(2, '0')}-01`;
      return { start, end: endOfMonth(addMonths(start, 2)) };
    }
    case 'this_year':
      return { start: startOfYear(today), end: endOfYear(today) };
    case 'last_year': {
      const prevYear = String(Number(today.slice(0, 4)) - 1);
      return { start: `${prevYear}-01-01`, end: `${prevYear}-12-31` };
    }
    case 'all':
      return { start: MIN_DATE, end: MAX_DATE };
    case 'custom': {
      const start = custom?.start && isIsoDate(custom.start) ? custom.start : today;
      const end = custom?.end && isIsoDate(custom.end) ? custom.end : today;
      return start <= end ? { start, end } : { start: end, end: start };
    }
    default:
      return { start: today, end: today };
  }
}

/**
 * Oldingi taqqoslash davri.
 *
 * Joriy (tugallanmagan) oy/yil uchun oldingi davr ham SHU KUNGACHA kesiladi —
 * aks holda "to'liq o'tgan oy" ga nisbatan taqqoslash har doim salbiy ko'rinadi.
 */
export function previousPeriod(preset: PeriodPreset, range: DateRange, today: IsoDate): DateRange {
  switch (preset) {
    case 'today':
      return { start: addDays(range.start, -1), end: addDays(range.start, -1) };
    case 'this_month': {
      const prevStart = addMonths(range.start, -1);
      const dayOffset = daysBetween(range.start, minIso(today, range.end)) - 1;
      return { start: prevStart, end: minIso(addDays(prevStart, dayOffset), endOfMonth(prevStart)) };
    }
    case 'last_month': {
      const prevStart = addMonths(range.start, -1);
      return { start: prevStart, end: endOfMonth(prevStart) };
    }
    case 'this_quarter': {
      const prevStart = addMonths(range.start, -3);
      const dayOffset = daysBetween(range.start, minIso(today, range.end)) - 1;
      return { start: prevStart, end: addDays(prevStart, dayOffset) };
    }
    case 'this_year': {
      const prevStart = addMonths(range.start, -12);
      const dayOffset = daysBetween(range.start, minIso(today, range.end)) - 1;
      return { start: prevStart, end: addDays(prevStart, dayOffset) };
    }
    case 'last_year': {
      const prevYear = String(Number(range.start.slice(0, 4)) - 1);
      return { start: `${prevYear}-01-01`, end: `${prevYear}-12-31` };
    }
    case 'all':
      return { start: MIN_DATE, end: MIN_DATE };
    case 'custom':
    default: {
      const length = daysBetween(range.start, range.end);
      const end = addDays(range.start, -1);
      return { start: addDays(end, -(length - 1)), end };
    }
  }
}

export function minIso(a: IsoDate, b: IsoDate): IsoDate {
  return a <= b ? a : b;
}

export function maxIso(a: IsoDate, b: IsoDate): IsoDate {
  return a >= b ? a : b;
}

/** Trend grafigi uchun guruhlash kaliti: kunlik yoki oylik. */
export function trendKey(date: IsoDate, monthly: boolean): string {
  return monthly ? date.slice(0, 7) : date;
}

/** Davr uzunligiga qarab kunlik yoki oylik guruhlashni tanlaydi. */
export function shouldGroupMonthly(range: DateRange): boolean {
  return daysBetween(range.start, range.end) > 62;
}

const MONTHS_UZ = [
  'Yanvar',
  'Fevral',
  'Mart',
  'Aprel',
  'May',
  'Iyun',
  'Iyul',
  'Avgust',
  'Sentabr',
  'Oktabr',
  'Noyabr',
  'Dekabr',
];

/** Tur-qo'riqchisi emas — `formatHuman` kabi joylarda `never` ga torayib qolmaslik uchun. */
function looksLikeIsoDate(value: string): boolean {
  return isIsoDate(value as unknown);
}

/** `2026-08-17` → `17 Avgust 2026`;  `2026-08` → `Avgust 2026`. */
export function formatHuman(value: string): string {
  if (looksLikeIsoDate(value)) {
    const { year, month, day } = parts(value);
    return `${day} ${MONTHS_UZ[month - 1] ?? month} ${year}`;
  }
  if (/^\d{4}-\d{2}$/.test(value)) {
    const year = Number(value.slice(0, 4));
    const month = Number(value.slice(5, 7));
    return `${MONTHS_UZ[month - 1] ?? month} ${year}`;
  }
  return value;
}

/** Grafik o'qi uchun qisqa yozuv. */
export function formatShort(value: string): string {
  if (looksLikeIsoDate(value)) return `${value.slice(8, 10)}.${value.slice(5, 7)}`;
  if (/^\d{4}-\d{2}$/.test(value)) {
    const month = Number(value.slice(5, 7));
    return (MONTHS_UZ[month - 1] ?? value).slice(0, 3);
  }
  return value;
}
