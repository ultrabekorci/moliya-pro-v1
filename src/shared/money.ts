/**
 * Pul bilan ishlash. Barcha ichki hisob-kitoblar BUTUN SON (minor birlik) ustida
 * bajariladi — `0.1 + 0.2 !== 0.3` muammosi umuman yuzaga kelmaydi.
 */

import type { CurrencyCode, CurrencyMeta, MinorUnits } from './types.js';

const DEFAULT_DECIMALS = 2;

/**
 * Ma'lum valyutalar uchun kasr xonalari. Ro'yxatda yo'q valyuta uchun 2 olinadi,
 * shuning uchun yangi valyuta qo'shish uchun kodni o'zgartirish shart emas.
 */
const CURRENCY_META: Readonly<Record<string, CurrencyMeta>> = {
  UZS: { code: 'UZS', decimals: 2, symbol: "so'm" },
  USD: { code: 'USD', decimals: 2, symbol: '$' },
  EUR: { code: 'EUR', decimals: 2, symbol: '€' },
  RUB: { code: 'RUB', decimals: 2, symbol: '₽' },
  KZT: { code: 'KZT', decimals: 2, symbol: '₸' },
  JPY: { code: 'JPY', decimals: 0, symbol: '¥' },
};

export function currencyMeta(currency: CurrencyCode): CurrencyMeta {
  const found = CURRENCY_META[currency.toUpperCase()];
  if (found) return found;
  return { code: currency.toUpperCase(), decimals: DEFAULT_DECIMALS, symbol: currency.toUpperCase() };
}

export function currencyDecimals(currency: CurrencyCode): number {
  return currencyMeta(currency).decimals;
}

function pow10(n: number): number {
  let result = 1;
  for (let i = 0; i < n; i += 1) result *= 10;
  return result;
}

/**
 * Foydalanuvchi kiritgan matnni minor birlikka aylantiradi.
 * Bo'sh joy, `,` guruh ajratgichi va vergul-kasr (`12,5`) qo'llab-quvvatlanadi.
 *
 * @throws {RangeError} qiymat son bo'lmasa yoki manfiy/cheksiz bo'lsa.
 */
export function parseAmountToMinor(input: string | number, currency: CurrencyCode): MinorUnits {
  const decimals = currencyDecimals(currency);

  let normalized: string;
  if (typeof input === 'number') {
    if (!Number.isFinite(input)) throw new RangeError('Summa noto‘g‘ri');
    normalized = input.toFixed(decimals + 3);
  } else {
    normalized = input.replace(/[\s\u00a0'\u2019]/g, '');
    // "1,234.56" → "1234.56";  "12,5" → "12.5"
    if (normalized.includes(',') && normalized.includes('.')) {
      normalized = normalized.replace(/,/g, '');
    } else {
      normalized = normalized.replace(/,/g, '.');
    }
  }

  if (normalized === '' || !/^-?\d*(\.\d*)?$/.test(normalized)) {
    throw new RangeError('Summa faqat sonlardan iborat bo‘lishi kerak');
  }

  const negative = normalized.startsWith('-');
  const unsigned = negative ? normalized.slice(1) : normalized;
  const dot = unsigned.indexOf('.');
  const wholePart = dot === -1 ? unsigned : unsigned.slice(0, dot);
  const fracPart = dot === -1 ? '' : unsigned.slice(dot + 1);

  const whole = wholePart === '' ? 0 : Number(wholePart);
  if (!Number.isSafeInteger(whole)) throw new RangeError('Summa juda katta');

  // Yaxlitlash: kerakli xonadan keyingisiga qarab (banker emas, oddiy half-up).
  const keep = fracPart.slice(0, decimals).padEnd(decimals, '0');
  const nextDigit = fracPart.charAt(decimals);
  let minor = whole * pow10(decimals) + (keep === '' ? 0 : Number(keep));
  if (nextDigit !== '' && Number(nextDigit) >= 5) minor += 1;

  if (!Number.isSafeInteger(minor)) throw new RangeError('Summa juda katta');
  return negative ? -minor : minor;
}

/** Minor birlikni major songa aylantiradi (faqat ko'rsatish/eksport uchun). */
export function minorToMajor(minor: MinorUnits, currency: CurrencyCode): number {
  return minor / pow10(currencyDecimals(currency));
}

/** Guruh ajratgichli matn: `1 234 567,89`. */
export function formatMinor(
  minor: MinorUnits,
  currency: CurrencyCode,
  options: { withSymbol?: boolean; maxDecimals?: number } = {},
): string {
  const decimals = currencyDecimals(currency);
  const maxDecimals = options.maxDecimals ?? decimals;
  const negative = minor < 0;
  const abs = Math.abs(minor);
  const factor = pow10(decimals);
  const whole = Math.floor(abs / factor);
  const frac = abs % factor;

  const wholeText = String(whole).replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
  let text = wholeText;
  if (maxDecimals > 0 && frac !== 0) {
    const fracText = String(frac).padStart(decimals, '0').slice(0, maxDecimals).replace(/0+$/, '');
    if (fracText !== '') text += `,${fracText}`;
  }
  if (negative) text = `−${text}`;
  if (options.withSymbol) {
    const meta = currencyMeta(currency);
    text = meta.symbol === '$' || meta.symbol === '€' ? `${meta.symbol}${text}` : `${text} ${meta.symbol}`;
  }
  return text;
}

/**
 * Valyutani kurs bo'yicha o'girish.
 *
 * @param rate `from` valyutasining bir birligi necha `to` birligiga teng.
 */
export function convertMinor(
  amountMinor: MinorUnits,
  from: CurrencyCode,
  to: CurrencyCode,
  rate: number,
): MinorUnits {
  if (from.toUpperCase() === to.toUpperCase()) return amountMinor;
  if (!Number.isFinite(rate) || rate <= 0) throw new RangeError('Valyuta kursi noto‘g‘ri');
  const fromFactor = pow10(currencyDecimals(from));
  const toFactor = pow10(currencyDecimals(to));
  const converted = (amountMinor / fromFactor) * rate * toFactor;
  return Math.round(converted);
}

/** Butun songa yaxlitlangan xavfsiz yig'indi. */
export function sumMinor(values: readonly MinorUnits[]): MinorUnits {
  let total = 0;
  for (const value of values) total += value;
  return total;
}

/**
 * Umumiy summani ulushlar bo'yicha taqsimlaydi. Yaxlitlash qoldig'i eng katta
 * ulushga qo'shiladi, shuning uchun bo'laklar yig'indisi HAR DOIM `totalMinor` ga teng.
 */
export function allocateMinor(totalMinor: MinorUnits, weights: readonly number[]): MinorUnits[] {
  const count = weights.length;
  if (count === 0) return [];

  const totalWeight = weights.reduce((acc, w) => acc + (w > 0 ? w : 0), 0);
  if (totalWeight <= 0) {
    // Ulushlar yo'q → teng bo'lamiz.
    const base = Math.trunc(totalMinor / count);
    const parts = new Array<number>(count).fill(base);
    let remainder = totalMinor - base * count;
    for (let i = 0; remainder !== 0 && i < count; i += 1) {
      const step = remainder > 0 ? 1 : -1;
      parts[i] = (parts[i] ?? 0) + step;
      remainder -= step;
    }
    return parts;
  }

  const parts: number[] = [];
  let allocated = 0;
  let largestIndex = 0;
  let largestWeight = -Infinity;

  for (let i = 0; i < count; i += 1) {
    const weight = weights[i] ?? 0;
    const share = weight > 0 ? Math.round((totalMinor * weight) / totalWeight) : 0;
    parts.push(share);
    allocated += share;
    if (weight > largestWeight) {
      largestWeight = weight;
      largestIndex = i;
    }
  }

  const diff = totalMinor - allocated;
  if (diff !== 0) parts[largestIndex] = (parts[largestIndex] ?? 0) + diff;
  return parts;
}
