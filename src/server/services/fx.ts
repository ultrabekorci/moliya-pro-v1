/**
 * Valyuta kurslari.
 *
 * v1 da kurs har bir chaqiruvda tashqi API'dan olinardi va agar u ishlamasa
 * dollar chiqimini umuman saqlab bo'lmasdi. Endi:
 *   - olingan kurs `FxRates` varag'iga yoziladi (tarix saqlanadi),
 *   - dam olish kunlari uchun eng yaqin oldingi kurs ishlatiladi,
 *   - internet ishlamasa kursni qo'lda kiritish mumkin,
 *   - takroriy chaqiruvlar keshdan javob oladi.
 */

import { AppError } from '../../shared/api.js';
import type { FxRateResult } from '../../shared/api.js';
import type { CurrencyCode, IsoDate } from '../../shared/types.js';
import { isIsoDate } from '../../shared/dates.js';
import { fxTable } from '../db.js';
import { newId, withLock } from '../sheets.js';
import { getConfig } from './config.js';

const CACHE_SECONDS = 6 * 60 * 60;

function cacheKey(date: IsoDate, currency: CurrencyCode): string {
  return `fx:${currency}:${date}`;
}

function storedRate(date: IsoDate, currency: CurrencyCode): FxRateResult | null {
  const exact = fxTable.find((row) => row.date === date && row.currency === currency && row.rate > 0);
  if (exact) {
    return { currency, rate: exact.rate, effectiveDate: exact.date, source: exact.source };
  }
  return null;
}

/** Berilgan sanadan oldingi (yoki shu kungi) eng yangi kurs. */
function nearestEarlierRate(date: IsoDate, currency: CurrencyCode): FxRateResult | null {
  const candidates = fxTable
    .filter((row) => row.currency === currency && row.rate > 0 && row.date <= date)
    .sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));
  const best = candidates[0];
  return best ? { currency, rate: best.rate, effectiveDate: best.date, source: best.source } : null;
}

function persist(result: FxRateResult): void {
  withLock(() => {
    const existing = fxTable.find(
      (row) => row.date === result.effectiveDate && row.currency === result.currency,
    );
    if (existing) {
      fxTable.update(existing.id, {
        rate: result.rate,
        source: result.source,
        fetchedAt: new Date().toISOString(),
      });
      return;
    }
    fxTable.insert({
      id: newId(),
      date: result.effectiveDate,
      currency: result.currency,
      rate: result.rate,
      source: result.source,
      fetchedAt: new Date().toISOString(),
    });
  });
}

interface CbuRow {
  Ccy?: string;
  Rate?: string;
  Date?: string;
}

/** `17.08.2026` → `2026-08-17`. */
function parseCbuDate(value: string | undefined, fallback: IsoDate): IsoDate {
  if (!value) return fallback;
  const match = /^(\d{2})\.(\d{2})\.(\d{4})$/.exec(value.trim());
  if (!match) return fallback;
  return `${match[3]}-${match[2]}-${match[1]}`;
}

function fetchFromCbu(date: IsoDate, currency: CurrencyCode): FxRateResult | null {
  const url = `https://cbu.uz/uz/arkhiv-kursov-valyut/json/${encodeURIComponent(currency)}/${date}/`;
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (response.getResponseCode() !== 200) return null;

    const parsed: unknown = JSON.parse(response.getContentText());
    if (!Array.isArray(parsed) || parsed.length === 0) return null;

    const row = parsed[0] as CbuRow;
    const rate = Number(row.Rate);
    if (!Number.isFinite(rate) || rate <= 0) return null;

    return {
      currency,
      rate,
      effectiveDate: parseCbuDate(row.Date, date),
      source: 'cbu.uz',
    };
  } catch (error) {
    console.warn(`FX fetch failed: ${String(error)}`);
    return null;
  }
}

/**
 * Kursni topadi. Tartib: kesh → jadval (aniq sana) → tashqi API → jadval (oldingi sana).
 * Hech qaysi manba javob bermasa xato qaytaradi — jimgina 0 qaytarilmaydi.
 */
export function getRate(date: IsoDate, currency: CurrencyCode): FxRateResult {
  const config = getConfig();
  const upper = currency.toUpperCase();
  if (upper === config.baseCurrency.toUpperCase()) {
    return { currency: upper, rate: 1, effectiveDate: date, source: 'base' };
  }
  if (!isIsoDate(date)) throw new AppError('VALIDATION', 'Sana noto‘g‘ri');

  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey(date, upper));
  if (cached) {
    try {
      return JSON.parse(cached) as FxRateResult;
    } catch {
      cache.remove(cacheKey(date, upper));
    }
  }

  const fromSheet = storedRate(date, upper);
  if (fromSheet) {
    cache.put(cacheKey(date, upper), JSON.stringify(fromSheet), CACHE_SECONDS);
    return fromSheet;
  }

  if (config.fxProvider === 'cbu.uz') {
    const fetched = fetchFromCbu(date, upper);
    if (fetched) {
      persist(fetched);
      cache.put(cacheKey(date, upper), JSON.stringify(fetched), CACHE_SECONDS);
      return fetched;
    }
  }

  const earlier = nearestEarlierRate(date, upper);
  if (earlier) {
    cache.put(cacheKey(date, upper), JSON.stringify(earlier), 60 * 30);
    return earlier;
  }

  throw new AppError(
    'NOT_FOUND',
    `${upper} uchun kurs topilmadi. Sozlamalardan kursni qo‘lda kiriting.`,
  );
}

export function setManualRate(date: IsoDate, currency: CurrencyCode, rate: number): FxRateResult {
  if (!isIsoDate(date)) throw new AppError('VALIDATION', 'Sana noto‘g‘ri');
  if (!Number.isFinite(rate) || rate <= 0) throw new AppError('VALIDATION', 'Kurs musbat son bo‘lishi kerak');

  const result: FxRateResult = {
    currency: currency.toUpperCase(),
    rate,
    effectiveDate: date,
    source: 'manual',
  };
  persist(result);
  CacheService.getScriptCache().put(
    cacheKey(date, result.currency),
    JSON.stringify(result),
    CACHE_SECONDS,
  );
  return result;
}

/**
 * Balansni hisobot valyutasiga o'girish uchun joriy kurslar.
 * Kursi topilmagan valyuta ro'yxatga kirmaydi (balans o'z valyutasida ko'rsatiladi).
 */
export function currentRates(currencies: readonly CurrencyCode[], today: IsoDate): Record<string, number> {
  const config = getConfig();
  const rates: Record<string, number> = { [config.baseCurrency.toUpperCase()]: 1 };

  for (const currency of currencies) {
    const upper = currency.toUpperCase();
    if (rates[upper] !== undefined) continue;
    try {
      rates[upper] = getRate(today, upper).rate;
    } catch {
      const fallback = nearestEarlierRate(today, upper);
      if (fallback) rates[upper] = fallback.rate;
    }
  }
  return rates;
}
