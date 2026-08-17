/**
 * Google Apps Script kirish nuqtalari.
 *
 * Diqqat: bu yerda sahifaga HECH QANDAY ma'lumot joylashtirilmaydi.
 * v1 da `doGet` barcha foydalanuvchilarni (ochiq matndagi parollari bilan)
 * HTML ichiga yozib yuborardi. Endi sahifa butunlay statik, hamma ma'lumot
 * autentifikatsiyadan o'tgan `rpc` chaqiruvlari orqali keladi.
 */

import { dispatch } from './router.js';
import { ensureSchema } from './db.js';
import { purgeExpiredSessions } from './services/auth.js';
import { seedDefaults } from './services/catalog.js';
import { getConfig } from './services/config.js';
export { migrateDryRun, migrateFromV1 } from './migrate.js';

export function doGet(): GoogleAppsScript.HTML.HtmlOutput {
  const config = getConfig();
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(config.organizationName || 'Moliya-Pro')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    // ALLOWALL emas: sahifani begona saytga joylash (clickjacking) taqiqlanadi.
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * Klientdan keladigan yagona chaqiruv. JSON matn qabul qiladi va JSON matn
 * qaytaradi — `google.script.run` orqali uzatishda tur ma'lumoti yo'qolmaydi.
 */
export function rpc(payload: string): string {
  let request: unknown;
  try {
    request = typeof payload === 'string' ? JSON.parse(payload) : payload;
  } catch {
    return JSON.stringify({
      ok: false,
      error: { code: 'VALIDATION', message: 'So‘rovni o‘qib bo‘lmadi' },
    });
  }
  return JSON.stringify(dispatch(request));
}

/**
 * Qo'lda ishga tushiriladi (Apps Script muharriridan): varaqlarni yaratadi va
 * boshlang'ich katalogni to'ldiradi. Administrator esa veb-sahifadagi
 * "birinchi ishga tushirish" oynasi orqali yaratiladi.
 */
export function setup(): string {
  ensureSchema();
  seedDefaults(getConfig().baseCurrency);
  return 'Tayyor. Endi veb-ilovani oching va birinchi administratorni yarating.';
}

/** Vaqtli trigger uchun: muddati o'tgan sessiyalarni tozalaydi. */
export function purgeSessions(): number {
  return purgeExpiredSessions();
}

/** UrlFetch ruxsatini bir marta so'rash uchun (kurs olish ishlashi uchun kerak). */
export function authorizeExternalRequests(): string {
  const response = UrlFetchApp.fetch('https://cbu.uz/uz/arkhiv-kursov-valyut/json/USD/', {
    muteHttpExceptions: true,
  });
  return `OK: ${response.getResponseCode()}`;
}
