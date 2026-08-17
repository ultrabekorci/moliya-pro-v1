/**
 * Ilova sozlamalari. Barcha "sozlanuvchi" xatti-harakat shu yerdan boshqariladi —
 * hech qanday biznes qoidasi kodga qotirib yozilmagan.
 */

import { configTable } from '../db.js';
import { withLock } from '../sheets.js';
import type { AppConfig, ExpenseAllocation } from '../../shared/types.js';
import { EXPENSE_ALLOCATIONS } from '../../shared/types.js';
import { requireCurrency, requireInt, requireString } from '../../shared/validation.js';

export const DEFAULT_CONFIG: AppConfig = {
  baseCurrency: 'UZS',
  organizationName: 'Moliya-Pro',
  sessionTtlMinutes: 60 * 8,
  maxLoginAttempts: 5,
  loginLockoutMinutes: 15,
  maxTransactionAmount: 1_000_000_000,
  maxNoteLength: 500,
  sharedExpenseAllocation: 'proRataIncome',
  fxProvider: 'cbu.uz',
};

let cached: AppConfig | null = null;

function parseValue(key: keyof AppConfig, raw: string): unknown {
  switch (key) {
    case 'sessionTtlMinutes':
    case 'maxLoginAttempts':
    case 'loginLockoutMinutes':
    case 'maxTransactionAmount':
    case 'maxNoteLength': {
      const num = Number(raw);
      return Number.isFinite(num) ? num : DEFAULT_CONFIG[key];
    }
    default:
      return raw;
  }
}

export function getConfig(): AppConfig {
  if (cached) return cached;

  const stored = configTable.all();
  const config: AppConfig = { ...DEFAULT_CONFIG };

  for (const row of stored) {
    const key = row.id as keyof AppConfig;
    if (!(key in DEFAULT_CONFIG)) continue;
    const value = parseValue(key, row.value);
    if (value === '' || value === undefined || value === null) continue;
    (config as unknown as Record<string, unknown>)[key] = value;
  }

  if (!(EXPENSE_ALLOCATIONS as readonly string[]).includes(config.sharedExpenseAllocation)) {
    config.sharedExpenseAllocation = DEFAULT_CONFIG.sharedExpenseAllocation;
  }
  if (config.fxProvider !== 'cbu.uz' && config.fxProvider !== 'none') {
    config.fxProvider = DEFAULT_CONFIG.fxProvider;
  }

  cached = config;
  return config;
}

export function invalidateConfigCache(): void {
  cached = null;
}

/** Faqat berilgan maydonlarni yangilaydi; noma'lum kalitlar e'tiborsiz qoldiriladi. */
export function saveConfig(patch: Partial<AppConfig>): AppConfig {
  const next: AppConfig = { ...getConfig() };

  if (patch.baseCurrency !== undefined) next.baseCurrency = requireCurrency(patch.baseCurrency, 'baseCurrency');
  if (patch.organizationName !== undefined) {
    next.organizationName = requireString(patch.organizationName, 'organizationName', { min: 1, max: 80 });
  }
  if (patch.sessionTtlMinutes !== undefined) {
    next.sessionTtlMinutes = requireInt(patch.sessionTtlMinutes, 'sessionTtlMinutes', 5, 60 * 24 * 7);
  }
  if (patch.maxLoginAttempts !== undefined) {
    next.maxLoginAttempts = requireInt(patch.maxLoginAttempts, 'maxLoginAttempts', 3, 50);
  }
  if (patch.loginLockoutMinutes !== undefined) {
    next.loginLockoutMinutes = requireInt(patch.loginLockoutMinutes, 'loginLockoutMinutes', 1, 24 * 60);
  }
  if (patch.maxTransactionAmount !== undefined) {
    next.maxTransactionAmount = requireInt(patch.maxTransactionAmount, 'maxTransactionAmount', 1, 1e15);
  }
  if (patch.maxNoteLength !== undefined) {
    next.maxNoteLength = requireInt(patch.maxNoteLength, 'maxNoteLength', 0, 2000);
  }
  if (patch.sharedExpenseAllocation !== undefined) {
    const value = patch.sharedExpenseAllocation;
    next.sharedExpenseAllocation = (EXPENSE_ALLOCATIONS as readonly string[]).includes(value)
      ? (value as ExpenseAllocation)
      : DEFAULT_CONFIG.sharedExpenseAllocation;
  }
  if (patch.fxProvider !== undefined) {
    next.fxProvider = patch.fxProvider === 'none' ? 'none' : 'cbu.uz';
  }

  withLock(() => {
    for (const [key, value] of Object.entries(next)) {
      const existing = configTable.findById(key);
      if (existing) configTable.update(key, { value: String(value) });
      else configTable.insert({ id: key, value: String(value) });
    }
  });

  invalidateConfigCache();
  return next;
}
