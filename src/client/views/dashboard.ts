/**
 * Dashboard: KPI plitalari, dinamika grafigi va kesimlar.
 * Firma tanlovi katalogdan keladi — hech qanday firma nomi kodda yozilmagan.
 */

import { call, errorMessage } from '../api.js';
import { el, fillSelect } from '../dom.js';
import { renderBarList, renderTrendChart, statTile } from '../charts.js';
import { emptyState, section, spinner, toast } from '../ui.js';
import { activeFirms, baseCurrency, state } from '../store.js';
import { PERIOD_PRESETS, previousPeriod, resolvePeriod } from '../../shared/dates.js';
import type { PeriodPreset } from '../../shared/dates.js';
import { growthPercent } from '../../shared/stats.js';
import { formatMinor } from '../../shared/money.js';
import type { DashboardStats } from '../../shared/types.js';

const PERIOD_LABELS: Record<PeriodPreset, string> = {
  today: 'Bugun',
  this_month: 'Shu oy',
  last_month: "O'tgan oy",
  this_quarter: 'Shu chorak',
  this_year: 'Shu yil',
  last_year: "O'tgan yil",
  all: 'Butun davr',
  custom: 'Boshqa davr',
};

interface DashboardControls {
  preset: PeriodPreset;
  firmId: string;
  customStart: string;
  customEnd: string;
}

const controls: DashboardControls = {
  preset: 'this_month',
  firmId: '',
  customStart: '',
  customEnd: '',
};

export function renderDashboard(container: HTMLElement): void {
  container.replaceChildren();

  const periodSelect = el('select', { class: 'input input--compact', name: 'period' });
  fillSelect(
    periodSelect,
    PERIOD_PRESETS.map((preset) => ({ value: preset, label: PERIOD_LABELS[preset] })),
    controls.preset,
  );

  const firmSelect = el('select', { class: 'input input--compact', name: 'firm' });
  fillSelect(
    firmSelect,
    activeFirms().map((firm) => ({ value: firm.id, label: firm.name })),
    controls.firmId,
    'Barcha firmalar',
  );

  const startInput = el('input', {
    class: 'input input--compact',
    type: 'date',
    value: controls.customStart || state.today,
  });
  const endInput = el('input', {
    class: 'input input--compact',
    type: 'date',
    value: controls.customEnd || state.today,
  });
  const customBox = el('div', { class: 'filters__custom' }, [startInput, el('span', { text: '—' }), endInput]);
  customBox.hidden = controls.preset !== 'custom';

  const content = el('div', { class: 'dashboard__content' }, [spinner()]);

  const reload = (): void => {
    controls.preset = periodSelect.value as PeriodPreset;
    controls.firmId = firmSelect.value;
    controls.customStart = startInput.value;
    controls.customEnd = endInput.value;
    customBox.hidden = controls.preset !== 'custom';
    void load(content);
  };

  periodSelect.addEventListener('change', reload);
  firmSelect.addEventListener('change', reload);
  startInput.addEventListener('change', reload);
  endInput.addEventListener('change', reload);

  container.append(
    el('div', { class: 'filters' }, [periodSelect, firmSelect, customBox]),
    content,
  );

  void load(content);
}

async function load(content: HTMLElement): Promise<void> {
  content.replaceChildren(spinner('Hisoblanmoqda…'));

  const range = resolvePeriod(controls.preset, state.today, {
    start: controls.customStart,
    end: controls.customEnd,
  });
  const previousRange = previousPeriod(controls.preset, range, state.today);

  try {
    const stats = await call('tx.dashboard', {
      range,
      previousRange,
      firmId: controls.firmId || null,
    });
    renderStats(content, stats);
  } catch (error) {
    content.replaceChildren(emptyState('Ma’lumotni yuklab bo‘lmadi', errorMessage(error)));
    toast(errorMessage(error), 'error');
  }
}

function renderStats(content: HTMLElement, stats: DashboardStats): void {
  const currency = stats.baseCurrency || baseCurrency();

  const tiles = el('div', { class: 'stat-grid' }, [
    statTile({
      label: 'Kirim',
      valueMinor: stats.current.incomeMinor,
      currency,
      deltaPercent: growthPercent(stats.current.incomeMinor, stats.previous.incomeMinor),
    }),
    statTile({
      label: 'Chiqim',
      valueMinor: stats.current.expenseMinor,
      currency,
      deltaPercent: growthPercent(stats.current.expenseMinor, stats.previous.expenseMinor),
      inverse: true,
    }),
    statTile({
      label: 'Foyda',
      valueMinor: stats.current.profitMinor,
      currency,
      deltaPercent: growthPercent(stats.current.profitMinor, stats.previous.profitMinor),
    }),
    statTile({
      label: 'Yozuvlar soni',
      text: String(stats.transactionCount),
      hint: 'tanlangan davrda',
    }),
  ]);

  const trendHost = el('div', { class: 'viz-host' });
  const pointsHost = el('div');
  const categoriesHost = el('div');
  const firmsHost = el('div');
  const accountsHost = el('div');

  content.replaceChildren(
    tiles,
    section('Kirim va chiqim dinamikasi', [trendHost]),
    el('div', { class: 'grid grid--2' }, [
      section('Savdo nuqtalari bo‘yicha kirim', [pointsHost]),
      section('Kategoriyalar bo‘yicha chiqim', [categoriesHost]),
    ]),
    el('div', { class: 'grid grid--2' }, [
      section('Firmalar kesimi', [firmsHost]),
      section('Hisoblar bo‘yicha kirim', [accountsHost]),
    ]),
  );

  renderTrendChart(trendHost, stats.trend, { currency });
  renderBarList(pointsHost, stats.byPoint, { currency });
  renderBarList(categoriesHost, stats.byCategory, { currency });
  renderBarList(firmsHost, stats.byFirm, { currency, emptyText: 'Firmalar bo‘yicha ma’lumot yo‘q' });
  renderBarList(accountsHost, stats.byAccount, { currency });

  // Rang yagona ajratuvchi belgi bo'lib qolmasligi uchun: raqamli xulosa matn ko'rinishida ham beriladi.
  content.appendChild(
    el('p', { class: 'dashboard__summary' }, [
      el('span', { text: 'Jami kirim: ' }),
      el('strong', { text: formatMinor(stats.current.incomeMinor, currency, { maxDecimals: 0 }) }),
      el('span', { text: ' · Jami chiqim: ' }),
      el('strong', { text: formatMinor(stats.current.expenseMinor, currency, { maxDecimals: 0 }) }),
      el('span', { text: ` · Valyuta: ${currency}` }),
    ]),
  );
}
