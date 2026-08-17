/**
 * Grafiklar — qo'lda yozilgan SVG. Chart.js yoki boshqa CDN kutubxonasi yo'q:
 * sahifa butunlay o'zicha ishlaydi va tashqi skript yuklamaydi.
 *
 * Ranglar `styles.css` dagi CSS o'zgaruvchilaridan olinadi (yorug'/qorong'i
 * rejim uchun alohida qiymatlar), shuning uchun bu yerda hex kod yozilmagan.
 */

import { el, svgEl } from './dom.js';
import { formatMinor } from '../shared/money.js';
import { formatHuman, formatShort } from '../shared/dates.js';
import type { NamedTotal, TrendPoint } from '../shared/types.js';

const NS_GAP = 2; // Qo'shni ustunlar orasidagi sirt oralig'i.

interface Size {
  width: number;
  height: number;
}

function niceCeil(value: number): number {
  if (value <= 0) return 1;
  const magnitude = Math.pow(10, Math.floor(Math.log10(value)));
  const normalized = value / magnitude;
  const step = normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10;
  return step * magnitude;
}

function compact(minor: number, currency: string): string {
  const abs = Math.abs(minor);
  const units = [
    { limit: 1e11, suffix: ' mlrd', divisor: 1e11 },
    { limit: 1e8, suffix: ' mln', divisor: 1e8 },
    { limit: 1e5, suffix: ' ming', divisor: 1e5 },
  ];
  for (const unit of units) {
    if (abs >= unit.limit) {
      const value = minor / unit.divisor;
      return `${value.toFixed(value >= 10 ? 0 : 1).replace('.', ',')}${unit.suffix}`;
    }
  }
  return formatMinor(minor, currency, { maxDecimals: 0 });
}

/** Ustun uchi 4px radius bilan yumaloqlangan, asos esa tekis. */
function barPath(x: number, y: number, width: number, height: number, radius = 4): string {
  const r = Math.max(0, Math.min(radius, width / 2, height));
  return [
    `M ${x} ${y + height}`,
    `L ${x} ${y + r}`,
    `Q ${x} ${y} ${x + r} ${y}`,
    `L ${x + width - r} ${y}`,
    `Q ${x + width} ${y} ${x + width} ${y + r}`,
    `L ${x + width} ${y + height}`,
    'Z',
  ].join(' ');
}

interface Tooltip {
  show(html: HTMLElement, x: number, y: number): void;
  hide(): void;
  node: HTMLElement;
}

function createTooltip(host: HTMLElement): Tooltip {
  const node = el('div', { class: 'viz-tooltip', attrs: { role: 'tooltip' } });
  node.hidden = true;
  host.appendChild(node);
  return {
    node,
    show(content, x, y) {
      node.replaceChildren(content);
      node.hidden = false;
      const bounds = host.getBoundingClientRect();
      const width = node.offsetWidth;
      const left = Math.min(Math.max(x - width / 2, 4), bounds.width - width - 4);
      node.style.left = `${left}px`;
      node.style.top = `${Math.max(y - node.offsetHeight - 10, 4)}px`;
    },
    hide() {
      node.hidden = true;
    },
  };
}

export interface TrendChartOptions {
  currency: string;
  incomeLabel?: string;
  expenseLabel?: string;
}

/**
 * Kirim/chiqim dinamikasi — guruhlangan ustunlar.
 * Ikki qator bo'lgani uchun legenda MAJBURIY (rang yagona ajratuvchi belgi bo'lib qolmasligi kerak).
 */
export function renderTrendChart(
  container: HTMLElement,
  points: readonly TrendPoint[],
  options: TrendChartOptions,
): void {
  container.replaceChildren();
  container.classList.add('viz');

  if (points.length === 0) {
    container.appendChild(el('p', { class: 'viz__empty', text: 'Bu davr uchun ma’lumot yo‘q' }));
    return;
  }

  const incomeLabel = options.incomeLabel ?? 'Kirim';
  const expenseLabel = options.expenseLabel ?? 'Chiqim';

  const legend = el('div', { class: 'viz-legend' }, [
    el('span', { class: 'viz-legend__item' }, [
      el('span', { class: 'viz-legend__swatch viz-legend__swatch--income' }),
      el('span', { text: incomeLabel }),
    ]),
    el('span', { class: 'viz-legend__item' }, [
      el('span', { class: 'viz-legend__swatch viz-legend__swatch--expense' }),
      el('span', { text: expenseLabel }),
    ]),
  ]);

  const size: Size = { width: 720, height: 260 };
  const padding = { top: 12, right: 8, bottom: 28, left: 56 };
  const plotWidth = size.width - padding.left - padding.right;
  const plotHeight = size.height - padding.top - padding.bottom;

  const maxValue = niceCeil(
    points.reduce((max, point) => Math.max(max, point.incomeMinor, point.expenseMinor), 0),
  );
  const scale = (value: number): number => (value / maxValue) * plotHeight;

  const svg = svgEl('svg', {
    class: 'viz-svg',
    viewBox: `0 0 ${size.width} ${size.height}`,
    preserveAspectRatio: 'none',
    role: 'img',
    'aria-label': `${incomeLabel} va ${expenseLabel} dinamikasi`,
  });

  // Retsessiv to'r va o'q yozuvlari.
  const ticks = 4;
  for (let i = 0; i <= ticks; i += 1) {
    const value = (maxValue / ticks) * i;
    const y = padding.top + plotHeight - scale(value);
    svg.appendChild(
      svgEl('line', {
        class: 'viz-grid',
        x1: padding.left,
        x2: size.width - padding.right,
        y1: y,
        y2: y,
      }),
    );
    svg.appendChild(
      svgEl('text', { class: 'viz-axis-label', x: padding.left - 8, y: y + 4, 'text-anchor': 'end' }, [
        compact(value, options.currency),
      ]),
    );
  }

  const slotWidth = plotWidth / points.length;
  const barWidth = Math.max(3, Math.min(18, slotWidth / 2 - NS_GAP));
  const tooltip = createTooltip(container);

  points.forEach((point, index) => {
    const slotStart = padding.left + slotWidth * index;
    const groupCenter = slotStart + slotWidth / 2;

    const bars: Array<{ value: number; className: string; offset: number }> = [
      { value: point.incomeMinor, className: 'viz-bar--income', offset: -barWidth - NS_GAP / 2 },
      { value: point.expenseMinor, className: 'viz-bar--expense', offset: NS_GAP / 2 },
    ];

    for (const bar of bars) {
      if (bar.value <= 0) continue;
      const height = Math.max(scale(bar.value), 2);
      svg.appendChild(
        svgEl('path', {
          class: `viz-bar ${bar.className}`,
          d: barPath(groupCenter + bar.offset, padding.top + plotHeight - height, barWidth, height),
        }),
      );
    }

    // Ko'rsatkichdan kengroq "sezgir" maydon.
    const hit = svgEl('rect', {
      class: 'viz-hit',
      x: slotStart,
      y: padding.top,
      width: slotWidth,
      height: plotHeight,
    });
    hit.addEventListener('mouseenter', () => {
      const content = el('div', {}, [
        el('div', { class: 'viz-tooltip__title', text: formatHuman(point.key) }),
        el('div', { class: 'viz-tooltip__row' }, [
          el('span', { class: 'viz-legend__swatch viz-legend__swatch--income' }),
          el('span', { text: `${incomeLabel}: ` }),
          el('strong', { text: formatMinor(point.incomeMinor, options.currency, { maxDecimals: 0 }) }),
        ]),
        el('div', { class: 'viz-tooltip__row' }, [
          el('span', { class: 'viz-legend__swatch viz-legend__swatch--expense' }),
          el('span', { text: `${expenseLabel}: ` }),
          el('strong', { text: formatMinor(point.expenseMinor, options.currency, { maxDecimals: 0 }) }),
        ]),
      ]);
      const rect = container.getBoundingClientRect();
      const scaleX = rect.width / size.width;
      tooltip.show(content, groupCenter * scaleX, (padding.top + plotHeight / 3) * (rect.height / size.height));
    });
    hit.addEventListener('mouseleave', () => tooltip.hide());
    svg.appendChild(hit);

    // O'q yozuvlari faqat tanlab qo'yiladi — har bir ustunga emas.
    const labelEvery = Math.ceil(points.length / 12);
    if (index % labelEvery === 0 || index === points.length - 1) {
      svg.appendChild(
        svgEl(
          'text',
          {
            class: 'viz-axis-label',
            x: groupCenter,
            y: size.height - 8,
            'text-anchor': 'middle',
          },
          [formatShort(point.key)],
        ),
      );
    }
  });

  container.appendChild(legend);
  container.appendChild(svg);
}

export interface BarListOptions {
  currency: string;
  limit?: number;
  emptyText?: string;
}

/**
 * Gorizontal ustunlar ro'yxati (savdo nuqtalari, kategoriyalar...).
 * Bitta qator bo'lgani uchun legenda kerak emas — sarlavha o'zi nomini aytadi.
 */
export function renderBarList(
  container: HTMLElement,
  items: readonly NamedTotal[],
  options: BarListOptions,
): void {
  container.replaceChildren();
  container.classList.add('viz-barlist');

  if (items.length === 0) {
    container.appendChild(
      el('p', { class: 'viz__empty', text: options.emptyText ?? 'Ma’lumot yo‘q' }),
    );
    return;
  }

  const limit = options.limit ?? 8;
  const visible = items.slice(0, limit);
  const rest = items.slice(limit);

  // Ortiqcha qatorlar yangi ranglarga bo'linmaydi — «Boshqalar» ga yig'iladi.
  if (rest.length > 0) {
    const total = rest.reduce((sum, item) => sum + item.totalMinor, 0);
    visible.push({ id: '__other__', name: `Boshqalar (${rest.length})`, totalMinor: total });
  }

  const max = visible.reduce((value, item) => Math.max(value, Math.abs(item.totalMinor)), 0) || 1;

  for (const item of visible) {
    const percent = Math.max((Math.abs(item.totalMinor) / max) * 100, 1);
    container.appendChild(
      el('div', { class: 'barlist__row' }, [
        el('div', { class: 'barlist__head' }, [
          el('span', { class: 'barlist__name', text: item.name, title: item.name }),
          el('span', {
            class: 'barlist__value',
            text: formatMinor(item.totalMinor, options.currency, { maxDecimals: 0 }),
          }),
        ]),
        el('div', { class: 'barlist__track' }, [
          el('div', { class: 'barlist__fill', style: { width: `${percent}%` } }),
        ]),
      ]),
    );
  }
}

/** KPI plitasi — bitta son uchun grafik shart emas. */
export interface StatTileOptions {
  label: string;
  /** Pul ko'rsatkichi — minor birlikda. `text` berilgan bo'lsa e'tiborsiz qoladi. */
  valueMinor?: number;
  currency?: string;
  /** Pul bo'lmagan ko'rsatkich (masalan yozuvlar soni) — tayyor matn. */
  text?: string;
  deltaPercent?: number | null;
  /** Xarajat uchun o'sish yomon belgi — rang teskari bo'ladi. */
  inverse?: boolean;
  hint?: string;
}

export function statTile(options: StatTileOptions): HTMLElement {
  const valueText =
    options.text ??
    formatMinor(options.valueMinor ?? 0, options.currency ?? 'UZS', { maxDecimals: 0 });

  const children: Array<HTMLElement | null> = [
    el('span', { class: 'stat__label', text: options.label }),
    el('strong', { class: 'stat__value', text: valueText }),
  ];

  if (options.deltaPercent === undefined || options.deltaPercent === null) {
    // Taqqoslash ma'nosiz bo'lgan ko'rsatkichda bo'sh "0%" qatorini chizmaymiz.
  } else {
    const rounded = Math.round(options.deltaPercent * 10) / 10;
    const good = options.inverse ? rounded <= 0 : rounded >= 0;
    const arrow = rounded > 0 ? '▲' : rounded < 0 ? '▼' : '■';
    children.push(
      el('span', {
        class: `stat__delta ${good ? 'stat__delta--good' : 'stat__delta--bad'}`,
        text: `${arrow} ${Math.abs(rounded).toFixed(1).replace('.', ',')}%`,
        title: 'Oldingi davrga nisbatan',
      }),
    );
  }

  if (options.hint) children.push(el('span', { class: 'stat__hint', text: options.hint }));
  return el('div', { class: 'stat' }, children);
}
