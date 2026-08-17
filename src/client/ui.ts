/**
 * Oddiy UI komponentlari: toast, modal, tasdiqlash oynasi, yuklanish holati.
 * Tashqi kutubxona ishlatilmaydi — sahifa CDN'ga umuman bog'liq emas.
 */

import { append, clear, el, replace } from './dom.js';
import type { Child } from './dom.js';

export type ToastKind = 'success' | 'error' | 'info' | 'warning';

let toastHost: HTMLElement | null = null;

function host(): HTMLElement {
  if (!toastHost) {
    toastHost = el('div', { class: 'toast-host', attrs: { role: 'status', 'aria-live': 'polite' } });
    document.body.appendChild(toastHost);
  }
  return toastHost;
}

export function toast(message: string, kind: ToastKind = 'info', durationMs = 3200): void {
  const node = el('div', { class: `toast toast--${kind}` }, [
    el('span', { class: 'toast__dot' }),
    el('span', { text: message }),
  ]);
  host().appendChild(node);
  window.setTimeout(() => {
    node.classList.add('toast--leaving');
    window.setTimeout(() => node.remove(), 220);
  }, durationMs);
}

export interface ModalHandle {
  close(): void;
  readonly body: HTMLElement;
  readonly footer: HTMLElement;
}

let openModals = 0;

export function openModal(options: {
  title: string;
  body: Child[];
  footer?: Child[];
  onClose?: () => void;
  wide?: boolean;
}): ModalHandle {
  const body = el('div', { class: 'modal__body' }, options.body);
  const footer = el('div', { class: 'modal__footer' }, options.footer ?? []);

  const close = (): void => {
    backdrop.classList.add('modal-backdrop--leaving');
    window.setTimeout(() => backdrop.remove(), 160);
    document.removeEventListener('keydown', onKeyDown);
    openModals = Math.max(0, openModals - 1);
    if (openModals === 0) document.body.classList.remove('modal-open');
    options.onClose?.();
  };

  const onKeyDown = (event: KeyboardEvent): void => {
    if (event.key === 'Escape') close();
  };

  const dialog = el(
    'div',
    {
      class: `modal${options.wide ? ' modal--wide' : ''}`,
      attrs: { role: 'dialog', 'aria-modal': 'true', 'aria-label': options.title },
    },
    [
      el('div', { class: 'modal__header' }, [
        el('h2', { class: 'modal__title', text: options.title }),
        el('button', {
          class: 'icon-button',
          type: 'button',
          text: '✕',
          attrs: { 'aria-label': 'Yopish' },
          on: { click: close },
        }),
      ]),
      body,
      footer,
    ],
  );

  const backdrop = el(
    'div',
    {
      class: 'modal-backdrop',
      on: {
        click: (event: MouseEvent) => {
          if (event.target === backdrop) close();
        },
      },
    },
    [dialog],
  );

  document.body.appendChild(backdrop);
  document.body.classList.add('modal-open');
  openModals += 1;
  document.addEventListener('keydown', onKeyDown);

  const firstField = dialog.querySelector<HTMLElement>('input, select, textarea, button');
  firstField?.focus();

  return { close, body, footer };
}

export function confirmDialog(options: {
  title: string;
  message: string;
  confirmText?: string;
  cancelText?: string;
  danger?: boolean;
}): Promise<boolean> {
  return new Promise((resolve) => {
    let settled = false;
    const finish = (value: boolean): void => {
      if (settled) return;
      settled = true;
      modal.close();
      resolve(value);
    };

    const modal = openModal({
      title: options.title,
      body: [el('p', { class: 'modal__message', text: options.message })],
      footer: [
        el('button', {
          class: 'button button--ghost',
          type: 'button',
          text: options.cancelText ?? 'Bekor qilish',
          on: { click: () => finish(false) },
        }),
        el('button', {
          class: `button ${options.danger ? 'button--danger' : 'button--primary'}`,
          type: 'button',
          text: options.confirmText ?? 'Ha',
          on: { click: () => finish(true) },
        }),
      ],
      onClose: () => {
        if (!settled) {
          settled = true;
          resolve(false);
        }
      },
    });
  });
}

/** Tugmani chaqiruv davomida bloklaydi — ikki marta yuborishning oldini oladi. */
export async function withBusy<T>(button: HTMLButtonElement | null, task: () => Promise<T>): Promise<T> {
  const original = button?.textContent ?? '';
  if (button) {
    button.disabled = true;
    button.classList.add('is-busy');
  }
  try {
    return await task();
  } finally {
    if (button) {
      button.disabled = false;
      button.classList.remove('is-busy');
      if (button.textContent !== original) button.textContent = original;
    }
  }
}

export function spinner(label = 'Yuklanmoqda…'): HTMLElement {
  return el('div', { class: 'loading' }, [el('span', { class: 'loading__spinner' }), el('span', { text: label })]);
}

export function emptyState(message: string, hint?: string): HTMLElement {
  return el('div', { class: 'empty' }, [
    el('p', { class: 'empty__title', text: message }),
    hint ? el('p', { class: 'empty__hint', text: hint }) : null,
  ]);
}

export function fieldRow(label: string, control: HTMLElement, hint?: string): HTMLElement {
  const id = control.id || `f_${Math.random().toString(36).slice(2, 9)}`;
  control.id = id;
  return el('label', { class: 'field', attrs: { for: id } }, [
    el('span', { class: 'field__label', text: label }),
    control,
    hint ? el('span', { class: 'field__hint', text: hint }) : null,
  ]);
}

export function setFieldError(form: HTMLElement, field: string | null, message: string): void {
  for (const node of Array.from(form.querySelectorAll('.field--invalid'))) {
    node.classList.remove('field--invalid');
  }
  const box = form.querySelector<HTMLElement>('.form-error');
  if (box) {
    box.textContent = message;
    box.hidden = message === '';
  }
  if (!field) return;
  const control = form.querySelector<HTMLElement>(`[name="${field}"]`);
  control?.closest('.field')?.classList.add('field--invalid');
  control?.focus();
}

export function section(title: string, children: Child[], actions?: Child[]): HTMLElement {
  return el('section', { class: 'card' }, [
    el('div', { class: 'card__header' }, [
      el('h2', { class: 'card__title', text: title }),
      actions ? el('div', { class: 'card__actions' }, actions) : null,
    ]),
    el('div', { class: 'card__body' }, children),
  ]);
}

export interface TableColumn<T> {
  header: string;
  align?: 'left' | 'right' | 'center';
  width?: string;
  render(row: T, index: number): Child;
}

export function dataTable<T>(columns: ReadonlyArray<TableColumn<T>>, rows: readonly T[], empty: HTMLElement): HTMLElement {
  if (rows.length === 0) return empty;

  const head = el('thead', {}, [
    el(
      'tr',
      {},
      columns.map((column) =>
        el('th', {
          text: column.header,
          class: column.align ? `align-${column.align}` : '',
          style: column.width ? { width: column.width } : {},
        }),
      ),
    ),
  ]);

  const body = el('tbody');
  rows.forEach((row, index) => {
    const tr = el('tr');
    for (const column of columns) {
      const cell = el('td', { class: column.align ? `align-${column.align}` : '' });
      append(cell, [column.render(row, index)]);
      tr.appendChild(cell);
    }
    body.appendChild(tr);
  });

  return el('div', { class: 'table-wrap' }, [el('table', { class: 'table' }, [head, body])]);
}

/** Ko'rinishni almashtirish uchun konteynerni tozalab, yangi tugunlarni qo'yadi. */
export function mount(container: HTMLElement, children: Child[]): void {
  replace(container, children);
}

export function detach(node: HTMLElement): void {
  clear(node);
}
