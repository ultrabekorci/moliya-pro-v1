/**
 * Kirim / Chiqim / O'tkazma ro'yxati va kiritish formasi.
 *
 * v1 dan farqlar:
 *  - tahrirlash va o'chirish tugmalari HAR UCHALA bo'limda ruxsatga qarab
 *    ko'rsatiladi (v1 da tekshiruv faqat "O'tkazma" jadvalida bor edi);
 *  - ro'yxat serverda filtrlanadi va sahifalanadi;
 *  - summalar hisobning o'z valyutasida ko'rsatiladi va tahrirlashda o'sha
 *    ko'rinishda qaytariladi — dollar chiqimi qayta konvertatsiya qilinmaydi;
 *  - to'lov turlari qotirib yozilmagan: hisoblar katalogdan keladi.
 */

import { call, errorField, errorMessage } from '../api.js';
import { debounce, el, fillSelect } from '../dom.js';
import {
  confirmDialog,
  dataTable,
  emptyState,
  fieldRow,
  openModal,
  section,
  setFieldError,
  spinner,
  toast,
  withBusy,
} from '../ui.js';
import {
  accountById,
  activeAccounts,
  activeCategories,
  activePoints,
  allowed,
  baseCurrency,
  categoryById,
  currencyOfAccount,
  pointById,
  state,
} from '../store.js';
import { formatMinor, minorToMajor } from '../../shared/money.js';
import { formatHuman } from '../../shared/dates.js';
import type { Action } from '../../shared/types.js';
import type { TransactionBatch, TransactionKind } from '../../shared/types.js';
import type { SaveTransactionInput } from '../../shared/api.js';

const KIND_LABELS: Record<TransactionKind, string> = {
  income: 'Kirim',
  expense: 'Chiqim',
  transfer: "O'tkazma",
};

const PAGE_SIZE = 50;

interface ViewState {
  kind: TransactionKind;
  search: string;
  start: string;
  end: string;
  offset: number;
}

const view: ViewState = { kind: 'income', search: '', start: '', end: '', offset: 0 };

let onDataChanged: (() => void) | null = null;

export function setTransactionsChangeHandler(handler: () => void): void {
  onDataChanged = handler;
}

function permit(kind: TransactionKind, action: Action): boolean {
  return allowed(kind, action);
}

export function renderTransactions(container: HTMLElement): void {
  container.replaceChildren();

  const tabs = el(
    'div',
    { class: 'tabs', attrs: { role: 'tablist' } },
    (Object.keys(KIND_LABELS) as TransactionKind[])
      .filter((kind) => permit(kind, 'view'))
      .map((kind) =>
        el('button', {
          class: `tab${view.kind === kind ? ' tab--active' : ''}`,
          type: 'button',
          text: KIND_LABELS[kind],
          attrs: { role: 'tab', 'aria-selected': String(view.kind === kind) },
          on: {
            click: () => {
              view.kind = kind;
              view.offset = 0;
              renderTransactions(container);
            },
          },
        }),
      ),
  );

  if (!permit(view.kind, 'view')) {
    const firstAllowed = (Object.keys(KIND_LABELS) as TransactionKind[]).find((kind) => permit(kind, 'view'));
    if (!firstAllowed) {
      container.replaceChildren(emptyState('Ruxsat yo‘q', 'Sizga yozuvlarni ko‘rish huquqi berilmagan.'));
      return;
    }
    view.kind = firstAllowed;
  }

  const searchInput = el('input', {
    class: 'input input--compact',
    type: 'search',
    placeholder: 'Izoh, nuqta, kategoriya…',
    value: view.search,
  });
  const startInput = el('input', { class: 'input input--compact', type: 'date', value: view.start });
  const endInput = el('input', { class: 'input input--compact', type: 'date', value: view.end });

  const listHost = el('div', {}, [spinner()]);

  const reload = (): void => {
    view.offset = 0;
    void loadList(listHost, container);
  };

  searchInput.addEventListener(
    'input',
    debounce(() => {
      view.search = searchInput.value;
      reload();
    }, 300),
  );
  startInput.addEventListener('change', () => {
    view.start = startInput.value;
    reload();
  });
  endInput.addEventListener('change', () => {
    view.end = endInput.value;
    reload();
  });

  const resetButton = el('button', {
    class: 'button button--ghost',
    type: 'button',
    text: 'Tozalash',
    on: {
      click: () => {
        view.search = '';
        view.start = '';
        view.end = '';
        renderTransactions(container);
      },
    },
  });

  const addButton = permit(view.kind, 'create')
    ? el('button', {
        class: 'button button--primary',
        type: 'button',
        text: `+ ${KIND_LABELS[view.kind]}`,
        on: { click: () => openTransactionForm(view.kind, null, container) },
      })
    : null;

  container.append(
    renderBalancePanel(),
    el('div', { class: 'toolbar' }, [tabs, addButton]),
    el('div', { class: 'filters' }, [searchInput, startInput, el('span', { text: '—' }), endInput, resetButton]),
    listHost,
  );

  void loadList(listHost, container);
}

// ---------------------------------------------------------------------------
// Balans paneli
// ---------------------------------------------------------------------------

function renderBalancePanel(): HTMLElement {
  const host = el('div', { class: 'balance-grid' });
  const dateInput = el('input', {
    class: 'input input--compact',
    type: 'date',
    value: state.today,
    max: state.today,
  });

  const paint = (): void => {
    host.replaceChildren(
      ...state.balances
        .filter((balance) => balance.showInBalance)
        .map((balance) =>
          el('div', { class: 'balance' }, [
            el('span', { class: 'balance__name', text: balance.name }),
            el('strong', {
              class: `balance__value${balance.balanceMinor < 0 ? ' balance__value--negative' : ''}`,
              text: formatMinor(balance.balanceMinor, balance.currency, { withSymbol: true }),
            }),
          ]),
        ),
    );
    if (host.childElementCount === 0) {
      host.replaceChildren(emptyState('Hisoblar yo‘q', 'Sozlamalar → Hisoblar bo‘limidan qo‘shing.'));
    }
  };

  dateInput.addEventListener('change', () => {
    void (async () => {
      try {
        state.balances = await call('tx.balances', { asOf: dateInput.value || null });
        paint();
      } catch (error) {
        toast(errorMessage(error), 'error');
      }
    })();
  });

  paint();
  return section('Hisoblar qoldig‘i', [host], [el('span', { class: 'card__hint', text: 'Sana:' }), dateInput]);
}

// ---------------------------------------------------------------------------
// Ro'yxat
// ---------------------------------------------------------------------------

function amountOfBatch(batch: TransactionBatch): string {
  if (batch.kind === 'transfer') {
    const entry = batch.entries[0];
    if (!entry) return '—';
    const out = formatMinor(entry.amountMinor, entry.currency, { withSymbol: true });
    const inValue =
      entry.counterAmountMinor !== null && entry.counterCurrency
        ? formatMinor(entry.counterAmountMinor, entry.counterCurrency, { withSymbol: true })
        : out;
    return out === inValue ? out : `${out} → ${inValue}`;
  }

  // Bir guruhda turli valyutalar bo'lishi mumkin — har birini alohida ko'rsatamiz.
  const byCurrency = new Map<string, number>();
  for (const entry of batch.entries) {
    byCurrency.set(entry.currency, (byCurrency.get(entry.currency) ?? 0) + entry.amountMinor);
  }
  return Array.from(byCurrency.entries())
    .map(([currency, minor]) => formatMinor(minor, currency, { withSymbol: true }))
    .join(' + ');
}

function accountsOfBatch(batch: TransactionBatch): string {
  if (batch.kind === 'transfer') {
    const entry = batch.entries[0];
    if (!entry) return '—';
    return `${accountById(entry.accountId)?.name ?? '—'} → ${accountById(entry.counterAccountId)?.name ?? '—'}`;
  }
  return batch.entries.map((entry) => accountById(entry.accountId)?.name ?? '—').join(', ');
}

async function loadList(host: HTMLElement, container: HTMLElement): Promise<void> {
  host.replaceChildren(spinner());

  const filter: Record<string, unknown> = { kinds: [view.kind] };
  if (view.search.trim() !== '') filter.search = view.search.trim();
  if (view.start && view.end) filter.range = { start: view.start, end: view.end };
  else if (view.start) filter.range = { start: view.start, end: '2999-12-31' };
  else if (view.end) filter.range = { start: '1970-01-01', end: view.end };

  try {
    const page = await call('tx.listGrouped', {
      filter: filter as never,
      offset: view.offset,
      limit: PAGE_SIZE,
    });
    renderList(host, container, page.items, page.total);
  } catch (error) {
    host.replaceChildren(emptyState('Ro‘yxatni yuklab bo‘lmadi', errorMessage(error)));
  }
}

function renderList(
  host: HTMLElement,
  container: HTMLElement,
  batches: readonly TransactionBatch[],
  total: number,
): void {
  const canEdit = permit(view.kind, 'edit');
  const canDelete = permit(view.kind, 'delete');

  const actionsCell = (batch: TransactionBatch): HTMLElement => {
    const wrap = el('div', { class: 'row-actions' });
    if (canEdit) {
      wrap.appendChild(
        el('button', {
          class: 'icon-button',
          type: 'button',
          text: '✎',
          title: 'Tahrirlash',
          on: { click: () => openTransactionForm(batch.kind, batch, container) },
        }),
      );
    }
    if (canDelete) {
      wrap.appendChild(
        el('button', {
          class: 'icon-button icon-button--danger',
          type: 'button',
          text: '🗑',
          title: "O'chirish",
          on: { click: () => void removeBatch(batch, container) },
        }),
      );
    }
    return wrap;
  };

  const columns = [
    { header: 'Sana', width: '130px', render: (batch: TransactionBatch) => formatHuman(batch.date) },
    view.kind === 'income'
      ? { header: 'Savdo nuqtasi', render: (batch: TransactionBatch) => pointById(batch.pointId)?.name ?? '—' }
      : view.kind === 'expense'
        ? { header: 'Kategoriya', render: (batch: TransactionBatch) => categoryById(batch.categoryId)?.name ?? '—' }
        : { header: 'Yo‘nalish', render: (batch: TransactionBatch) => accountsOfBatch(batch) },
    ...(view.kind === 'transfer'
      ? []
      : [{ header: 'Hisob', render: (batch: TransactionBatch) => accountsOfBatch(batch) }]),
    {
      header: 'Summa',
      align: 'right' as const,
      render: (batch: TransactionBatch) =>
        el('span', { class: `amount amount--${batch.kind}`, text: amountOfBatch(batch) }),
    },
    { header: 'Izoh', render: (batch: TransactionBatch) => batch.note || '—' },
    ...(canEdit || canDelete
      ? [{ header: '', width: '90px', align: 'right' as const, render: actionsCell }]
      : []),
  ];

  const table = dataTable(
    columns,
    batches,
    emptyState('Yozuv topilmadi', 'Filtrlarni o‘zgartiring yoki yangi yozuv qo‘shing.'),
  );

  host.replaceChildren(table);

  if (total > PAGE_SIZE) {
    const from = view.offset + 1;
    const to = Math.min(view.offset + PAGE_SIZE, total);
    host.appendChild(
      el('div', { class: 'pager' }, [
        el('button', {
          class: 'button button--ghost',
          type: 'button',
          text: '← Oldingi',
          disabled: view.offset === 0,
          on: {
            click: () => {
              view.offset = Math.max(0, view.offset - PAGE_SIZE);
              void loadList(host, container);
            },
          },
        }),
        el('span', { class: 'pager__info', text: `${from}–${to} / ${total}` }),
        el('button', {
          class: 'button button--ghost',
          type: 'button',
          text: 'Keyingi →',
          disabled: to >= total,
          on: {
            click: () => {
              view.offset += PAGE_SIZE;
              void loadList(host, container);
            },
          },
        }),
      ]),
    );
  }
}

async function removeBatch(batch: TransactionBatch, container: HTMLElement): Promise<void> {
  const confirmed = await confirmDialog({
    title: 'Yozuvni o‘chirish',
    message: `${formatHuman(batch.date)} sanasidagi yozuv o‘chiriladi. Tarixda saqlanadi, lekin hisob-kitobga kirmaydi.`,
    confirmText: "O'chirish",
    danger: true,
  });
  if (!confirmed) return;

  try {
    await call('tx.delete', { id: batch.batchId });
    toast('Yozuv o‘chirildi', 'success');
    onDataChanged?.();
    renderTransactions(container);
  } catch (error) {
    toast(errorMessage(error), 'error');
  }
}

// ---------------------------------------------------------------------------
// Kiritish / tahrirlash formasi
// ---------------------------------------------------------------------------

function amountInput(name: string, value = ''): HTMLInputElement {
  return el('input', {
    class: 'input',
    name,
    value,
    placeholder: '0',
    inputMode: 'decimal',
    autocomplete: 'off',
  });
}

function openTransactionForm(
  kind: TransactionKind,
  batch: TransactionBatch | null,
  container: HTMLElement,
): void {
  const form = el('form', { class: 'form' });
  const isEdit = batch !== null;

  const dateInput = el('input', {
    class: 'input',
    type: 'date',
    name: 'date',
    value: batch?.date ?? state.today,
    required: true,
  });
  const noteInput = el('textarea', {
    class: 'input',
    name: 'note',
    rows: 2,
    maxLength: state.config?.maxNoteLength ?? 500,
    value: batch?.note ?? '',
  }) as unknown as HTMLTextAreaElement;

  const fields: HTMLElement[] = [fieldRow('Sana', dateInput)];
  const accounts = activeAccounts();

  if (accounts.length === 0) {
    openModal({
      title: 'Hisoblar yo‘q',
      body: [
        el('p', {
          class: 'modal__message',
          text: 'Avval Sozlamalar → Hisoblar bo‘limida kamida bitta hisob (Naqd, Plastik…) yarating.',
        }),
      ],
    });
    return;
  }

  // --- Kirim: bir nechta hisobga bir vaqtda ---------------------------------
  const entryInputs = new Map<string, HTMLInputElement>();
  let pointSelect: HTMLSelectElement | null = null;
  let categorySelect: HTMLSelectElement | null = null;
  let accountSelect: HTMLSelectElement | null = null;
  let counterSelect: HTMLSelectElement | null = null;
  let singleAmount: HTMLInputElement | null = null;
  let counterAmount: HTMLInputElement | null = null;
  let counterAmountRow: HTMLElement | null = null;

  if (kind === 'income') {
    pointSelect = el('select', { class: 'input', name: 'pointId' });
    fillSelect(
      pointSelect,
      activePoints().map((point) => ({ value: point.id, label: point.name })),
      batch?.pointId ?? '',
      'Tanlanmagan',
    );
    fields.push(fieldRow('Savdo nuqtasi', pointSelect, 'Firma shu nuqta orqali aniqlanadi'));

    const grid = el('div', { class: 'amount-grid' });
    for (const account of accounts) {
      const existing = batch?.entries.find((entry) => entry.accountId === account.id);
      const input = amountInput(
        `amount_${account.id}`,
        existing ? String(minorToMajor(existing.amountMinor, existing.currency)) : '',
      );
      entryInputs.set(account.id, input);
      grid.appendChild(
        el('label', { class: 'amount-grid__row' }, [
          el('span', { class: 'amount-grid__label' }, [
            el('span', { text: account.name }),
            el('span', { class: 'amount-grid__currency', text: account.currency }),
          ]),
          input,
        ]),
      );
    }
    fields.push(el('div', { class: 'field' }, [el('span', { class: 'field__label', text: 'Summalar' }), grid]));
  } else {
    accountSelect = el('select', { class: 'input', name: 'accountId', required: true });
    fillSelect(
      accountSelect,
      accounts.map((account) => ({ value: account.id, label: `${account.name} (${account.currency})` })),
      batch?.entries[0]?.accountId ?? '',
      'Tanlang…',
    );

    const firstEntry = batch?.entries[0];
    singleAmount = amountInput(
      'amount',
      firstEntry ? String(minorToMajor(firstEntry.amountMinor, firstEntry.currency)) : '',
    );

    const currencyHint = el('span', { class: 'field__hint' });
    const updateHint = (): void => {
      const currency = currencyOfAccount(accountSelect?.value);
      currencyHint.textContent = `Summa ${currency} da kiritiladi`;
      if (kind === 'transfer' && counterAmountRow) {
        const from = currencyOfAccount(accountSelect?.value);
        const to = currencyOfAccount(counterSelect?.value);
        counterAmountRow.hidden = from === to;
      }
    };
    accountSelect.addEventListener('change', updateHint);

    if (kind === 'expense') {
      categorySelect = el('select', { class: 'input', name: 'categoryId', required: true });
      fillSelect(
        categorySelect,
        activeCategories('expense').map((category) => ({ value: category.id, label: category.name })),
        batch?.categoryId ?? '',
        'Tanlang…',
      );
      fields.push(fieldRow('Kategoriya', categorySelect));

      const expensePoint = el('select', { class: 'input', name: 'pointId' });
      fillSelect(
        expensePoint,
        activePoints().map((point) => ({ value: point.id, label: point.name })),
        batch?.pointId ?? '',
        'Umumiy (firmaga bog‘lanmagan)',
      );
      pointSelect = expensePoint;
      fields.push(
        fieldRow(
          'Savdo nuqtasi',
          expensePoint,
          'Bo‘sh qoldirilsa xarajat umumiy hisoblanadi va sozlamadagi qoidaga ko‘ra taqsimlanadi',
        ),
      );

      fields.push(fieldRow('Hisob', accountSelect));
      fields.push(fieldRow('Summa', singleAmount));
      fields.push(el('div', { class: 'field' }, [currencyHint]));
    } else {
      counterSelect = el('select', { class: 'input', name: 'counterAccountId', required: true });
      fillSelect(
        counterSelect,
        accounts.map((account) => ({ value: account.id, label: `${account.name} (${account.currency})` })),
        batch?.entries[0]?.counterAccountId ?? '',
        'Tanlang…',
      );
      counterSelect.addEventListener('change', updateHint);

      const entry = batch?.entries[0];
      counterAmount = amountInput(
        'counterAmount',
        entry?.counterAmountMinor !== null && entry?.counterAmountMinor !== undefined && entry.counterCurrency
          ? String(minorToMajor(entry.counterAmountMinor, entry.counterCurrency))
          : '',
      );
      counterAmountRow = fieldRow('Tushadigan summa', counterAmount, 'Valyutalar har xil bo‘lsa majburiy');

      fields.push(fieldRow('Qayerdan', accountSelect));
      fields.push(fieldRow('Chiqadigan summa', singleAmount));
      fields.push(fieldRow('Qayerga', counterSelect));
      fields.push(counterAmountRow);
      fields.push(el('div', { class: 'field' }, [currencyHint]));
    }
    updateHint();
  }

  fields.push(fieldRow('Izoh', noteInput as unknown as HTMLElement));
  fields.push(el('p', { class: 'form-error' }));

  form.append(...fields);
  (form.querySelector('.form-error') as HTMLElement).hidden = true;

  const submit = el('button', {
    class: 'button button--primary',
    type: 'submit',
    text: isEdit ? 'Saqlash' : 'Qo‘shish',
  });

  const modal = openModal({
    title: `${KIND_LABELS[kind]}${isEdit ? ' — tahrirlash' : ''}`,
    body: [form],
    footer: [
      el('button', {
        class: 'button button--ghost',
        type: 'button',
        text: 'Bekor qilish',
        on: { click: () => modal.close() },
      }),
      submit,
    ],
    wide: kind === 'income',
  });

  form.addEventListener('submit', (event) => {
    event.preventDefault();

    const input: SaveTransactionInput = {
      batchId: batch?.batchId ?? null,
      kind,
      date: dateInput.value,
      note: noteInput.value,
      pointId: pointSelect?.value || null,
      categoryId: categorySelect?.value || null,
    };

    if (kind === 'income') {
      input.entries = Array.from(entryInputs.entries())
        .map(([accountId, field]) => ({ accountId, amount: field.value.trim() }))
        .filter((entry) => entry.amount !== '');
    } else {
      input.accountId = accountSelect?.value ?? '';
      input.amount = singleAmount?.value ?? '';
      if (kind === 'transfer') {
        input.counterAccountId = counterSelect?.value ?? '';
        const counterValue = counterAmount?.value.trim() ?? '';
        input.counterAmount = counterValue === '' ? null : counterValue;
      }
    }

    void withBusy(submit, async () => {
      try {
        await call('tx.save', input);
        modal.close();
        toast(isEdit ? 'Yozuv yangilandi' : 'Yozuv qo‘shildi', 'success');
        onDataChanged?.();
        renderTransactions(container);
      } catch (error) {
        setFieldError(form, errorField(error), errorMessage(error));
      }
    });
  });
}

/** Balans panelini yangilash uchun (yozuv o'zgargandan keyin chaqiriladi). */
export async function refreshBalances(): Promise<void> {
  try {
    state.balances = await call('tx.balances', {});
  } catch {
    // Balans yangilanmasa ham ro'yxat ishlayveradi.
  }
}

export function currentBaseCurrency(): string {
  return baseCurrency();
}
