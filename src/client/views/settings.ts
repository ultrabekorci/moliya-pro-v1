/**
 * Sozlamalar: ma'lumotnomalar, xodimlar, umumiy parametrlar, valyuta kursi.
 *
 * Bu yerda v1 dagi "Greenpen/Smartmiz kodga qotirib yozilgan" muammosi hal
 * bo'ladi — firmalar, savdo nuqtalari, hisoblar va xarajat taqsimoti qoidasi
 * to'liq shu oynadan boshqariladi.
 */

import { call, errorField, errorMessage } from '../api.js';
import { el, fillSelect } from '../dom.js';
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
import { allowed, firmById, state } from '../store.js';
import { ACTION_LABELS, RESOURCE_LABELS, ROLE_LABELS } from '../../shared/permissions.js';
import { ACTIONS, EXPENSE_ALLOCATIONS, RESOURCES, ROLES } from '../../shared/types.js';
import type {
  Account,
  Action,
  AppConfig,
  AuditEntry,
  Category,
  ExpenseAllocation,
  Firm,
  Permissions,
  Point,
  PublicUser,
  Resource,
  Role,
} from '../../shared/types.js';
import { checkPasswordStrength } from '../../shared/validation.js';

type SettingsTab = 'catalog' | 'users' | 'general' | 'audit';

const TAB_LABELS: Record<SettingsTab, string> = {
  catalog: 'Ma’lumotnomalar',
  users: 'Xodimlar',
  general: 'Umumiy',
  audit: 'Amallar tarixi',
};

const ALLOCATION_LABELS: Record<ExpenseAllocation, string> = {
  direct: 'Taqsimlanmasin (faqat umumiy hisobotda)',
  proRataIncome: 'Daromadga proporsional',
  equal: 'Firmalar orasida teng',
};

let activeTab: SettingsTab = 'catalog';

let onCatalogChanged: (() => void | Promise<void>) | null = null;

export function setSettingsChangeHandler(handler: () => void | Promise<void>): void {
  onCatalogChanged = handler;
}

function availableTabs(): SettingsTab[] {
  const tabs: SettingsTab[] = [];
  if (allowed('catalog', 'view')) tabs.push('catalog');
  if (allowed('users', 'view')) tabs.push('users');
  if (allowed('config', 'view')) tabs.push('general');
  if (allowed('audit', 'view')) tabs.push('audit');
  return tabs;
}

export function renderSettings(container: HTMLElement): void {
  const tabs = availableTabs();
  if (tabs.length === 0) {
    container.replaceChildren(emptyState('Ruxsat yo‘q', 'Sozlamalarni ko‘rish huquqi berilmagan.'));
    return;
  }
  if (!tabs.includes(activeTab)) activeTab = tabs[0]!;

  const body = el('div', { class: 'settings__body' }, [spinner()]);

  container.replaceChildren(
    el(
      'div',
      { class: 'tabs', attrs: { role: 'tablist' } },
      tabs.map((tab) =>
        el('button', {
          class: `tab${activeTab === tab ? ' tab--active' : ''}`,
          type: 'button',
          text: TAB_LABELS[tab],
          attrs: { role: 'tab', 'aria-selected': String(activeTab === tab) },
          on: {
            click: () => {
              activeTab = tab;
              renderSettings(container);
            },
          },
        }),
      ),
    ),
    body,
  );

  switch (activeTab) {
    case 'catalog':
      void renderCatalogTab(body, container);
      break;
    case 'users':
      void renderUsersTab(body);
      break;
    case 'general':
      void renderGeneralTab(body);
      break;
    case 'audit':
      void renderAuditTab(body);
      break;
    default:
      break;
  }
}

async function reloadCatalog(): Promise<void> {
  state.catalog = await call('catalog.get', {});
  await onCatalogChanged?.();
}

// ---------------------------------------------------------------------------
// Ma'lumotnomalar
// ---------------------------------------------------------------------------

function statusBadge(active: boolean): HTMLElement {
  return el('span', {
    class: `badge ${active ? 'badge--ok' : 'badge--muted'}`,
    text: active ? 'Faol' : 'Nofaol',
  });
}

function catalogActions(onEdit: () => void, onDelete: () => Promise<void>): HTMLElement {
  const wrap = el('div', { class: 'row-actions' });
  if (allowed('catalog', 'edit')) {
    wrap.appendChild(
      el('button', { class: 'icon-button', type: 'button', text: '✎', title: 'Tahrirlash', on: { click: onEdit } }),
    );
  }
  if (allowed('catalog', 'delete')) {
    wrap.appendChild(
      el('button', {
        class: 'icon-button icon-button--danger',
        type: 'button',
        text: '🗑',
        title: "O'chirish",
        on: { click: () => void onDelete() },
      }),
    );
  }
  return wrap;
}

async function renderCatalogTab(body: HTMLElement, container: HTMLElement): Promise<void> {
  body.replaceChildren(spinner());
  try {
    await reloadCatalog();
  } catch (error) {
    body.replaceChildren(emptyState('Yuklab bo‘lmadi', errorMessage(error)));
    return;
  }

  const canEdit = allowed('catalog', 'edit');
  const refresh = (): void => renderSettings(container);

  const addButton = (label: string, onClick: () => void): HTMLElement | null =>
    canEdit
      ? el('button', { class: 'button button--primary button--small', type: 'button', text: label, on: { click: onClick } })
      : null;

  // --- Firmalar -------------------------------------------------------------
  const firmsTable = dataTable<Firm>(
    [
      { header: 'Nom', render: (firm) => firm.name },
      { header: 'Umumiy xarajat', render: (firm) => ALLOCATION_LABELS[firm.expenseAllocation] },
      { header: 'Holat', render: (firm) => statusBadge(firm.active) },
      {
        header: '',
        align: 'right',
        width: '90px',
        render: (firm) =>
          catalogActions(
            () => openFirmForm(firm, refresh),
            () => deleteCatalogItem('catalog.deleteFirm', firm.id, firm.name, refresh),
          ),
      },
    ],
    state.catalog.firms,
    emptyState('Firmalar yo‘q', 'Birinchi firmani qo‘shing — savdo nuqtalari shunga bog‘lanadi.'),
  );

  // --- Hisoblar -------------------------------------------------------------
  const accountsTable = dataTable<Account>(
    [
      { header: 'Nom', render: (account) => account.name },
      { header: 'Valyuta', render: (account) => account.currency },
      { header: 'Balansda', render: (account) => (account.showInBalance ? 'Ha' : "Yo'q") },
      { header: 'Holat', render: (account) => statusBadge(account.active) },
      {
        header: '',
        align: 'right',
        width: '90px',
        render: (account) =>
          catalogActions(
            () => openAccountForm(account, refresh),
            () => deleteCatalogItem('catalog.deleteAccount', account.id, account.name, refresh),
          ),
      },
    ],
    state.catalog.accounts,
    emptyState('Hisoblar yo‘q', 'Naqd, Plastik, Bank, Dollar kassa… — nechta kerak bo‘lsa shuncha.'),
  );

  // --- Savdo nuqtalari ------------------------------------------------------
  const pointsTable = dataTable<Point>(
    [
      { header: 'Nom', render: (point) => point.name },
      { header: 'Firma', render: (point) => firmById(point.firmId)?.name ?? '—' },
      {
        header: 'Daromadga kiradi',
        render: (point) => (point.excludeFromRevenue ? "Yo'q" : 'Ha'),
      },
      { header: 'Holat', render: (point) => statusBadge(point.active) },
      {
        header: '',
        align: 'right',
        width: '90px',
        render: (point) =>
          catalogActions(
            () => openPointForm(point, refresh),
            () => deleteCatalogItem('catalog.deletePoint', point.id, point.name, refresh),
          ),
      },
    ],
    state.catalog.points,
    emptyState('Savdo nuqtalari yo‘q'),
  );

  // --- Kategoriyalar --------------------------------------------------------
  const categoriesTable = dataTable<Category>(
    [
      { header: 'Nom', render: (category) => category.name },
      {
        header: 'Turi',
        render: (category) =>
          category.kind === 'income' ? 'Kirim' : category.kind === 'expense' ? 'Chiqim' : 'Ikkalasi',
      },
      { header: 'Holat', render: (category) => statusBadge(category.active) },
      {
        header: '',
        align: 'right',
        width: '90px',
        render: (category) =>
          catalogActions(
            () => openCategoryForm(category, refresh),
            () => deleteCatalogItem('catalog.deleteCategory', category.id, category.name, refresh),
          ),
      },
    ],
    state.catalog.categories,
    emptyState('Kategoriyalar yo‘q'),
  );

  body.replaceChildren(
    section('Firmalar', [firmsTable], [addButton('+ Firma', () => openFirmForm(null, refresh))]),
    section('Hisoblar (to‘lov turlari)', [accountsTable], [addButton('+ Hisob', () => openAccountForm(null, refresh))]),
    section('Savdo nuqtalari', [pointsTable], [addButton('+ Nuqta', () => openPointForm(null, refresh))]),
    section(
      'Kategoriyalar',
      [categoriesTable],
      [addButton('+ Kategoriya', () => openCategoryForm(null, refresh))],
    ),
  );
}

async function deleteCatalogItem(
  method: 'catalog.deleteFirm' | 'catalog.deleteAccount' | 'catalog.deletePoint' | 'catalog.deleteCategory',
  id: string,
  name: string,
  refresh: () => void,
): Promise<void> {
  const confirmed = await confirmDialog({
    title: 'O‘chirish',
    message: `«${name}» o‘chirilsinmi?`,
    confirmText: "O'chirish",
    danger: true,
  });
  if (!confirmed) return;

  try {
    await call(method, { id });
    toast('O‘chirildi', 'success');
    refresh();
  } catch (error) {
    // Ishlatilayotgan yozuvni o'chirish taqiqlangan — server sababini tushuntiradi.
    toast(errorMessage(error), 'error', 6000);
  }
}

function nameInput(value = ''): HTMLInputElement {
  return el('input', { class: 'input', name: 'name', value, required: true, maxLength: 60 });
}

function checkbox(name: string, label: string, checked: boolean): HTMLElement {
  const input = el('input', { type: 'checkbox', name, checked });
  return el('label', { class: 'checkbox' }, [input, el('span', { text: label })]);
}

function checkboxValue(form: HTMLElement, name: string): boolean {
  return form.querySelector<HTMLInputElement>(`[name="${name}"]`)?.checked ?? false;
}

function simpleForm(options: {
  title: string;
  fields: HTMLElement[];
  submit(form: HTMLFormElement): Promise<void>;
}): void {
  const form = el('form', { class: 'form' }, [...options.fields, el('p', { class: 'form-error' })]);
  (form.querySelector('.form-error') as HTMLElement).hidden = true;

  const submitButton = el('button', { class: 'button button--primary', type: 'submit', text: 'Saqlash' });
  const modal = openModal({
    title: options.title,
    body: [form],
    footer: [
      el('button', {
        class: 'button button--ghost',
        type: 'button',
        text: 'Bekor qilish',
        on: { click: () => modal.close() },
      }),
      submitButton,
    ],
  });

  form.addEventListener('submit', (event) => {
    event.preventDefault();
    void withBusy(submitButton, async () => {
      try {
        await options.submit(form);
        modal.close();
        toast('Saqlandi', 'success');
      } catch (error) {
        setFieldError(form, errorField(error), errorMessage(error));
      }
    });
  });
}

function openFirmForm(firm: Firm | null, refresh: () => void): void {
  const name = nameInput(firm?.name ?? '');
  const allocation = el('select', { class: 'input', name: 'expenseAllocation' });
  fillSelect(
    allocation,
    EXPENSE_ALLOCATIONS.map((value) => ({ value, label: ALLOCATION_LABELS[value] })),
    firm?.expenseAllocation ?? 'proRataIncome',
  );

  simpleForm({
    title: firm ? 'Firmani tahrirlash' : 'Yangi firma',
    fields: [
      fieldRow('Nom', name),
      fieldRow('Umumiy xarajatlar', allocation, 'Firmaga bog‘lanmagan xarajatlar qanday taqsimlansin'),
      checkbox('active', 'Faol', firm?.active ?? true),
    ],
    submit: async (form) => {
      await call('catalog.saveFirm', {
        id: firm?.id ?? null,
        name: name.value,
        active: checkboxValue(form, 'active'),
        expenseAllocation: allocation.value as ExpenseAllocation,
      });
      refresh();
    },
  });
}

function openAccountForm(account: Account | null, refresh: () => void): void {
  const name = nameInput(account?.name ?? '');
  const currency = el('input', {
    class: 'input',
    name: 'currency',
    value: account?.currency ?? state.config?.baseCurrency ?? 'UZS',
    maxLength: 3,
    required: true,
    placeholder: 'UZS',
  });

  simpleForm({
    title: account ? 'Hisobni tahrirlash' : 'Yangi hisob',
    fields: [
      fieldRow('Nom', name, 'Masalan: Naqd, Uzcard, Humo, Bank hisobi'),
      fieldRow('Valyuta', currency, '3 harfli kod: UZS, USD, EUR…'),
      checkbox('showInBalance', 'Balans panelida ko‘rsatilsin', account?.showInBalance ?? true),
      checkbox('active', 'Faol', account?.active ?? true),
    ],
    submit: async (form) => {
      await call('catalog.saveAccount', {
        id: account?.id ?? null,
        name: name.value,
        currency: currency.value,
        showInBalance: checkboxValue(form, 'showInBalance'),
        active: checkboxValue(form, 'active'),
      });
      refresh();
    },
  });
}

function openPointForm(point: Point | null, refresh: () => void): void {
  const name = nameInput(point?.name ?? '');
  const firmSelect = el('select', { class: 'input', name: 'firmId', required: true });
  fillSelect(
    firmSelect,
    state.catalog.firms.map((firm) => ({ value: firm.id, label: firm.name })),
    point?.firmId ?? '',
    'Tanlang…',
  );

  simpleForm({
    title: point ? 'Savdo nuqtasini tahrirlash' : 'Yangi savdo nuqtasi',
    fields: [
      fieldRow('Nom', name),
      fieldRow('Firma', firmSelect),
      checkbox(
        'excludeFromRevenue',
        'Daromad statistikasiga kirmasin',
        point?.excludeFromRevenue ?? false,
      ),
      checkbox('active', 'Faol', point?.active ?? true),
    ],
    submit: async (form) => {
      await call('catalog.savePoint', {
        id: point?.id ?? null,
        name: name.value,
        firmId: firmSelect.value,
        excludeFromRevenue: checkboxValue(form, 'excludeFromRevenue'),
        active: checkboxValue(form, 'active'),
      });
      refresh();
    },
  });
}

function openCategoryForm(category: Category | null, refresh: () => void): void {
  const name = nameInput(category?.name ?? '');
  const kind = el('select', { class: 'input', name: 'kind' });
  fillSelect(
    kind,
    [
      { value: 'expense', label: 'Chiqim' },
      { value: 'income', label: 'Kirim' },
      { value: 'both', label: 'Ikkalasi' },
    ],
    category?.kind ?? 'expense',
  );

  simpleForm({
    title: category ? 'Kategoriyani tahrirlash' : 'Yangi kategoriya',
    fields: [fieldRow('Nom', name), fieldRow('Turi', kind), checkbox('active', 'Faol', category?.active ?? true)],
    submit: async (form) => {
      await call('catalog.saveCategory', {
        id: category?.id ?? null,
        name: name.value,
        kind: kind.value as Category['kind'],
        active: checkboxValue(form, 'active'),
      });
      refresh();
    },
  });
}

// ---------------------------------------------------------------------------
// Xodimlar
// ---------------------------------------------------------------------------

async function renderUsersTab(body: HTMLElement): Promise<void> {
  body.replaceChildren(spinner());
  let users: PublicUser[];
  try {
    users = await call('users.list', {});
  } catch (error) {
    body.replaceChildren(emptyState('Yuklab bo‘lmadi', errorMessage(error)));
    return;
  }

  const refresh = (): void => void renderUsersTab(body);
  const canEdit = allowed('users', 'edit');
  const canDelete = allowed('users', 'delete');

  const table = dataTable<PublicUser>(
    [
      { header: 'Login', render: (user) => user.login },
      { header: 'Ism', render: (user) => user.displayName },
      { header: 'Rol', render: (user) => ROLE_LABELS[user.role] },
      { header: 'Holat', render: (user) => statusBadge(user.status === 'active') },
      {
        header: 'Oxirgi kirish',
        render: (user) => (user.lastLoginAt ? new Date(user.lastLoginAt).toLocaleString('uz-UZ') : '—'),
      },
      {
        header: '',
        align: 'right',
        width: '90px',
        render: (user) => {
          const wrap = el('div', { class: 'row-actions' });
          if (canEdit) {
            wrap.appendChild(
              el('button', {
                class: 'icon-button',
                type: 'button',
                text: '✎',
                title: 'Tahrirlash',
                on: { click: () => openUserForm(user, refresh) },
              }),
            );
          }
          if (canDelete && user.id !== state.user?.id) {
            wrap.appendChild(
              el('button', {
                class: 'icon-button icon-button--danger',
                type: 'button',
                text: '🗑',
                title: "O'chirish",
                on: {
                  click: () => {
                    void (async () => {
                      const confirmed = await confirmDialog({
                        title: 'Xodimni o‘chirish',
                        message: `«${user.login}» o‘chirilsinmi?`,
                        confirmText: "O'chirish",
                        danger: true,
                      });
                      if (!confirmed) return;
                      try {
                        await call('users.delete', { id: user.id });
                        toast('O‘chirildi', 'success');
                        refresh();
                      } catch (error) {
                        toast(errorMessage(error), 'error');
                      }
                    })();
                  },
                },
              }),
            );
          }
          return wrap;
        },
      },
    ],
    users,
    emptyState('Xodimlar yo‘q'),
  );

  body.replaceChildren(
    section(
      'Xodimlar',
      [table],
      canEdit
        ? [
            el('button', {
              class: 'button button--primary button--small',
              type: 'button',
              text: '+ Xodim',
              on: { click: () => openUserForm(null, refresh) },
            }),
          ]
        : [],
    ),
  );
}

function permissionMatrix(permissions: Permissions, role: Role): HTMLElement {
  const table = el('table', { class: 'table table--compact perm-table' });
  const head = el('thead', {}, [
    el('tr', {}, [
      el('th', { text: 'Bo‘lim' }),
      ...ACTIONS.map((action) => el('th', { class: 'align-center', text: ACTION_LABELS[action] })),
    ]),
  ]);

  const bodyRows = el('tbody');
  for (const resource of RESOURCES) {
    const row = el('tr', {}, [el('td', { text: RESOURCE_LABELS[resource] })]);
    for (const action of ACTIONS) {
      const input = el('input', {
        type: 'checkbox',
        name: `perm_${resource}_${action}`,
        checked: permissions[resource]?.[action] === true,
        disabled: role === 'admin',
      });
      row.appendChild(el('td', { class: 'align-center' }, [input]));
    }
    bodyRows.appendChild(row);
  }

  table.append(head, bodyRows);
  return el('div', { class: 'table-wrap' }, [table]);
}

function collectPermissions(form: HTMLElement): Permissions {
  const permissions: Permissions = {};
  for (const resource of RESOURCES as readonly Resource[]) {
    const entry: Partial<Record<Action, boolean>> = {};
    for (const action of ACTIONS as readonly Action[]) {
      const input = form.querySelector<HTMLInputElement>(`[name="perm_${resource}_${action}"]`);
      if (input?.checked) entry[action] = true;
    }
    if (Object.keys(entry).length > 0) permissions[resource] = entry;
  }
  return permissions;
}

function openUserForm(user: PublicUser | null, refresh: () => void): void {
  const login = el('input', {
    class: 'input',
    name: 'login',
    value: user?.login ?? '',
    required: true,
    maxLength: 32,
    autocomplete: 'off',
  });
  const displayName = el('input', {
    class: 'input',
    name: 'displayName',
    value: user?.displayName ?? '',
    maxLength: 60,
  });
  const password = el('input', {
    class: 'input',
    type: 'password',
    name: 'password',
    autocomplete: 'new-password',
    placeholder: user ? 'O‘zgartirmaslik uchun bo‘sh qoldiring' : 'kamida 8 ta belgi',
  });
  const role = el('select', { class: 'input', name: 'role' });
  fillSelect(
    role,
    ROLES.map((value) => ({ value, label: ROLE_LABELS[value] })),
    user?.role ?? 'operator',
  );
  const status = el('select', { class: 'input', name: 'status' });
  fillSelect(
    status,
    [
      { value: 'active', label: 'Faol' },
      { value: 'blocked', label: 'Bloklangan' },
    ],
    user?.status ?? 'active',
  );

  const matrixHost = el('div', {}, [permissionMatrix(user?.permissions ?? {}, user?.role ?? 'operator')]);
  const adminNote = el('p', { class: 'field__hint', text: 'Administrator barcha huquqlarga avtomatik ega.' });
  adminNote.hidden = (user?.role ?? 'operator') !== 'admin';

  role.addEventListener('change', () => {
    const selected = role.value as Role;
    adminNote.hidden = selected !== 'admin';
    matrixHost.replaceChildren(permissionMatrix(collectPermissions(matrixHost), selected));
  });

  simpleForm({
    title: user ? 'Xodimni tahrirlash' : 'Yangi xodim',
    fields: [
      fieldRow('Login', login),
      fieldRow('Ism', displayName),
      fieldRow(
        'Parol',
        password,
        user ? 'Bo‘sh qoldirilsa eski parol saqlanadi' : 'Kamida 8 ta belgi, harf va raqam',
      ),
      fieldRow('Rol', role),
      fieldRow('Holat', status),
      el('div', { class: 'field' }, [
        el('span', { class: 'field__label', text: 'Qo‘shimcha huquqlar' }),
        adminNote,
        matrixHost,
      ]),
    ],
    submit: async (form) => {
      const passwordValue = password.value.trim();
      if (!user && passwordValue === '') {
        throw Object.assign(new Error('Parolni kiriting'), { field: 'password' });
      }
      if (passwordValue !== '') {
        const strength = checkPasswordStrength(passwordValue);
        if (!strength.ok) throw new Error(strength.message ?? 'Parol talabga javob bermaydi');
      }

      await call('users.save', {
        id: user?.id ?? null,
        login: login.value,
        displayName: displayName.value,
        password: passwordValue === '' ? null : passwordValue,
        role: role.value as Role,
        status: status.value as PublicUser['status'],
        permissions: collectPermissions(form),
      });
      refresh();
    },
  });
}

// ---------------------------------------------------------------------------
// Umumiy sozlamalar
// ---------------------------------------------------------------------------

async function renderGeneralTab(body: HTMLElement): Promise<void> {
  body.replaceChildren(spinner());
  let config: AppConfig;
  try {
    config = await call('config.get', {});
  } catch (error) {
    body.replaceChildren(emptyState('Yuklab bo‘lmadi', errorMessage(error)));
    return;
  }

  const canEdit = allowed('config', 'edit');
  const form = el('form', { class: 'form form--grid' });

  const organizationName = el('input', {
    class: 'input',
    name: 'organizationName',
    value: config.organizationName,
    maxLength: 80,
    disabled: !canEdit,
  });
  const baseCurrencyInput = el('input', {
    class: 'input',
    name: 'baseCurrency',
    value: config.baseCurrency,
    maxLength: 3,
    disabled: !canEdit,
  });
  const allocation = el('select', { class: 'input', name: 'sharedExpenseAllocation', disabled: !canEdit });
  fillSelect(
    allocation,
    EXPENSE_ALLOCATIONS.map((value) => ({ value, label: ALLOCATION_LABELS[value] })),
    config.sharedExpenseAllocation,
  );
  const fxProvider = el('select', { class: 'input', name: 'fxProvider', disabled: !canEdit });
  fillSelect(
    fxProvider,
    [
      { value: 'cbu.uz', label: 'Markaziy bank (cbu.uz)' },
      { value: 'none', label: 'Faqat qo‘lda kiritish' },
    ],
    config.fxProvider,
  );
  const sessionTtl = el('input', {
    class: 'input',
    type: 'number',
    name: 'sessionTtlMinutes',
    value: String(config.sessionTtlMinutes),
    min: '5',
    max: '10080',
    disabled: !canEdit,
  });
  const maxAmount = el('input', {
    class: 'input',
    type: 'number',
    name: 'maxTransactionAmount',
    value: String(config.maxTransactionAmount),
    min: '1',
    disabled: !canEdit,
  });
  const maxNote = el('input', {
    class: 'input',
    type: 'number',
    name: 'maxNoteLength',
    value: String(config.maxNoteLength),
    min: '0',
    max: '2000',
    disabled: !canEdit,
  });

  const submit = el('button', {
    class: 'button button--primary',
    type: 'submit',
    text: 'Saqlash',
    disabled: !canEdit,
  });

  form.append(
    fieldRow('Tashkilot nomi', organizationName),
    fieldRow('Hisobot valyutasi', baseCurrencyInput, 'Barcha umumiy ko‘rsatkichlar shu valyutada'),
    fieldRow('Umumiy xarajatlarni taqsimlash', allocation),
    fieldRow('Valyuta kursi manbasi', fxProvider),
    fieldRow('Sessiya muddati (daqiqa)', sessionTtl),
    fieldRow('Maksimal summa', maxAmount),
    fieldRow('Izoh uzunligi', maxNote),
    el('p', { class: 'form-error' }),
    submit,
  );
  (form.querySelector('.form-error') as HTMLElement).hidden = true;

  form.addEventListener('submit', (event) => {
    event.preventDefault();
    void withBusy(submit, async () => {
      try {
        const saved = await call('config.save', {
          organizationName: organizationName.value,
          baseCurrency: baseCurrencyInput.value,
          sharedExpenseAllocation: allocation.value as ExpenseAllocation,
          fxProvider: fxProvider.value as AppConfig['fxProvider'],
          sessionTtlMinutes: Number(sessionTtl.value),
          maxTransactionAmount: Number(maxAmount.value),
          maxNoteLength: Number(maxNote.value),
        });
        state.config = saved;
        toast('Sozlamalar saqlandi', 'success');
        await onCatalogChanged?.();
      } catch (error) {
        setFieldError(form, errorField(error), errorMessage(error));
      }
    });
  });

  body.replaceChildren(section('Umumiy sozlamalar', [form]), renderFxSection(canEdit), renderPasswordSection());
}

function renderFxSection(canEdit: boolean): HTMLElement {
  const dateInput = el('input', { class: 'input', type: 'date', value: state.today, disabled: !canEdit });
  const currencyInput = el('input', {
    class: 'input',
    value: 'USD',
    maxLength: 3,
    disabled: !canEdit,
  });
  const rateInput = el('input', {
    class: 'input',
    inputMode: 'decimal',
    placeholder: '12500',
    disabled: !canEdit,
  });
  const result = el('p', { class: 'field__hint' });

  const checkButton = el('button', {
    class: 'button button--ghost',
    type: 'button',
    text: 'Kursni tekshirish',
    on: {
      click: () => {
        void (async () => {
          try {
            const rate = await call('fx.rate', {
              date: dateInput.value,
              currency: currencyInput.value.toUpperCase(),
            });
            result.textContent = `${rate.currency}: ${rate.rate} (${rate.effectiveDate}, manba: ${rate.source})`;
            rateInput.value = String(rate.rate);
          } catch (error) {
            result.textContent = errorMessage(error);
          }
        })();
      },
    },
  });

  const saveButton = el('button', {
    class: 'button button--primary',
    type: 'button',
    text: 'Qo‘lda saqlash',
    disabled: !canEdit,
    on: {
      click: () => {
        void (async () => {
          try {
            const saved = await call('fx.setManual', {
              date: dateInput.value,
              currency: currencyInput.value.toUpperCase(),
              rate: Number(rateInput.value.replace(',', '.')),
            });
            result.textContent = `Saqlandi: ${saved.currency} = ${saved.rate}`;
            toast('Kurs saqlandi', 'success');
          } catch (error) {
            toast(errorMessage(error), 'error');
          }
        })();
      },
    },
  });

  return section('Valyuta kursi', [
    el('div', { class: 'form form--grid' }, [
      fieldRow('Sana', dateInput),
      fieldRow('Valyuta', currencyInput),
      fieldRow('Kurs', rateInput, 'Internet ishlamasa kursni qo‘lda kiriting'),
    ]),
    el('div', { class: 'row-actions' }, [checkButton, saveButton]),
    result,
  ]);
}

function renderPasswordSection(): HTMLElement {
  const current = el('input', { class: 'input', type: 'password', name: 'currentPassword', autocomplete: 'current-password' });
  const next = el('input', { class: 'input', type: 'password', name: 'newPassword', autocomplete: 'new-password' });
  const repeat = el('input', { class: 'input', type: 'password', name: 'repeatPassword', autocomplete: 'new-password' });

  const form = el('form', { class: 'form form--grid' }, [
    fieldRow('Joriy parol', current),
    fieldRow('Yangi parol', next),
    fieldRow('Takrorlang', repeat),
    el('p', { class: 'form-error' }),
  ]);
  (form.querySelector('.form-error') as HTMLElement).hidden = true;

  const submit = el('button', { class: 'button button--primary', type: 'submit', text: 'Parolni o‘zgartirish' });
  form.appendChild(submit);

  form.addEventListener('submit', (event) => {
    event.preventDefault();
    if (next.value !== repeat.value) {
      setFieldError(form, 'repeatPassword', 'Parollar mos kelmadi');
      return;
    }
    void withBusy(submit, async () => {
      try {
        await call('auth.changePassword', {
          currentPassword: current.value,
          newPassword: next.value,
        });
        toast('Parol o‘zgartirildi. Qaytadan kiring.', 'success', 5000);
      } catch (error) {
        setFieldError(form, errorField(error), errorMessage(error));
      }
    });
  });

  return section('Parolni o‘zgartirish', [form]);
}

// ---------------------------------------------------------------------------
// Amallar tarixi
// ---------------------------------------------------------------------------

async function renderAuditTab(body: HTMLElement): Promise<void> {
  body.replaceChildren(spinner());
  try {
    const page = await call('audit.list', { offset: 0, limit: 100 });
    const table = dataTable<AuditEntry>(
      [
        { header: 'Vaqt', render: (entry) => new Date(entry.at).toLocaleString('uz-UZ') },
        { header: 'Xodim', render: (entry) => entry.userLogin },
        { header: 'Amal', render: (entry) => entry.action },
        { header: 'Obyekt', render: (entry) => entry.entity },
        { header: 'Tafsilot', render: (entry) => entry.details || '—' },
      ],
      page.items,
      emptyState('Tarix bo‘sh'),
    );
    body.replaceChildren(section(`Amallar tarixi (jami ${page.total})`, [table]));
  } catch (error) {
    body.replaceChildren(emptyState('Yuklab bo‘lmadi', errorMessage(error)));
  }
}
