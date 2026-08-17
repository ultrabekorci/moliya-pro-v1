/**
 * Ilova qobig'i: navigatsiya, mavzu (yorug'/qorong'i), sessiya nazorati.
 */

import { call, errorMessage, setToken, setUnauthorizedHandler } from './api.js';
import { el, throttle } from './dom.js';
import { emptyState, toast } from './ui.js';
import {
  allowed,
  applyBootstrap,
  applySession,
  clearSession,
  loadStoredToken,
  state,
} from './store.js';
import { renderLogin } from './views/login.js';
import { renderDashboard } from './views/dashboard.js';
import {
  refreshBalances,
  renderTransactions,
  setTransactionsChangeHandler,
} from './views/transactions.js';
import { renderSettings, setSettingsChangeHandler } from './views/settings.js';

type Route = 'dashboard' | 'transactions' | 'settings';

const ROUTE_LABELS: Record<Route, string> = {
  dashboard: 'Dashboard',
  transactions: 'Yozuvlar',
  settings: 'Sozlamalar',
};

let currentRoute: Route = 'transactions';
let sessionTimer: number | undefined;

const root = (): HTMLElement => {
  const node = document.getElementById('app');
  if (!node) throw new Error('#app topilmadi');
  return node;
};

// ---------------------------------------------------------------------------
// Mavzu
// ---------------------------------------------------------------------------

const THEME_KEY = 'moliya.theme';

function applyTheme(theme: 'light' | 'dark'): void {
  document.documentElement.dataset.theme = theme;
  try {
    window.localStorage.setItem(THEME_KEY, theme);
  } catch {
    // localStorage bloklangan bo'lishi mumkin.
  }
}

function initTheme(): void {
  let stored: string | null = null;
  try {
    stored = window.localStorage.getItem(THEME_KEY);
  } catch {
    stored = null;
  }
  if (stored === 'light' || stored === 'dark') {
    document.documentElement.dataset.theme = stored;
  }
}

function toggleTheme(): void {
  const current =
    document.documentElement.dataset.theme ??
    (window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
  applyTheme(current === 'dark' ? 'light' : 'dark');
}

// ---------------------------------------------------------------------------
// Sessiya nazorati
// ---------------------------------------------------------------------------

function scheduleSessionCheck(): void {
  if (sessionTimer !== undefined) window.clearInterval(sessionTimer);
  sessionTimer = window.setInterval(() => {
    if (!state.user) return;
    if (state.expiresAt > 0 && Date.now() > state.expiresAt) {
      toast('Sessiya muddati tugadi. Qaytadan kiring.', 'warning', 5000);
      void logout();
    }
  }, 30_000);
}

/**
 * Faollik bo'lganda sessiyani uzaytiradi. v1 da bu har bir `mousemove` da
 * `sessionStorage` ga yozardi — bu yerda 5 daqiqada bir martadan tez emas.
 */
const renewSession = throttle(() => {
  if (!state.user) return;
  const remaining = state.expiresAt - Date.now();
  const ttl = (state.config?.sessionTtlMinutes ?? 480) * 60_000;
  if (remaining > ttl / 2) return;

  void call('auth.session', {})
    .then((session) => applySession(session))
    .catch(() => {
      /* Sessiya yangilanmasa, keyingi so'rov o'zi login ekraniga qaytaradi. */
    });
}, 5 * 60_000);

async function logout(): Promise<void> {
  try {
    await call('auth.logout', {});
  } catch {
    // Server javob bermasa ham lokal sessiyani tozalaymiz.
  }
  clearSession();
  showLogin();
}

// ---------------------------------------------------------------------------
// Ko'rinishlar
// ---------------------------------------------------------------------------

function showLogin(): void {
  if (sessionTimer !== undefined) window.clearInterval(sessionTimer);
  renderLogin(root(), { onSuccess: () => startApp() });
}

function availableRoutes(): Route[] {
  const routes: Route[] = [];
  if (allowed('dashboard', 'view')) routes.push('dashboard');
  if (allowed('income', 'view') || allowed('expense', 'view') || allowed('transfer', 'view')) {
    routes.push('transactions');
  }
  if (
    allowed('catalog', 'view') ||
    allowed('users', 'view') ||
    allowed('config', 'view') ||
    allowed('audit', 'view')
  ) {
    routes.push('settings');
  }
  return routes;
}

function renderShell(): void {
  const routes = availableRoutes();
  if (routes.length === 0) {
    root().replaceChildren(
      emptyState('Ruxsat yo‘q', 'Hisobingizga hech qanday bo‘lim biriktirilmagan. Administratorga murojaat qiling.'),
    );
    return;
  }
  if (!routes.includes(currentRoute)) currentRoute = routes[0]!;

  const content = el('main', { class: 'content', id: 'view-content' });

  const nav = el(
    'nav',
    { class: 'nav' },
    routes.map((route) =>
      el('button', {
        class: `nav__item${currentRoute === route ? ' nav__item--active' : ''}`,
        type: 'button',
        text: ROUTE_LABELS[route],
        on: {
          click: () => {
            currentRoute = route;
            renderShell();
          },
        },
      }),
    ),
  );

  const header = el('header', { class: 'header' }, [
    el('div', { class: 'header__brand' }, [
      el('span', { class: 'header__logo', text: '₴' }),
      el('span', { class: 'header__title', text: state.config?.organizationName ?? 'Moliya-Pro' }),
    ]),
    nav,
    el('div', { class: 'header__right' }, [
      el('button', {
        class: 'icon-button',
        type: 'button',
        text: '◐',
        title: 'Mavzuni almashtirish',
        on: { click: toggleTheme },
      }),
      el('span', { class: 'header__user', text: state.user?.displayName ?? '' }),
      el('button', {
        class: 'button button--ghost button--small',
        type: 'button',
        text: 'Chiqish',
        on: { click: () => void logout() },
      }),
    ]),
  ]);

  root().replaceChildren(header, content);

  switch (currentRoute) {
    case 'dashboard':
      renderDashboard(content);
      break;
    case 'transactions':
      renderTransactions(content);
      break;
    case 'settings':
      renderSettings(content);
      break;
    default:
      break;
  }
}

async function startApp(): Promise<void> {
  try {
    const bootstrap = await call('app.bootstrap', {});
    applyBootstrap(bootstrap);
    scheduleSessionCheck();
    renderShell();
  } catch (error) {
    toast(errorMessage(error), 'error', 6000);
    clearSession();
    showLogin();
  }
}

async function restoreSession(): Promise<void> {
  const token = loadStoredToken();
  if (!token) {
    showLogin();
    return;
  }

  setToken(token);
  try {
    // Rol va ruxsatlar HAR DOIM serverdan olinadi — brauzerdagi qiymatga ishonilmaydi.
    const session = await call('auth.session', {});
    applySession(session);
    await startApp();
  } catch {
    clearSession();
    showLogin();
  }
}

function boot(): void {
  initTheme();

  setUnauthorizedHandler(() => {
    clearSession();
    showLogin();
  });

  setTransactionsChangeHandler(() => {
    void refreshBalances();
  });

  setSettingsChangeHandler(async () => {
    const bootstrap = await call('app.bootstrap', {});
    applyBootstrap(bootstrap);
  });

  for (const event of ['click', 'keydown', 'touchstart'] as const) {
    document.addEventListener(event, renewSession, { passive: true });
  }

  void restoreSession();
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', boot);
} else {
  boot();
}
