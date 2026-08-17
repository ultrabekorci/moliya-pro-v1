/**
 * RPC yo'naltiruvchisi — serverga kiradigan YAGONA eshik.
 *
 * Har bir metod uchun kerakli ruxsat shu yerda e'lon qilingan. Ruxsat tekshiruvi
 * markazlashgan: yangi metod qo'shilganda uni ro'yxatga kiritish shart, aks holda
 * u umuman chaqirilmaydi (fail-closed).
 */

import { AppError } from '../shared/api.js';
import type {
  ApiMethod,
  ApiParams,
  ApiResult,
  RpcRequest,
  RpcResponse,
  SessionInfo,
} from '../shared/api.js';
import type { Action, Resource, TransactionKind } from '../shared/types.js';
import { TRANSACTION_KINDS } from '../shared/types.js';
import { can, effectivePermissions } from '../shared/permissions.js';
import { todayIso } from '../shared/dates.js';
import { normalizeLogin } from '../shared/validation.js';
import { ensureSchema } from './db.js';
import type { StoredUser } from './db.js';
import { scriptTimeZone } from './sheets.js';
import * as auth from './services/auth.js';
import * as audit from './services/audit.js';
import * as catalog from './services/catalog.js';
import * as configService from './services/config.js';
import * as fx from './services/fx.js';
import * as transactions from './services/transactions.js';
import * as users from './services/users.js';

type Guard =
  | { kind: 'public' }
  | { kind: 'authenticated' }
  | { kind: 'permission'; resource: Resource; action: Action };

const PUBLIC: Guard = { kind: 'public' };
const AUTH: Guard = { kind: 'authenticated' };

function perm(resource: Resource, action: Action): Guard {
  return { kind: 'permission', resource, action };
}

/** Metod → talab qilinadigan ruxsat. */
const GUARDS: Record<ApiMethod, Guard> = {
  'app.status': PUBLIC,
  'auth.setup': PUBLIC,
  'auth.login': PUBLIC,
  'auth.logout': AUTH,
  'auth.session': AUTH,
  'auth.changePassword': AUTH,

  'app.bootstrap': AUTH,

  // Ro'yxat uchun ruxsat tur bo'yicha `handle` ichida cheklanadi: foydalanuvchi
  // faqat o'zi ko'rishga haqli turlarni oladi.
  'tx.list': AUTH,
  'tx.listGrouped': AUTH,
  'tx.save': AUTH, // Aniq tur bo'yicha tekshiruv `handle` ichida.
  'tx.delete': AUTH,
  'tx.balances': AUTH,
  'tx.dashboard': perm('dashboard', 'view'),

  'catalog.get': AUTH,
  'catalog.saveFirm': perm('catalog', 'edit'),
  'catalog.saveAccount': perm('catalog', 'edit'),
  'catalog.savePoint': perm('catalog', 'edit'),
  'catalog.saveCategory': perm('catalog', 'edit'),
  'catalog.deleteFirm': perm('catalog', 'delete'),
  'catalog.deleteAccount': perm('catalog', 'delete'),
  'catalog.deletePoint': perm('catalog', 'delete'),
  'catalog.deleteCategory': perm('catalog', 'delete'),

  'users.list': perm('users', 'view'),
  'users.save': perm('users', 'edit'),
  'users.delete': perm('users', 'delete'),

  'config.get': AUTH,
  'config.save': perm('config', 'edit'),

  'fx.rate': AUTH,
  'fx.setManual': perm('config', 'edit'),

  'audit.list': perm('audit', 'view'),
};

function isApiMethod(value: unknown): value is ApiMethod {
  return typeof value === 'string' && Object.prototype.hasOwnProperty.call(GUARDS, value);
}

function sessionInfo(user: StoredUser, token: string, expiresAt: Date): SessionInfo {
  return {
    token,
    expiresAt: expiresAt.toISOString(),
    user: auth.toPublicUser(user),
    permissions: effectivePermissions(user.role, user.permissions),
  };
}

function requirePermission(user: StoredUser, resource: Resource, action: Action): void {
  const permissions = effectivePermissions(user.role, user.permissions);
  if (!can(permissions, resource, action)) {
    throw new AppError('FORBIDDEN', 'Bu amal uchun ruxsatingiz yo‘q');
  }
}

/**
 * So'rovni foydalanuvchi ko'ra oladigan turlar bilan cheklaydi. Klient qanday
 * filtr yuborishidan qat'i nazar, ruxsatsiz tur natijaga tushmaydi.
 */
function scopeToVisibleKinds<T extends { filter?: { kinds?: TransactionKind[] } }>(
  query: T,
  user: StoredUser,
): T {
  const permissions = effectivePermissions(user.role, user.permissions);
  const visible = TRANSACTION_KINDS.filter((kind) => can(permissions, kind, 'view'));
  if (visible.length === 0) throw new AppError('FORBIDDEN', 'Yozuvlarni ko‘rish huquqingiz yo‘q');

  const requested = query.filter?.kinds;
  const kinds = requested && requested.length > 0 ? requested.filter((kind) => visible.includes(kind)) : visible;
  if (kinds.length === 0) throw new AppError('FORBIDDEN', 'Bu turdagi yozuvlarni ko‘rish huquqingiz yo‘q');

  return { ...query, filter: { ...(query.filter ?? {}), kinds } };
}

/** `tx.save` / `tx.delete` uchun aniq tranzaksiya turi bo'yicha tekshiruv. */
function requireTransactionPermission(user: StoredUser, kind: unknown, action: Action): void {
  const resource: Resource =
    kind === 'expense' ? 'expense' : kind === 'transfer' ? 'transfer' : 'income';
  requirePermission(user, resource, action);
}

function handle<M extends ApiMethod>(
  method: M,
  params: ApiParams<M>,
  user: StoredUser | null,
  token: string | null,
): ApiResult<M> {
  // `as never` — TypeScript uzatilayotgan aniq metodni bu yerda tor qila olmaydi,
  // lekin har bir shox o'z shartnomasiga mos qiymat qaytaradi.
  switch (method) {
    case 'app.status': {
      const config = configService.getConfig();
      return {
        configured: auth.hasAnyUser(),
        organizationName: config.organizationName,
      } as ApiResult<M>;
    }

    case 'auth.setup': {
      const input = params as ApiParams<'auth.setup'>;
      const created = users.createFirstAdmin(input.login, input.password);
      const stored = auth.findUserByLogin(created.login);
      if (!stored) throw new AppError('INTERNAL', 'Administrator yaratilmadi');
      catalog.seedDefaults(configService.getConfig().baseCurrency);
      const session = auth.createSession(stored.id);
      return sessionInfo(stored, session.token, session.expiresAt) as ApiResult<M>;
    }

    case 'auth.login': {
      const input = params as ApiParams<'auth.login'>;
      const login = normalizeLogin(input.login);
      auth.assertNotLockedOut(login);

      const found = auth.findUserByLogin(login);
      // Foydalanuvchi topilmasa ham parol tekshiruviga teng vaqt sarflanadi.
      const ok = found ? auth.verifyPassword(found, String(input.password ?? '')) : false;

      if (!found || !ok) {
        auth.registerFailedLogin(login);
        throw new AppError('UNAUTHENTICATED', 'Login yoki parol noto‘g‘ri');
      }
      if (found.status !== 'active') {
        throw new AppError('FORBIDDEN', 'Profilingiz bloklangan. Administrator bilan bog‘laning.');
      }

      auth.clearFailedLogins(login);
      users.touchLastLogin(found.id);
      const session = auth.createSession(found.id);
      audit.record(found, 'auth.login', 'user', found.id, null);
      return sessionInfo(found, session.token, session.expiresAt) as ApiResult<M>;
    }

    case 'auth.logout': {
      auth.destroySession(token);
      return { ok: true } as ApiResult<M>;
    }

    case 'auth.session': {
      const current = user!;
      const session = auth.createSession(current.id);
      auth.destroySession(token);
      return sessionInfo(current, session.token, session.expiresAt) as ApiResult<M>;
    }

    case 'auth.changePassword': {
      users.changeOwnPassword(params as ApiParams<'auth.changePassword'>, user!);
      return { ok: true } as ApiResult<M>;
    }

    case 'app.bootstrap': {
      const config = configService.getConfig();
      const timeZone = scriptTimeZone();
      const today = todayIso(timeZone);
      const cat = catalog.getCatalog();
      return {
        config,
        catalog: cat,
        balances: transactions.balances({}),
        rates: fx.currentRates(
          cat.accounts.map((account) => account.currency),
          today,
        ),
        today,
        serverTimeZone: timeZone,
      } as ApiResult<M>;
    }

    case 'tx.list':
      return transactions.list(scopeToVisibleKinds(params as ApiParams<'tx.list'>, user!)) as ApiResult<M>;

    case 'tx.listGrouped':
      return transactions.listGrouped(
        scopeToVisibleKinds(params as ApiParams<'tx.listGrouped'>, user!),
      ) as ApiResult<M>;

    case 'tx.save': {
      const input = params as ApiParams<'tx.save'>;
      requireTransactionPermission(user!, input.kind, input.batchId ? 'edit' : 'create');
      return transactions.save(input, user!) as ApiResult<M>;
    }

    case 'tx.delete': {
      const input = params as ApiParams<'tx.delete'>;
      const kind = transactions.batchKind(input.id);
      if (!kind) throw new AppError('NOT_FOUND', 'Yozuv topilmadi');
      requireTransactionPermission(user!, kind, 'delete');
      transactions.remove(input.id, user!);
      return { ok: true } as ApiResult<M>;
    }

    case 'tx.balances':
      return transactions.balances(params as ApiParams<'tx.balances'>) as ApiResult<M>;

    case 'tx.dashboard':
      return transactions.dashboard(params as ApiParams<'tx.dashboard'>) as ApiResult<M>;

    case 'catalog.get':
      return catalog.getCatalog() as ApiResult<M>;

    case 'catalog.saveFirm': {
      const result = catalog.saveFirm(params as ApiParams<'catalog.saveFirm'>);
      audit.record(user!, 'catalog.saveFirm', 'firm', result.id, params);
      return result as ApiResult<M>;
    }

    case 'catalog.saveAccount': {
      const result = catalog.saveAccount(params as ApiParams<'catalog.saveAccount'>);
      audit.record(user!, 'catalog.saveAccount', 'account', result.id, params);
      return result as ApiResult<M>;
    }

    case 'catalog.savePoint': {
      const result = catalog.savePoint(params as ApiParams<'catalog.savePoint'>);
      audit.record(user!, 'catalog.savePoint', 'point', result.id, params);
      return result as ApiResult<M>;
    }

    case 'catalog.saveCategory': {
      const result = catalog.saveCategory(params as ApiParams<'catalog.saveCategory'>);
      audit.record(user!, 'catalog.saveCategory', 'category', result.id, params);
      return result as ApiResult<M>;
    }

    case 'catalog.deleteFirm': {
      const input = params as ApiParams<'catalog.deleteFirm'>;
      catalog.deleteFirm(input.id);
      audit.record(user!, 'catalog.deleteFirm', 'firm', input.id, null);
      return { ok: true } as ApiResult<M>;
    }

    case 'catalog.deleteAccount': {
      const input = params as ApiParams<'catalog.deleteAccount'>;
      catalog.deleteAccount(input.id);
      audit.record(user!, 'catalog.deleteAccount', 'account', input.id, null);
      return { ok: true } as ApiResult<M>;
    }

    case 'catalog.deletePoint': {
      const input = params as ApiParams<'catalog.deletePoint'>;
      catalog.deletePoint(input.id);
      audit.record(user!, 'catalog.deletePoint', 'point', input.id, null);
      return { ok: true } as ApiResult<M>;
    }

    case 'catalog.deleteCategory': {
      const input = params as ApiParams<'catalog.deleteCategory'>;
      catalog.deleteCategory(input.id);
      audit.record(user!, 'catalog.deleteCategory', 'category', input.id, null);
      return { ok: true } as ApiResult<M>;
    }

    case 'users.list':
      return users.list() as ApiResult<M>;

    case 'users.save':
      return users.save(params as ApiParams<'users.save'>, user!) as ApiResult<M>;

    case 'users.delete': {
      const input = params as ApiParams<'users.delete'>;
      users.remove(input.id, user!);
      return { ok: true } as ApiResult<M>;
    }

    case 'config.get':
      return configService.getConfig() as ApiResult<M>;

    case 'config.save': {
      const saved = configService.saveConfig(params as ApiParams<'config.save'>);
      audit.record(user!, 'config.save', 'config', 'app', params);
      return saved as ApiResult<M>;
    }

    case 'fx.rate': {
      const input = params as ApiParams<'fx.rate'>;
      return fx.getRate(input.date, input.currency) as ApiResult<M>;
    }

    case 'fx.setManual': {
      const input = params as ApiParams<'fx.setManual'>;
      const result = fx.setManualRate(input.date, input.currency, input.rate);
      audit.record(user!, 'fx.setManual', 'fx', `${input.currency}:${input.date}`, input);
      return result as ApiResult<M>;
    }

    case 'audit.list': {
      const input = params as ApiParams<'audit.list'>;
      return audit.list(input.offset, input.limit) as ApiResult<M>;
    }

    default: {
      throw new AppError('NOT_FOUND', `Noma'lum metod: ${String(method)}`);
    }
  }
}

/** Kirish nuqtasi. Hech qanday istisno tashqariga chiqmaydi — hammasi `RpcResponse` ga o'raladi. */
export function dispatch(request: unknown): RpcResponse {
  try {
    if (typeof request !== 'object' || request === null) {
      throw new AppError('VALIDATION', 'So‘rov formati noto‘g‘ri');
    }
    const { method, params, token } = request as RpcRequest;
    if (!isApiMethod(method)) throw new AppError('NOT_FOUND', 'Noma’lum metod');

    ensureSchema();

    const guard = GUARDS[method];
    let user: StoredUser | null = null;

    if (guard.kind !== 'public') {
      user = auth.resolveSession(token);
      if (guard.kind === 'permission') {
        requirePermission(user, guard.resource, guard.action);
      }
    }

    const data = handle(method, (params ?? {}) as ApiParams<typeof method>, user, token ?? null);
    return { ok: true, data };
  } catch (error) {
    if (error instanceof AppError) return { ok: false, error: error.toRpcError() };
    const message = error instanceof Error ? error.message : String(error);
    console.error(`RPC error: ${message}`);
    // Ichki xatolar tafsiloti klientga chiqmaydi.
    return {
      ok: false,
      error: { code: 'INTERNAL', message: 'Serverda kutilmagan xatolik yuz berdi' },
    };
  }
}
