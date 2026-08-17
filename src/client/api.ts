/**
 * Tipizatsiyalangan RPC klienti.
 *
 * v1 da har bir chaqiruvda `withFailureHandler` yozish esdan chiqardi va server
 * xatosi foydalanuvchiga umuman ko'rinmasdi (jimgina "Muvaffaqiyatli!" chiqardi).
 * Bu yerda xato yo'li YAGONA: `call()` Promise qaytaradi va xato bo'lsa
 * `AppError` bilan reject bo'ladi — uni e'tiborsiz qoldirib bo'lmaydi.
 */

import { AppError } from '../shared/api.js';
import type { ApiMethod, ApiParams, ApiResult, RpcResponse } from '../shared/api.js';

/** `google.script.run` ning bizga kerakli qismi. */
interface ScriptRunner {
  withSuccessHandler(handler: (value: string) => void): ScriptRunner;
  withFailureHandler(handler: (error: Error) => void): ScriptRunner;
  rpc(payload: string): void;
}

declare global {
  interface Window {
    google?: {
      script?: {
        run: ScriptRunner;
      };
    };
  }
}

let sessionToken: string | null = null;

export function setToken(token: string | null): void {
  sessionToken = token;
}

export function getToken(): string | null {
  return sessionToken;
}

/** Sessiya tugaganda ilova login ekraniga qaytishi uchun. */
type UnauthorizedHandler = (error: AppError) => void;
let onUnauthorized: UnauthorizedHandler | null = null;

export function setUnauthorizedHandler(handler: UnauthorizedHandler): void {
  onUnauthorized = handler;
}

function transport(payload: string): Promise<string> {
  const runner = window.google?.script?.run;
  if (!runner) {
    return Promise.reject(
      new AppError('INTERNAL', 'Google Apps Script muhiti topilmadi. Sahifani veb-ilova havolasi orqali oching.'),
    );
  }
  return new Promise<string>((resolve, reject) => {
    runner
      .withSuccessHandler(resolve)
      .withFailureHandler((error: Error) => {
        reject(new AppError('INTERNAL', error?.message || 'Server bilan aloqa uzildi'));
      })
      .rpc(payload);
  });
}

export async function call<M extends ApiMethod>(method: M, params: ApiParams<M>): Promise<ApiResult<M>> {
  const payload = JSON.stringify({ method, params, token: sessionToken });
  const raw = await transport(payload);

  let response: RpcResponse<M>;
  try {
    response = JSON.parse(raw) as RpcResponse<M>;
  } catch {
    throw new AppError('INTERNAL', 'Serverdan tushunarsiz javob keldi');
  }

  if (response.ok) return response.data;

  const error = new AppError(response.error.code, response.error.message, response.error.field);
  if (error.code === 'UNAUTHENTICATED' || (error.code === 'FORBIDDEN' && !sessionToken)) {
    onUnauthorized?.(error);
  }
  throw error;
}

/** Xato xabarini foydalanuvchiga ko'rsatish uchun matnga aylantiradi. */
export function errorMessage(error: unknown): string {
  if (error instanceof AppError) return error.message;
  if (error instanceof Error) return error.message;
  return 'Kutilmagan xatolik';
}

export function errorField(error: unknown): string | null {
  return error instanceof AppError && error.field ? error.field : null;
}
