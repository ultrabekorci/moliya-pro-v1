/**
 * Kirish ekrani va birinchi ishga tushirish (birinchi administratorni yaratish).
 */

import { call, errorField, errorMessage } from '../api.js';
import { el } from '../dom.js';
import { fieldRow, setFieldError, toast, withBusy } from '../ui.js';
import { applySession } from '../store.js';
import { checkPasswordStrength } from '../../shared/validation.js';

export interface LoginViewOptions {
  onSuccess(): void | Promise<void>;
}

export function renderLogin(container: HTMLElement, options: LoginViewOptions): void {
  container.replaceChildren(el('div', { class: 'auth' }, [el('div', { class: 'auth__card' }, [])]));
  const card = container.querySelector<HTMLElement>('.auth__card');
  if (!card) return;

  void bootstrapStatus(card, options);
}

async function bootstrapStatus(card: HTMLElement, options: LoginViewOptions): Promise<void> {
  card.replaceChildren(el('p', { class: 'auth__loading', text: 'Tekshirilmoqda…' }));
  try {
    const status = await call('app.status', {});
    if (status.configured) renderLoginForm(card, status.organizationName, options);
    else renderSetupForm(card, status.organizationName, options);
  } catch (error) {
    card.replaceChildren(
      el('p', { class: 'auth__title', text: 'Ulanib bo‘lmadi' }),
      el('p', { class: 'form-error', text: errorMessage(error) }),
      el('button', {
        class: 'button button--primary',
        type: 'button',
        text: 'Qayta urinish',
        on: { click: () => void bootstrapStatus(card, options) },
      }),
    );
  }
}

function passwordField(name: string, placeholder: string): HTMLElement {
  const input = el('input', {
    class: 'input',
    type: 'password',
    name,
    placeholder,
    required: true,
    autocomplete: name === 'password' ? 'current-password' : 'new-password',
  });

  const toggle = el('button', {
    class: 'input-affix',
    type: 'button',
    text: '👁',
    attrs: { 'aria-label': 'Parolni ko‘rsatish', tabindex: '-1' },
    on: {
      click: () => {
        input.type = input.type === 'password' ? 'text' : 'password';
      },
    },
  });

  return el('div', { class: 'input-group' }, [input, toggle]);
}

function renderLoginForm(card: HTMLElement, organizationName: string, options: LoginViewOptions): void {
  const form = el('form', { class: 'auth__form' });
  const submit = el('button', { class: 'button button--primary button--block', type: 'submit', text: 'Kirish' });

  const loginInput = el('input', {
    class: 'input',
    name: 'login',
    placeholder: 'login',
    required: true,
    autocomplete: 'username',
    maxLength: 32,
  });

  form.append(
    fieldRow('Login', loginInput),
    fieldRow('Parol', passwordField('password', '••••••••')),
    el('p', { class: 'form-error' }),
    submit,
  );
  (form.querySelector('.form-error') as HTMLElement).hidden = true;

  form.addEventListener('submit', (event) => {
    event.preventDefault();
    const password = form.querySelector<HTMLInputElement>('[name="password"]')?.value ?? '';
    void withBusy(submit, async () => {
      try {
        const session = await call('auth.login', { login: loginInput.value, password });
        applySession(session);
        await options.onSuccess();
      } catch (error) {
        setFieldError(form, errorField(error), errorMessage(error));
      }
    });
  });

  card.replaceChildren(
    el('h1', { class: 'auth__title', text: organizationName || 'Moliya-Pro' }),
    el('p', { class: 'auth__subtitle', text: 'Hisobingizga kiring' }),
    form,
  );
  loginInput.focus();
}

function renderSetupForm(card: HTMLElement, organizationName: string, options: LoginViewOptions): void {
  const form = el('form', { class: 'auth__form' });
  const submit = el('button', {
    class: 'button button--primary button--block',
    type: 'submit',
    text: 'Administrator yaratish',
  });

  const loginInput = el('input', {
    class: 'input',
    name: 'login',
    placeholder: 'admin',
    required: true,
    autocomplete: 'username',
    maxLength: 32,
  });
  const passwordGroup = passwordField('newPassword', 'kamida 8 ta belgi');
  const repeatGroup = passwordField('repeatPassword', 'parolni takrorlang');

  form.append(
    fieldRow('Login', loginInput, 'Faqat harf, raqam va . _ - belgilari'),
    fieldRow('Parol', passwordGroup, 'Kamida 8 ta belgi, harf va raqam bo‘lsin'),
    fieldRow('Parolni takrorlang', repeatGroup),
    el('p', { class: 'form-error' }),
    submit,
  );
  (form.querySelector('.form-error') as HTMLElement).hidden = true;

  form.addEventListener('submit', (event) => {
    event.preventDefault();
    const password = form.querySelector<HTMLInputElement>('[name="newPassword"]')?.value ?? '';
    const repeat = form.querySelector<HTMLInputElement>('[name="repeatPassword"]')?.value ?? '';

    if (password !== repeat) {
      setFieldError(form, 'repeatPassword', 'Parollar mos kelmadi');
      return;
    }
    const strength = checkPasswordStrength(password);
    if (!strength.ok) {
      setFieldError(form, 'newPassword', strength.message ?? 'Parol talabga javob bermaydi');
      return;
    }

    void withBusy(submit, async () => {
      try {
        const session = await call('auth.setup', { login: loginInput.value, password });
        applySession(session);
        toast('Tizim sozlandi. Xush kelibsiz!', 'success');
        await options.onSuccess();
      } catch (error) {
        setFieldError(form, errorField(error), errorMessage(error));
      }
    });
  });

  card.replaceChildren(
    el('h1', { class: 'auth__title', text: organizationName || 'Moliya-Pro' }),
    el('p', { class: 'auth__subtitle', text: 'Birinchi ishga tushirish — administrator hisobini yarating' }),
    el('p', {
      class: 'auth__note',
      text: 'Bu oyna faqat bir marta, tizimda hech qanday foydalanuvchi bo‘lmaganda ko‘rinadi.',
    }),
    form,
  );
  loginInput.focus();
}
