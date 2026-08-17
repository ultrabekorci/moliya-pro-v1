/**
 * Xavfsiz DOM qurish.
 *
 * v1 da jadvallar `innerHTML += \`<td>${item.note}</td>\`` tarzida qurilardi —
 * izohga `<img onerror=...>` yozgan xodim boshqalarning brauzerida kod ishga
 * tushira olardi. Bu yerdagi yordamchilar matnni FAQAT `textContent` orqali
 * qo'yadi, shuning uchun XSS uchun joy qolmaydi.
 */

export type Child = Node | string | number | null | undefined | false;

export interface ElementOptions {
  class?: string;
  text?: string | number;
  title?: string;
  id?: string;
  type?: string;
  value?: string;
  placeholder?: string;
  name?: string;
  href?: string;
  disabled?: boolean;
  checked?: boolean;
  required?: boolean;
  min?: string;
  max?: string;
  step?: string;
  maxLength?: number;
  autocomplete?: string;
  inputMode?: string;
  rows?: number;
  colSpan?: number;
  dataset?: Record<string, string>;
  attrs?: Record<string, string>;
  style?: Partial<CSSStyleDeclaration>;
  on?: Partial<{
    [K in keyof HTMLElementEventMap]: (event: HTMLElementEventMap[K]) => void;
  }>;
}

export function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  options: ElementOptions = {},
  children: Child[] = [],
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);

  if (options.class) node.className = options.class;
  if (options.id) node.id = options.id;
  if (options.title) node.title = options.title;
  if (options.text !== undefined) node.textContent = String(options.text);

  const anyNode = node as unknown as Record<string, unknown>;
  if (options.type !== undefined) anyNode.type = options.type;
  if (options.value !== undefined) anyNode.value = options.value;
  if (options.placeholder !== undefined) anyNode.placeholder = options.placeholder;
  if (options.name !== undefined) anyNode.name = options.name;
  if (options.href !== undefined) anyNode.href = options.href;
  if (options.disabled !== undefined) anyNode.disabled = options.disabled;
  if (options.checked !== undefined) anyNode.checked = options.checked;
  if (options.required !== undefined) anyNode.required = options.required;
  if (options.min !== undefined) anyNode.min = options.min;
  if (options.max !== undefined) anyNode.max = options.max;
  if (options.step !== undefined) anyNode.step = options.step;
  if (options.maxLength !== undefined) anyNode.maxLength = options.maxLength;
  if (options.autocomplete !== undefined) anyNode.autocomplete = options.autocomplete;
  if (options.inputMode !== undefined) anyNode.inputMode = options.inputMode;
  if (options.rows !== undefined) anyNode.rows = options.rows;
  if (options.colSpan !== undefined) anyNode.colSpan = options.colSpan;

  if (options.dataset) {
    for (const [key, value] of Object.entries(options.dataset)) node.dataset[key] = value;
  }
  if (options.attrs) {
    for (const [key, value] of Object.entries(options.attrs)) node.setAttribute(key, value);
  }
  if (options.style) Object.assign(node.style, options.style);
  if (options.on) {
    for (const [event, handler] of Object.entries(options.on)) {
      node.addEventListener(event, handler as EventListener);
    }
  }

  append(node, children);
  return node;
}

export function append(parent: Node, children: Child[]): void {
  for (const child of children) {
    if (child === null || child === undefined || child === false) continue;
    parent.appendChild(typeof child === 'object' ? child : document.createTextNode(String(child)));
  }
}

export function clear(node: Node): void {
  while (node.firstChild) node.removeChild(node.firstChild);
}

export function replace(node: Node, children: Child[]): void {
  clear(node);
  append(node, children);
}

export function byId<T extends HTMLElement = HTMLElement>(id: string): T {
  const node = document.getElementById(id);
  if (!node) throw new Error(`DOM elementi topilmadi: #${id}`);
  return node as T;
}

export function query<T extends HTMLElement = HTMLElement>(selector: string, root: ParentNode = document): T | null {
  return root.querySelector<T>(selector);
}

export function queryAll<T extends HTMLElement = HTMLElement>(
  selector: string,
  root: ParentNode = document,
): T[] {
  return Array.from(root.querySelectorAll<T>(selector));
}

export interface Option {
  value: string;
  label: string;
  disabled?: boolean;
}

export function fillSelect(
  select: HTMLSelectElement,
  options: readonly Option[],
  selected?: string | null,
  placeholder?: string,
): void {
  clear(select);
  if (placeholder !== undefined) {
    select.appendChild(el('option', { value: '', text: placeholder }));
  }
  for (const option of options) {
    select.appendChild(
      el('option', { value: option.value, text: option.label, disabled: option.disabled ?? false }),
    );
  }
  select.value = selected ?? '';
  if (select.value === '' && selected) {
    // Tanlangan qiymat ro'yxatda yo'q (masalan nofaol qilingan) — birinchisiga qaytamiz.
    select.selectedIndex = 0;
  }
}

/** SVG elementlari uchun alohida yordamchi (createElement SVG uchun ishlamaydi). */
export function svgEl<K extends keyof SVGElementTagNameMap>(
  tag: K,
  attrs: Record<string, string | number> = {},
  children: Array<SVGElement | string> = [],
): SVGElementTagNameMap[K] {
  const node = document.createElementNS('http://www.w3.org/2000/svg', tag);
  for (const [key, value] of Object.entries(attrs)) node.setAttribute(key, String(value));
  for (const child of children) {
    node.appendChild(typeof child === 'string' ? document.createTextNode(child) : child);
  }
  return node;
}

/** Bir necha marta chaqirilsa ham oxirgisini bajaradi (qidiruv maydoni uchun). */
export function debounce<T extends unknown[]>(fn: (...args: T) => void, delay: number): (...args: T) => void {
  let timer: number | undefined;
  return (...args: T): void => {
    if (timer !== undefined) window.clearTimeout(timer);
    timer = window.setTimeout(() => fn(...args), delay);
  };
}

/** Tez-tez uchraydigan hodisalar uchun (mousemove) — v1 da bu yo'q edi. */
export function throttle<T extends unknown[]>(fn: (...args: T) => void, interval: number): (...args: T) => void {
  let last = 0;
  return (...args: T): void => {
    const now = Date.now();
    if (now - last < interval) return;
    last = now;
    fn(...args);
  };
}
