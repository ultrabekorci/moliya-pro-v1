#!/usr/bin/env node
/**
 * Qurish (build) skripti.
 *
 * TypeScript manbani Google Apps Script tushunadigan ikkita faylga aylantiradi:
 *   dist/Code.gs     — server kodi (bitta IIFE + global funksiyalar qobig'i)
 *   dist/index.html  — klient (CSS va JS ichiga joylashtirilgan, CDN yo'q)
 *   dist/appsscript.json — manifest
 *
 * `clasp push` shu `dist` papkasini yuklaydi.
 */

import { build } from 'esbuild';
import { mkdir, readFile, rm, writeFile, copyFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const dist = resolve(root, 'dist');

/** Apps Script global funksiyalari — bundle ichidagi eksportlarga ko'prik. */
const GLOBAL_ENTRY_POINTS = [
  ['doGet', 'e'],
  ['rpc', 'payload'],
  ['setup', ''],
  ['purgeSessions', ''],
  ['authorizeExternalRequests', ''],
  ['migrateDryRun', ''],
  ['migrateFromV1', ''],
];

function globalsShim(namespace) {
  const lines = GLOBAL_ENTRY_POINTS.map(([name, args]) => {
    const params = args === '' ? '' : args;
    return `function ${name}(${params}) {\n  return ${namespace}.${name}(${params});\n}`;
  });
  return [
    '',
    '// --- Google Apps Script global kirish nuqtalari (build tomonidan yaratilgan) ---',
    ...lines,
    '',
  ].join('\n');
}

async function buildServer() {
  const result = await build({
    entryPoints: [resolve(root, 'src/server/main.ts')],
    bundle: true,
    write: false,
    format: 'iife',
    globalName: 'MoliyaProServer',
    platform: 'neutral',
    target: 'es2019',
    charset: 'utf8',
    legalComments: 'none',
    logLevel: 'warning',
  });

  const output = result.outputFiles[0];
  if (!output) throw new Error('Server bundle yaratilmadi');

  const banner = [
    '/**',
    ' * Moliya-Pro — server kodi.',
    ' * DIQQAT: bu fayl avtomatik yaratilgan. Manba: src/server/**.ts',
    ' * O‘zgartirish uchun TypeScript manbani tahrirlang va `npm run build` ni ishga tushiring.',
    ' */',
    '',
  ].join('\n');

  await writeFile(resolve(dist, 'Code.gs'), banner + output.text + globalsShim('MoliyaProServer'), 'utf8');
}

async function buildClient() {
  const result = await build({
    entryPoints: [resolve(root, 'src/client/main.ts')],
    bundle: true,
    write: false,
    format: 'iife',
    platform: 'browser',
    target: ['es2019'],
    charset: 'utf8',
    minify: process.env.NODE_ENV !== 'development',
    legalComments: 'none',
    logLevel: 'warning',
  });

  const script = result.outputFiles[0];
  if (!script) throw new Error('Klient bundle yaratilmadi');

  const cssResult = await build({
    entryPoints: [resolve(root, 'src/client/styles.css')],
    bundle: true,
    write: false,
    minify: process.env.NODE_ENV !== 'development',
    logLevel: 'warning',
  });
  const css = cssResult.outputFiles[0];
  if (!css) throw new Error('CSS bundle yaratilmadi');

  const template = await readFile(resolve(root, 'src/client/index.html'), 'utf8');
  const html = template
    .replace('/*__STYLES__*/', () => css.text)
    .replace('/*__SCRIPT__*/', () => script.text);

  assertSelfContained(html);
  await writeFile(resolve(dist, 'index.html'), html, 'utf8');
}

/**
 * Sahifa hech qanday tashqi manbaga murojaat qilmasligi kerak: barcha CSS va JS
 * ichiga joylashtirilgan. v1 da to'rtta CDN'dan kutubxona yuklanardi (ikkitasi
 * ikki marta), bu esa sahifani begona serverlarga bog'lab qo'yardi.
 */
function assertSelfContained(html) {
  const external = [
    /<script[^>]+src\s*=/i,
    /<link[^>]+rel\s*=\s*["']?stylesheet/i,
    /https?:\/\/(cdn|unpkg|fonts|ajax)\./i,
  ];
  for (const pattern of external) {
    if (pattern.test(html)) {
      throw new Error(`index.html tashqi manbaga murojaat qilmoqda: ${pattern}`);
    }
  }
}

async function main() {
  await rm(dist, { recursive: true, force: true });
  await mkdir(dist, { recursive: true });

  await Promise.all([buildServer(), buildClient()]);
  await copyFile(resolve(root, 'appsscript.json'), resolve(dist, 'appsscript.json'));

  console.log('✔ dist/Code.gs, dist/index.html, dist/appsscript.json tayyor');
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
