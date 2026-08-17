import js from '@eslint/js';
import tseslint from '@typescript-eslint/eslint-plugin';
import tsparser from '@typescript-eslint/parser';
import globals from 'globals';

export default [
  {
    ignores: ['dist/**', 'coverage/**', 'node_modules/**'],
  },
  js.configs.recommended,
  {
    files: ['src/**/*.ts', 'tests/**/*.ts'],
    languageOptions: {
      parser: tsparser,
      parserOptions: {
        ecmaVersion: 2022,
        sourceType: 'module',
      },
      globals: {
        ...globals.browser,
        ...globals.node,
        // Google Apps Script muhitidagi global obyektlar.
        SpreadsheetApp: 'readonly',
        PropertiesService: 'readonly',
        CacheService: 'readonly',
        LockService: 'readonly',
        UrlFetchApp: 'readonly',
        Utilities: 'readonly',
        HtmlService: 'readonly',
        Session: 'readonly',
        Logger: 'readonly',
        GoogleAppsScript: 'readonly',
      },
    },
    plugins: {
      '@typescript-eslint': tseslint,
    },
    rules: {
      ...tseslint.configs.recommended.rules,
      'no-undef': 'off', // TypeScript o'zi tekshiradi.
      'no-unused-vars': 'off',
      '@typescript-eslint/no-unused-vars': ['error', { argsIgnorePattern: '^_' }],
      '@typescript-eslint/no-explicit-any': 'error',
      '@typescript-eslint/no-non-null-assertion': 'off',
      eqeqeq: ['error', 'always'],
      'no-console': ['warn', { allow: ['warn', 'error', 'log'] }],
      'prefer-const': 'error',
    },
  },
  {
    files: ['scripts/**/*.mjs'],
    languageOptions: {
      globals: { ...globals.node },
    },
  },
];
