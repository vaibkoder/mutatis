import { createRequire } from 'module';
import tseslint from '@typescript-eslint/eslint-plugin';
import tsParser from '@typescript-eslint/parser';
import globals from 'globals';

const require = createRequire(import.meta.url);
const invariants = require('./eslint-plugins/invariants.cjs');

export default [
  {
    ignores: ['node_modules/**', 'dist/**', 'spec/**', 'docs/**', '**/*.md',
      // USER-RUN Playwright gate. `@playwright/test` is deliberately NOT in the
      // toolchain (large dependency + browser binaries) and vitest only collects
      // `*.test.ts`, so this file is never installed, linted or run here. Excluded
      // rather than `@ts-nocheck`-suppressed, which would weaken ban-ts-comment.
      'tests/integration/pdf-network.spec.ts'],
  },
  {
    files: ['src/**/*.{ts,tsx}'],
    languageOptions: {
      parser: tsParser,
      parserOptions: {
        ecmaVersion: 2020,
        sourceType: 'module',
        ecmaFeatures: { jsx: true },
      },
      globals: {
        ...globals.browser,
        ...globals.es2020,
      },
    },
    plugins: {
      '@typescript-eslint': tseslint,
      'local-invariants': invariants,
    },
    rules: {
      ...tseslint.configs.recommended.rules,
      '@typescript-eslint/no-unused-vars': ['error', { argsIgnorePattern: '^_' }],
      '@typescript-eslint/no-explicit-any': 'error',
      'local-invariants/no-network-egress': 'error',
      'local-invariants/no-us-literals': 'error',
      'local-invariants/no-cut-feature-refs': 'error',
      'local-invariants/no-host-imports': 'error',
    },
  },
  {
    files: ['eslint-plugins/**/*.cjs'],
    languageOptions: {
      sourceType: 'commonjs',
      globals: globals.node,
    },
  },
  {
    files: ['tests/**/*.{ts,tsx}'],
    languageOptions: {
      parser: tsParser,
      parserOptions: { ecmaVersion: 2020, sourceType: 'module' },
      globals: { ...globals.browser, ...globals.es2020 },
    },
    plugins: {
      '@typescript-eslint': tseslint,
      'local-invariants': invariants,
    },
    rules: {
      ...tseslint.configs.recommended.rules,
      // Same shape as `src`, plus `ignoreRestSiblings` for the rest-omit idiom
      // (`const { source: _omitted, ...rest } = x`) that tests use to drop a key.
      '@typescript-eslint/no-unused-vars': ['error', {
        argsIgnorePattern: '^_',
        varsIgnorePattern: '^_',
        caughtErrorsIgnorePattern: '^_',
        ignoreRestSiblings: true,
      }],
      // `eslint-plugins/invariants.cjs` exports FOUR rules. Three are `src`-only
      // by design — a test must be free to write a US literal, name a cut feature
      // or import a host in order to prove the shipped code rejects it.
      'local-invariants/no-us-literals': 'off',
      'local-invariants/no-cut-feature-refs': 'off',
      'local-invariants/no-host-imports': 'off',
      // The fourth is NOT one of those, and leaving it out of this block was an
      // omission rather than a decision. ADR-0001/ADR-0003 are the catastrophic
      // invariant: no document content leaves the machine. A network call is
      // exactly as forbidden in a fixture loader or a test helper as in shipped
      // code — more so, because helpers get promoted into `src`.
      'local-invariants/no-network-egress': 'error',
    },
  },
];
