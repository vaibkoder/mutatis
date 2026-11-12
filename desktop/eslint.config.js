import js from '@eslint/js';
import tseslint from 'typescript-eslint';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const invariants = require('./eslint-plugins/invariants.cjs');

export default tseslint.config(
  { ignores: ['out/**', 'dist/**', 'release/**', 'node_modules/**'] },

  js.configs.recommended,

  // The invariant plugin is CommonJS by design — ESLint plugins load as CJS.
  {
    files: ['eslint-plugins/**/*.cjs'],
    languageOptions: {
      sourceType: 'commonjs',
      globals: { module: 'writable', require: 'readonly', __dirname: 'readonly' },
    },
  },
  ...tseslint.configs.recommended,

  {
    files: ['src/**/*.{ts,tsx}'],
    plugins: { invariants },
    rules: {
      // The one catastrophic invariant. See eslint-plugins/invariants.cjs.
      'invariants/no-renderer-provider-access': 'error',
      '@typescript-eslint/no-unused-vars': [
        'error',
        { argsIgnorePattern: '^_', varsIgnorePattern: '^_' },
      ],
    },
  },

  // The engine lives upstream and is never linted from here — it is read-only.
  { ignores: ['../Mutatis/**'] },
);
