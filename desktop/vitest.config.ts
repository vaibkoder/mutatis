import { defineConfig } from 'vitest/config';
import { bytesPlugin } from './scripts/bytes-plugin.mjs';
import { resolve } from 'node:path';
import { engineAliases, engineDedupe } from './scripts/engine-aliases.mjs';

export default defineConfig({
  // The bytes plugin must be present in the TEST runner too, not just the
  // bundler — otherwise every `?bytes` import fails the moment a test touches
  // the render path.
  plugins: [bytesPlugin()],
  resolve: {
    // Not optional, and not inherited from the bundler config. The engine
    // aliases resolve in tsc and in `vite build` and STILL throw under vitest
    // without this — the three paths are independent.
    alias: { ...engineAliases, '@common': resolve(import.meta.dirname, 'src/common') },
    dedupe: engineDedupe,
  },
  test: {
    environment: 'node',
    include: ['tests/**/*.test.ts'],
  },
});
