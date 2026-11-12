import { defineConfig } from 'vitest/config';
import { resolve } from 'path';
import { bytesPlugin } from './scripts/bytes-plugin.mjs';

export default defineConfig({
  plugins: [bytesPlugin()],
  resolve: {
    alias: {
      '@shared': resolve(__dirname, 'src/shared'),
      '@taskpane': resolve(__dirname, 'src/taskpane'),
      '@features': resolve(__dirname, 'src/features'),
      '@host-office': resolve(__dirname, 'src/host-office'),
      '@host-web': resolve(__dirname, 'src/host-web'),
      '@fonts': resolve(__dirname, 'public/fonts'),
    },
  },
  esbuild: {
    jsx: 'automatic',
    jsxImportSource: 'preact',
  },
  test: {
    environment: 'jsdom',
    globals: true,
    pool: 'forks',
    include: ['tests/**/*.test.ts', 'tests/**/*.test.tsx'],
    setupFiles: ['tests/setup.ts'],
  },
});
