/**
 * Root vite.config.ts — present for the shadcn CLI's project detection ONLY.
 *
 * The real build is driven by electron.vite.config.ts. shadcn's detection is
 * broken for Electron-shaped projects (PR #9632, open since Feb 2026): without a
 * root vite.config.ts it reports "We could not detect a supported framework", and
 * an EMPTY file does not satisfy it — it must have real contents. This file is
 * never used to build anything.
 */
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';
import { resolve } from 'node:path';

export default defineConfig({
  plugins: [react(), tailwindcss()],
  resolve: {
    alias: {
      '@': resolve(import.meta.dirname, 'src/renderer/src'),
    },
  },
});
