import { defineConfig, externalizeDepsPlugin } from 'electron-vite';
import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';
import { resolve } from 'node:path';
import { bytesPlugin } from './scripts/bytes-plugin.mjs';
import { engineAliases, engineDedupe } from './scripts/engine-aliases.mjs';
import { pdfjsAssetsPlugin } from './scripts/pdfjs-assets-plugin.mjs';

const here = import.meta.dirname;


export default defineConfig({
  main: {
    plugins: [externalizeDepsPlugin()],
    resolve: {
      alias: {
        ...engineAliases,
        '@main': resolve(here, 'src/main'),
        '@common': resolve(here, 'src/common'),
      },
    },
    build: {
      rollupOptions: {
        // app-protocol is emitted as its own entry so it can be loaded and
        // exercised in isolation (scripts/smoke.mjs) rather than only as part of
        // the bundled main entry. The app:// scheme and its CSP header are the
        // load-bearing security surface; they deserve a test that runs the real
        // module, not a copy of it.
        input: {
          index: resolve(here, 'src/main/index.ts'),
          'app-protocol': resolve(here, 'src/main/app-protocol.ts'),
        },
      },
    },
  },
  // main and preload deliberately set no build.target — electron-vite pins those
  // to the Node target of the bundled Electron and rejects any other value.
  preload: {
    plugins: [externalizeDepsPlugin()],
    resolve: { alias: { '@common': resolve(here, 'src/common') } },
    build: {
      // MUST be CommonJS. We run with sandbox:true, and Electron cannot load an
      // ESM preload into a sandboxed renderer — it fails at runtime, not at
      // build time, so nothing here would catch it.
      rollupOptions: {
        output: { format: 'cjs', entryFileNames: 'index.cjs' },
      },
    },
  },
  renderer: {
    root: resolve(here, 'src/renderer'),
    plugins: [react(), tailwindcss(), bytesPlugin(), pdfjsAssetsPlugin(resolve(here, 'out/renderer'))],
    resolve: {
      alias: {
        '@': resolve(here, 'src/renderer/src'),
        '@common': resolve(here, 'src/common'),
        ...engineAliases,
      },
      dedupe: engineDedupe,
    },
    // The bytes plugin must run in the worker bundle too — any transitive
    // `@fonts/X.ttf?bytes` import fails silently without it.
    worker: {
      format: 'es',
      plugins: () => [bytesPlugin()],
    },
    optimizeDeps: {
      // harfbuzzjs self-instantiates via a top-level await that the dev
      // dep-optimizer's esbuild target cannot emit; the optimized chunk 504s and
      // the dynamic import fails. Served raw, its `new URL('harfbuzz.wasm',
      // import.meta.url)` also resolves next to the real file. Dev-only —
      // the production rollup build is unaffected.
      exclude: ['harfbuzzjs'],
    },
    build: {
      // es2022 carries harfbuzzjs's top-level await. Required by RenderCore's
      // PDF/layout path (not by the prose redline, which needs no fonts).
      target: 'es2022',
      outDir: resolve(here, 'out/renderer'),
      emptyOutDir: true,
    },
  },
});
