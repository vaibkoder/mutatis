/**
 * STANDALONE / EXPERIMENTAL build config — NOT a product target.
 *
 * Packages the existing Office-free web app (`src/web/`, ADR-0022) for trying in
 * a plain browser, separate from the Word add-in build (`vite.config.ts`). It does
 * NOT touch the add-in, the manifest, or any product surface — it only re-bundles
 * code that already ships, so confidentiality holds by construction (no backend,
 * no network; ADR-0001/0003).
 *
 * Two modes, selected by the DP_TARGET env var:
 *   DP_TARGET=served  → multi-file bundle in `dist-web/` (index.html at root,
 *                       `/assets/*`, `/pdf-assets/*`). Serve the folder over http
 *                       and EVERY feature works, incl. PDF compare + rendered
 *                       redline preview (Web Workers + wasm need an http origin).
 *   DP_TARGET=single  → one self-contained `dist-single/index.html` with JS/CSS
 *                       inlined. Double-click it (file://) and the pure-JS
 *                       features work: Proof checkers + Word/Text redline .docx.
 *                       PDF compare + the rendered preview can't run from file://
 *                       (Chrome blocks workers + same-origin asset fetches there).
 *
 * Run via the npm scripts `build:web` / `build:single`.
 */
import { defineConfig, type Plugin } from 'vite';
import preact from '@preact/preset-vite';
import { viteSingleFile } from 'vite-plugin-singlefile';
import { resolve, join } from 'path';
import { existsSync, mkdirSync, cpSync, copyFileSync, readdirSync, rmSync } from 'fs';
import { bytesPlugin } from './scripts/bytes-plugin.mjs';

const SINGLE = process.env.DP_TARGET === 'single';
const outDir = resolve(__dirname, SINGLE ? 'dist-single' : 'dist-web');

/** Copy the vendored PDF.js assets into the served bundle (no CDN — ADR-0001/0003). */
function pdfjsAssetsCopy(): Plugin {
  const pkgRoot = resolve(__dirname, 'node_modules/pdfjs-dist');
  const dirs: Array<[string, string]> = [
    ['cmaps', 'cmaps'],
    ['standard_fonts', 'standard_fonts'],
    ['iccs', 'iccs'],
    ['wasm', 'wasm'],
  ];
  const files: Array<[string, string]> = [['build/pdf.worker.min.mjs', 'pdf.worker.min.mjs']];
  return {
    name: 'mutatis-standalone-pdfjs-assets',
    closeBundle() {
      const outRoot = join(outDir, 'pdf-assets');
      mkdirSync(outRoot, { recursive: true });
      for (const [from, to] of dirs) {
        const src = join(pkgRoot, from);
        if (existsSync(src)) cpSync(src, join(outRoot, to), { recursive: true });
      }
      for (const [from, to] of files) {
        const src = join(pkgRoot, from);
        if (existsSync(src)) copyFileSync(src, join(outRoot, to));
      }
    },
  };
}

/**
 * SINGLE mode leaves the entry fully inlined in `index.html`, but vite still emits
 * the (now-unreferenced) Worker build's sibling chunks. Delete them so the output
 * dir is just the one self-contained file — nothing to mistakenly copy/forget.
 */
function pruneToIndexHtml(): Plugin {
  return {
    name: 'mutatis-standalone-prune',
    closeBundle() {
      for (const name of readdirSync(outDir)) {
        if (name !== 'index.html') rmSync(join(outDir, name), { recursive: true, force: true });
      }
    },
  };
}

export default defineConfig({
  // `src/web/index.html` becomes the bundle root so the output is a top-level
  // `index.html` (cleaner to serve / double-click than a nested path).
  root: resolve(__dirname, 'src/web'),
  plugins: [
    preact(),
    bytesPlugin(),
    // Single-file inlines JS/CSS into the html; the served bundle keeps the
    // vendored PDF.js assets as same-origin files (can't be inlined and still run).
    ...(SINGLE ? [viteSingleFile(), pruneToIndexHtml()] : [pdfjsAssetsCopy()]),
  ],
  worker: { format: 'es', plugins: () => [bytesPlugin()] },
  optimizeDeps: { exclude: ['harfbuzzjs'] },
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
  // SINGLE mode: force the in-process (main-thread) render client so there is no
  // Worker chunk to fetch — `file://` blocks `new Worker(...)`. The wasm is already
  // a data-URL asset, so the PDF redline then works from a double-clicked file.
  define: SINGLE ? { 'import.meta.env.VITE_DP_INPROCESS': JSON.stringify('1') } : {},
  build: {
    target: 'es2022',
    outDir,
    emptyOutDir: true,
    // Quieten the (expected) large-chunk warning for the single-file bundle.
    chunkSizeWarningLimit: 8000,
    // SINGLE mode: collapse every dynamic import() into the entry chunk so the one
    // inlined HTML needs no sibling .js — `file://` can't dynamic-import a relative
    // file (origin null / CORS). The served bundle keeps normal code-splitting.
    ...(SINGLE ? { rollupOptions: { output: { inlineDynamicImports: true } } } : {}),
  },
});
