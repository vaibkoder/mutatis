import { defineConfig, type Plugin } from 'vite';
import preact from '@preact/preset-vite';
import basicSsl from '@vitejs/plugin-basic-ssl';
import { resolve, join } from 'path';
import { existsSync, mkdirSync, cpSync, copyFileSync, createReadStream } from 'fs';
import { bytesPlugin } from './scripts/bytes-plugin.mjs';

/**
 * Vendor PDF.js assets locally so the PDF compare path never fetches the worker,
 * cMaps, standard fonts, ICC profiles or wasm from a CDN (the PDF-offline guard,
 * confidentiality invariant ADR-0001/0003). Served same-origin at `/pdf-assets/`
 * in dev (middleware) and copied into `dist/pdf-assets/` on build. Keeping the
 * source in node_modules avoids committing thousands of asset files.
 */
function pdfjsAssets(): Plugin {
  const pkgRoot = resolve(__dirname, 'node_modules/pdfjs-dist');
  // [source-in-pkg, served-as] pairs.
  const dirs: Array<[string, string]> = [
    ['cmaps', 'cmaps'],
    ['standard_fonts', 'standard_fonts'],
    ['iccs', 'iccs'],
    ['wasm', 'wasm'],
  ];
  const files: Array<[string, string]> = [
    ['build/pdf.worker.min.mjs', 'pdf.worker.min.mjs'],
  ];
  return {
    name: 'mutatis-pdfjs-assets',
    configureServer(server) {
      server.middlewares.use('/pdf-assets', (req, res, next) => {
        const rel = decodeURIComponent((req.url ?? '').split('?')[0] ?? '').replace(/^\/+/, '');
        if (rel === '' || rel.includes('..')) return next();
        // Map served name back to its package location.
        const fileHit = files.find(([, served]) => served === rel);
        const src = fileHit ? join(pkgRoot, fileHit[0]) : join(pkgRoot, rel);
        if (!src.startsWith(pkgRoot) || !existsSync(src)) return next();
        if (rel.endsWith('.mjs')) res.setHeader('Content-Type', 'text/javascript');
        createReadStream(src).pipe(res);
      });
    },
    closeBundle() {
      const outRoot = resolve(__dirname, 'dist/pdf-assets');
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

export default defineConfig({
  plugins: [preact(), basicSsl(), pdfjsAssets(), bytesPlugin()],
  // The RenderCore worker (src/workers/render-core.worker.ts) code-splits — it
  // pulls in the engine + the same-origin harfbuzz wasm asset — so it must build
  // as an ES module (IIFE/UMD cannot do code-splitting). The bytes plugin is
  // applied here too so any transitive `?bytes` import resolves in the worker
  // bundle (zero runtime fetch — confidentiality invariant ADR-0001/0003).
  worker: {
    format: 'es',
    plugins: () => [bytesPlugin()],
  },
  optimizeDeps: {
    // Do NOT pre-bundle harfbuzzjs in dev. Its `index.mjs` self-instantiates via
    // a top-level await, which the dev dep-optimizer's esbuild target can't emit
    // (only `build.target:'es2022'` below covers the production bundle, not the
    // dev pre-bundler) — so the optimized chunk 504s and the worker's dynamic
    // `import('harfbuzzjs')` fails ("Importing a module script failed."). Served
    // raw from node_modules, `index.mjs`/`harfbuzz.js` load natively (TLA is fine
    // in a module worker) AND the emscripten factory's `new URL('harfbuzz.wasm',
    // import.meta.url)` resolves next to the real file — which it could not from
    // `.vite/deps/`. Dev-only; the production build (rollup) is unaffected.
    exclude: ['harfbuzzjs'],
  },
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
  build: {
    // es2022 so the RenderCore worker can carry harfbuzzjs's top-level await
    // (its index.mjs self-instantiates via `init(await createHarfBuzz())`). TLA
    // is es2022; all RenderCore-target engines (Office WebView2 Chromium,
    // Safari 15+, modern Chrome/Firefox) support it. The main app bundles use no
    // es2022-only syntax, so this only relaxes the floor where harfbuzz needs it.
    target: 'es2022',
    outDir: 'dist',
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, 'src/taskpane/index.html'),
        web: resolve(__dirname, 'src/web/index.html'),
      },
    },
  },
  server: {
    port: 3000,
    https: {},
  },
});
