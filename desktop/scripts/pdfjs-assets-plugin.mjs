import { cpSync, createReadStream, existsSync, mkdirSync, copyFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

/**
 * Vendor the PDF.js runtime assets next to the renderer so reading a PDF never
 * fetches a worker, character map or font from anywhere but this app.
 *
 * Served same-origin at `/pdf-assets/` in dev and copied into the renderer
 * output on build — the layout the engine's PDF opener
 * (@features/compare/io/pdf/pdfjs-loader) expects. Sourced from node_modules so
 * thousands of asset files are not committed. `wasm/` and `iccs/` are NOT
 * vendored: the engine opens documents with `useWasm: false` and nothing here
 * renders images, so neither is ever requested.
 */
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const pkgRoot = join(root, 'node_modules/pdfjs-dist');

const DIRS = ['cmaps', 'standard_fonts'];
const FILES = [['build/pdf.worker.min.mjs', 'pdf.worker.min.mjs']];

export function pdfjsAssetsPlugin(outDir) {
  return {
    name: 'mutatis-pdfjs-assets',
    configureServer(server) {
      server.middlewares.use('/pdf-assets', (req, res, next) => {
        const rel = decodeURIComponent((req.url ?? '').split('?')[0] ?? '').replace(/^\/+/, '');
        if (rel === '' || rel.includes('..')) return next();
        const hit = FILES.find(([, served]) => served === rel);
        const src = hit ? join(pkgRoot, hit[0]) : join(pkgRoot, rel);
        const allowed = hit || DIRS.some((d) => rel.startsWith(`${d}/`));
        if (!allowed || !src.startsWith(pkgRoot) || !existsSync(src)) return next();
        if (rel.endsWith('.mjs')) res.setHeader('Content-Type', 'text/javascript');
        createReadStream(src).pipe(res);
      });
    },
    closeBundle() {
      const target = join(outDir, 'pdf-assets');
      mkdirSync(target, { recursive: true });
      for (const dir of DIRS) cpSync(join(pkgRoot, dir), join(target, dir), { recursive: true });
      for (const [from, to] of FILES) copyFileSync(join(pkgRoot, from), join(target, to));
    },
  };
}
