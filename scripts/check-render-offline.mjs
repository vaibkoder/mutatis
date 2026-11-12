/**
 * RenderCore Phase 2 D1 — render-path built-bundle guard.
 *
 * Sibling to `check-pdf-offline.mjs` (which scans ALL of `dist/` for off-origin
 * CDN hosts — that egress gate already covers the render chunks). This guard
 * adds the two render-specific invariants from ADR-0024 / Phase 2:
 *
 *   1. HOST-FREE WALL at the artefact level — no render-path chunk may carry an
 *      Office/Word JS-API reference. The engine + adapter import only `@shared/*`;
 *      this asserts that survived bundling (the worker especially must be pure).
 *   2. BUNDLE BUDGET — the render path (worker + same-origin harfbuzz wasm + the
 *      lazy adapter/engine/font chunks) must stay within ~5 MB. It is all-new,
 *      lazy-loaded code, so its total IS the bundle delta. Logged either way.
 *
 *   npm run build && node scripts/check-render-offline.mjs
 *
 * Render-path assets are identified structurally: every `.wasm`, the
 * `render-core.worker-*.js` entry, and any JS chunk that references the worker
 * URL or the harfbuzz runtime (`render-core.worker` / `harfbuzz`). The wasm is
 * delivered SAME-ORIGIN (Vite-emitted) — that is offline and allowed; only
 * off-origin/CDN fetches are banned, and those are gated by check-pdf-offline.
 */
import { readdirSync, readFileSync, statSync, existsSync } from 'node:fs';
import { join } from 'node:path';

const DIST = 'dist';
const BUDGET_BYTES = 5 * 1024 * 1024; // ~5 MB render-path ceiling (ADR-0024)

// Office/Word JS-API surfaces that must never reach a render-path chunk.
const HOST_API_MARKERS = ['OfficeExtension', 'Word.run', 'Office.context', 'Office.onReady'];

// A JS chunk is render-path if it boots the worker or carries the harfbuzz runtime.
const RENDER_JS_MARKER = /render-core\.worker|harfbuzz/;

if (!existsSync(DIST)) {
  console.error(`check-render-offline: ${DIST}/ not found — run \`npm run build\` first.`);
  process.exit(2);
}

function walk(dir) {
  const out = [];
  for (const name of readdirSync(dir)) {
    const p = join(dir, name);
    const st = statSync(p);
    if (st.isDirectory()) out.push(...walk(p));
    else out.push(p);
  }
  return out;
}

const files = walk(DIST);
const renderAssets = [];

for (const file of files) {
  // PDF.js vendors its own wasm/assets under dist/pdf-assets/ — those belong to
  // the PDF compare path (gated by check-pdf-offline), not the render path.
  if (file.includes('/pdf-assets/')) continue;
  const base = file.split('/').pop() ?? '';
  if (file.endsWith('.wasm')) {
    renderAssets.push(file);
    continue;
  }
  if (/\.(js|mjs)$/.test(base)) {
    const text = readFileSync(file, 'utf8');
    if (/render-core\.worker-/.test(base) || RENDER_JS_MARKER.test(text)) {
      renderAssets.push(file);
    }
  }
}

if (renderAssets.length === 0) {
  console.error('check-render-offline: FAIL — no render-path chunks found in dist/. ' +
    'Did the build emit the RenderCore worker? (Expected after Phase 2 wiring.)');
  process.exit(1);
}

// 1. Host-free wall — no Office/Word API in any render-path chunk.
const hostHits = [];
let total = 0;
for (const file of renderAssets) {
  const size = statSync(file).size;
  total += size;
  if (!file.endsWith('.wasm')) {
    const text = readFileSync(file, 'utf8');
    for (const marker of HOST_API_MARKERS) {
      if (text.includes(marker)) hostHits.push(`${file}: ${marker}`);
    }
  }
}

console.log('check-render-offline: render-path assets —');
for (const file of renderAssets) {
  console.log(`  ${(statSync(file).size / 1024).toFixed(1).padStart(9)} kB  ${file}`);
}
console.log(`  ${'─'.repeat(9)}`);
console.log(`  ${(total / 1024).toFixed(1).padStart(9)} kB  total (budget ${(BUDGET_BYTES / 1024 / 1024).toFixed(0)} MB)`);

let failed = false;
if (hostHits.length > 0) {
  console.error('\ncheck-render-offline: FAIL — Office/Word API reference in a render-path chunk ' +
    '(the engine + adapter must be host-free — `shared ↛ features`, no Office import):');
  for (const h of hostHits) console.error('  ' + h);
  failed = true;
}

if (total > BUDGET_BYTES) {
  console.error(`\ncheck-render-offline: FAIL — render path ${(total / 1024 / 1024).toFixed(2)} MB ` +
    `exceeds the ~${(BUDGET_BYTES / 1024 / 1024).toFixed(0)} MB budget (ADR-0024).`);
  failed = true;
}

if (failed) process.exit(1);
console.log('\ncheck-render-offline: OK — render path host-free and within budget.');
