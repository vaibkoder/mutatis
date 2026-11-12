/**
 * A1 confidentiality guard — LAYER 1 (built-bundle scan).
 *
 * After `npm run build`, scans the emitted `dist/` JavaScript for any hardcoded
 * off-origin CDN reference that could make PDF.js (or anything else) fetch an
 * asset off-machine. Exits non-zero on a hit so it can gate CI / pre-commit.
 *
 *   npm run build && node scripts/check-pdf-offline.mjs
 *
 * This complements the static unit guard (tests/unit/pdf/pdfjs-offline.guard.test.ts)
 * and the runtime Playwright zero-network spec (tests/integration/pdf-network.spec.ts).
 */
import { readdirSync, readFileSync, statSync, existsSync } from 'node:fs';
import { join } from 'node:path';

const DIST = 'dist';
const FORBIDDEN = [
  'cdn.jsdelivr.net',
  'unpkg.com',
  'cdnjs.cloudflare.com',
  'mozilla.github.io',
  'fonts.googleapis.com',
  'fonts.gstatic.com',
  // PDF.js's documented default CDN base, in case a config slips through.
  'cdn.jsdelivr.net/npm/pdfjs-dist',
];

if (!existsSync(DIST)) {
  console.error(`check-pdf-offline: ${DIST}/ not found — run \`npm run build\` first.`);
  process.exit(2);
}

/** Recursively collect text-ish asset files from dist. */
function walk(dir) {
  const out = [];
  for (const name of readdirSync(dir)) {
    const p = join(dir, name);
    const st = statSync(p);
    if (st.isDirectory()) out.push(...walk(p));
    else if (/\.(js|mjs|css|html|map)$/.test(name)) out.push(p);
  }
  return out;
}

const hits = [];
for (const file of walk(DIST)) {
  const text = readFileSync(file, 'utf8');
  for (const host of FORBIDDEN) {
    if (text.includes(host)) hits.push(`${file}: ${host}`);
  }
}

if (hits.length > 0) {
  console.error('check-pdf-offline: FAIL — off-origin asset reference(s) in the built bundle:');
  for (const h of hits) console.error('  ' + h);
  console.error('All PDF.js assets must be vendored locally (ADR-0001/0003).');
  process.exit(1);
}
console.log('check-pdf-offline: OK — no off-origin asset references in dist/.');
