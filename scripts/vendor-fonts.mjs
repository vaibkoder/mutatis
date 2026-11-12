/**
 * Dev-time font vendoring (RenderCore, feature 16 / ADR-0024, Phase 2 Task A1).
 *
 * Acquires the OFL metric-clone set and writes raw **sfnt (TTF)** bytes + OFL
 * licences into `public/fonts/`. The engine needs sfnt for HarfBuzz shaping and
 * fontkit subset embedding (WOFF/WOFF2 are compressed containers it cannot read
 * directly), so this decodes WOFF v1 → sfnt in-process.
 *
 *   Times New Roman → Liberation Serif  (vendored as Tinos — same TNR metrics)
 *   Arial           → Liberation Sans   (from pdfjs-dist standard_fonts)
 *   Calibri         → Carlito
 *   Cambria         → Caladea
 *
 * This is a **dev-time** tool: it uses `npm pack` (network) ONCE; its OUTPUT (the
 * committed `public/fonts/*.ttf`) is what ships, embedded via the `?bytes` plugin
 * with ZERO runtime egress (ADR-0001/0003). Re-runnable + deterministic.
 *
 *   node scripts/vendor-fonts.mjs
 */
import { execFileSync } from 'node:child_process';
import { inflateSync } from 'node:zlib';
import {
  mkdirSync, readFileSync, writeFileSync, existsSync, readdirSync, rmSync, copyFileSync,
} from 'node:fs';
import { resolve, join } from 'node:path';

const root = resolve(import.meta.dirname, '..');
const OUT = join(root, 'public', 'fonts');
const TMP = join(root, '.fontprobe');
mkdirSync(OUT, { recursive: true });
mkdirSync(TMP, { recursive: true });

// ── WOFF v1 → sfnt ───────────────────────────────────────────────────────────
function woffToSfnt(woff) {
  const u = new DataView(woff.buffer, woff.byteOffset, woff.byteLength);
  if (u.getUint32(0) !== 0x774f4646) throw new Error('not a WOFF file'); // 'wOFF'
  const flavor = u.getUint32(4);
  const numTables = u.getUint16(12);

  // Directory entries (20 bytes each) start at byte 44.
  const entries = [];
  for (let i = 0; i < numTables; i++) {
    const o = 44 + i * 20;
    entries.push({
      tag: u.getUint32(o + 0),
      offset: u.getUint32(o + 4),
      compLength: u.getUint32(o + 8),
      origLength: u.getUint32(o + 12),
      checksum: u.getUint32(o + 16),
    });
  }
  // sfnt requires table records sorted by tag; data may be in any order.
  entries.sort((a, b) => a.tag - b.tag);

  const pad4 = (n) => (n + 3) & ~3;
  let dataLen = 0;
  for (const e of entries) dataLen += pad4(e.origLength);
  const headerLen = 12 + numTables * 16;
  const out = Buffer.alloc(headerLen + dataLen);

  // Offset table.
  const maxPow2 = Math.pow(2, Math.floor(Math.log2(numTables)));
  out.writeUInt32BE(flavor, 0);
  out.writeUInt16BE(numTables, 4);
  out.writeUInt16BE(maxPow2 * 16, 6);                     // searchRange
  out.writeUInt16BE(Math.floor(Math.log2(numTables)), 8); // entrySelector
  out.writeUInt16BE(numTables * 16 - maxPow2 * 16, 10);   // rangeShift

  let cursor = headerLen;
  entries.forEach((e, i) => {
    const raw = woff.subarray(e.offset, e.offset + e.compLength);
    const tbl = e.compLength === e.origLength ? raw : inflateSync(raw);
    if (tbl.length !== e.origLength) throw new Error(`table length mismatch tag=${e.tag.toString(16)}`);
    tbl.copy(out, cursor);
    const rec = 12 + i * 16;
    out.writeUInt32BE(e.tag, rec + 0);
    out.writeUInt32BE(e.checksum, rec + 4);
    out.writeUInt32BE(cursor, rec + 8);
    out.writeUInt32BE(e.origLength, rec + 12);
    cursor += pad4(e.origLength);
  });
  return out;
}

// ── Acquire one @fontsource clone via npm pack, decode its 4 styles ───────────
const STYLES = [
  ['400', 'normal', 'Regular', false, false],
  ['700', 'normal', 'Bold', true, false],
  ['400', 'italic', 'Italic', false, true],
  ['700', 'italic', 'BoldItalic', true, true],
];

function extractTgz(tgz, dest) {
  rmSync(dest, { recursive: true, force: true });
  mkdirSync(dest, { recursive: true });
  execFileSync('tar', ['xzf', tgz, '-C', dest], { stdio: 'ignore' });
  return join(dest, 'package');
}

// On-disk file prefixes are space-free (import-specifier friendly); the logical
// clone family names (with spaces, matching registry CLONE_MAP) live in faces.ts.
function vendorFontsource(pkg, filePrefix) {
  execFileSync('npm', ['pack', `@fontsource/${pkg}`, '--pack-destination', TMP], {
    cwd: root, stdio: 'ignore',
  });
  const tgz = readdirSync(TMP).filter((f) => f.startsWith(`fontsource-${pkg}-`)).sort().pop();
  const dir = extractTgz(join(TMP, tgz), join(TMP, pkg));
  const files = join(dir, 'files');
  for (const [w, s, label] of STYLES) {
    const woffName = `${pkg}-latin-${w}-${s}.woff`;
    const sfnt = woffToSfnt(readFileSync(join(files, woffName)));
    writeFileSync(join(OUT, `${filePrefix}-${label}.ttf`), sfnt);
  }
  copyFileSync(join(dir, 'LICENSE'), join(OUT, `OFL-${filePrefix}.txt`));
}

// ── Liberation Sans (Arial clone) — already offline in pdfjs-dist ─────────────
function vendorLiberationSans() {
  const src = join(root, 'node_modules', 'pdfjs-dist', 'standard_fonts');
  const map = [
    ['LiberationSans-Regular.ttf', 'LiberationSans-Regular.ttf'],
    ['LiberationSans-Bold.ttf', 'LiberationSans-Bold.ttf'],
    ['LiberationSans-Italic.ttf', 'LiberationSans-Italic.ttf'],
    ['LiberationSans-BoldItalic.ttf', 'LiberationSans-BoldItalic.ttf'],
  ];
  for (const [from, to] of map) copyFileSync(join(src, from), join(OUT, to));
  if (existsSync(join(src, 'LICENSE_LIBERATION'))) {
    copyFileSync(join(src, 'LICENSE_LIBERATION'), join(OUT, 'OFL-LiberationSans.txt'));
  }
}

console.log('Vendoring OFL metric-clone fonts → public/fonts/');
vendorLiberationSans();
vendorFontsource('tinos', 'LiberationSerif'); // Tinos == Liberation Serif (TNR metrics)
vendorFontsource('carlito', 'Carlito');       // Calibri clone
vendorFontsource('caladea', 'Caladea');       // Cambria clone

// ── Manifest ──────────────────────────────────────────────────────────────────
let total = 0;
for (const f of readdirSync(OUT).filter((f) => f.endsWith('.ttf')).sort()) {
  const sz = readFileSync(join(OUT, f)).length;
  total += sz;
  console.log(`  ${f.padEnd(34)} ${(sz / 1024).toFixed(1)} KB`);
}
console.log(`  TOTAL sfnt ${(total / 1024 / 1024).toFixed(2)} MB across 16 faces`);
rmSync(TMP, { recursive: true, force: true });
