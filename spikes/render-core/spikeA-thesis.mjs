// Spike A — the thesis test (THROWAWAY, not bundled).
//
// Proves the RenderCore thesis on real font bytes, fully offline:
//  (1) METRIC-CLONE IDENTITY — Liberation Sans (OFL) advance widths == Arial
//      (the source MS font) with kern/liga off → identical line breaks.
//  (2) usWin LINE-HEIGHT formula computes from OS/2 (vs typo metrics).
//  (3) GREEDY first-fit breaker + Word JUSTIFICATION (space expansion +
//      up-to-~20% inter-word shrink) fills an A4 column with no overflow.
//  (4) Emits an HTML overlay (positioned runs, embedded clone font) for the
//      user to eyeball against Word's rendering of the same text.
//
// Run: node spikes/render-core/spikeA-thesis.mjs
import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import * as fontkit from 'fontkit';
import { loadFont, shapeRun, widthPt } from './hb-util.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const asset = (f) => join(HERE, 'assets', f);
const outDir = join(HERE, 'out');

const PT_PER_IN = 72;
const MM_PER_IN = 25.4;
const mm = (x) => (x / MM_PER_IN) * PT_PER_IN;

// A4 portrait, Word default 1-inch (25.4mm) margins.
const PAGE_W = mm(210);
const MARGIN = PT_PER_IN; // 1 inch
const COL_W = PAGE_W - 2 * MARGIN; // ≈ 451.3 pt
const SIZE = 11; // 11pt body

// ── (1) metric-clone identity ───────────────────────────────────────────────
const lib = loadFont(asset('LiberationSans-Regular.ttf'));
let arial = null;
try { arial = loadFont(asset('Arial-Reference.ttf')); } catch { /* reference optional */ }

const SAMPLE_WORDS = [
  'Agreement', 'Confidential', 'Information', 'hereinafter', 'Parties',
  'indemnification', 'jurisdiction', 'Singapore', 'obligations', 'pursuant',
  'Clause', 'Schedule', 'whatsoever', 'notwithstanding', 'representations',
];
let maxDeltaPct = 0, sumDeltaPct = 0;
const rows = [];
for (const w of SAMPLE_WORDS) {
  const wl = widthPt(lib.font, lib.upem, w, SIZE);
  const wa = arial ? widthPt(arial.font, arial.upem, w, SIZE) : NaN;
  const dPct = arial ? Math.abs(wl - wa) / wa * 100 : 0;
  maxDeltaPct = Math.max(maxDeltaPct, dPct);
  sumDeltaPct += dPct;
  rows.push({ w, wl: wl.toFixed(3), wa: arial ? wa.toFixed(3) : 'n/a', dPct: dPct.toFixed(3) });
}

// ── (2) usWin line-height ────────────────────────────────────────────────────
function lineMetrics(path) {
  const f = fontkit.openSync(path);
  const os2 = f['OS/2'];
  const upem = f.unitsPerEm;
  const useTypo = os2 ? (os2.fsSelection & 0x80) !== 0 : false; // USE_TYPO_METRICS bit 7
  const winAsc = os2?.winAscent ?? f.ascent;
  const winDesc = os2?.winDescent ?? Math.abs(f.descent);
  const typoAsc = os2?.typoAscender ?? f.ascent;
  const typoDesc = os2?.typoDescender ?? f.descent;
  const typoGap = os2?.typoLineGap ?? f.lineGap ?? 0;
  const winSingle = ((winAsc + winDesc) / upem) * SIZE;
  const typoSingle = ((typoAsc - typoDesc + typoGap) / upem) * SIZE;
  return { upem, useTypo, winAsc, winDesc, winSingle, typoSingle };
}
const lm = lineMetrics(asset('LiberationSans-Regular.ttf'));
const singleLineHeight = lm.useTypo ? lm.typoSingle : lm.winSingle;

// ── (3) greedy first-fit + Word justification ────────────────────────────────
const PARA =
  'The Borrower shall, on each Interest Payment Date and in accordance with ' +
  'Clause 8.2 (Payment mechanics), pay to the Agent for the account of each ' +
  'Lender the accrued interest on the Loan which has accrued during the ' +
  'Interest Period ending on that date, provided that no such payment shall ' +
  'be required to the extent that it would cause the aggregate outstanding ' +
  'amount to exceed the Total Commitments then in force under this Agreement.';

const SPACE_W = widthPt(lib.font, lib.upem, ' ', SIZE);
const SHRINK_MAX = 0.20; // Word 2013+ up-to-20% inter-word shrink

function wordWidth(w) { return shapeRun(lib.font, w).widthFu / lib.upem * SIZE; }

function breakGreedy(text, colW) {
  const words = text.split(/\s+/).filter(Boolean);
  const widths = words.map(wordWidth);
  const lines = [];
  let i = 0;
  while (i < words.length) {
    let j = i;
    let natural = widths[i]; // sum of word widths, no spaces yet
    let count = 1;
    // add words while they fit with normal spaces
    while (j + 1 < words.length) {
      const tentative = natural + SPACE_W + widths[j + 1];
      if (tentative <= colW) { natural = tentative; j++; count++; }
      else {
        // Word shrink rule: fit one more if existing spaces can shrink ≤20%.
        const gaps = count; // inter-word gaps if we add the next word
        const overflow = tentative - colW;
        const shrinkable = gaps * SPACE_W * SHRINK_MAX;
        if (gaps > 0 && overflow <= shrinkable) { natural = tentative; j++; count++; }
        break;
      }
    }
    lines.push({ from: i, to: j, words: words.slice(i, j + 1), widths: widths.slice(i, j + 1) });
    i = j + 1;
  }
  return { words, widths, lines };
}

const { lines } = breakGreedy(PARA, COL_W);
const lineReport = lines.map((ln, idx) => {
  const isLast = idx === lines.length - 1;
  const wordSum = ln.widths.reduce((a, b) => a + b, 0);
  const gaps = ln.words.length - 1;
  const naturalWidth = wordSum + gaps * SPACE_W;
  let spaceW = SPACE_W, finalWidth = naturalWidth, mode = 'flush-left(last)';
  if (!isLast && gaps > 0) {
    spaceW = (COL_W - wordSum) / gaps; // justify: solve so line == COL_W
    finalWidth = COL_W;
    mode = spaceW >= SPACE_W ? `expand×${(spaceW / SPACE_W).toFixed(3)}` : `shrink×${(spaceW / SPACE_W).toFixed(3)}`;
  }
  return { idx, words: ln.words.length, wordsArr: ln.words, wordSum, gaps, spaceW, finalWidth, mode, text: ln.words.join(' ') };
});

const overflow = lineReport.filter((l) => l.finalWidth > COL_W + 0.01);
const maxShrink = Math.min(1, ...lineReport.filter((l) => l.mode.startsWith('shrink')).map((l) => l.spaceW / SPACE_W));

// ── (4) HTML overlay (positioned runs, embedded clone font) ──────────────────
const fontB64 = readFileSync(asset('LiberationSans-Regular.ttf')).toString('base64');
let html = `<!doctype html><meta charset=utf8><title>Spike A overlay</title>
<style>
@font-face{font-family:'LibClone';src:url(data:font/ttf;base64,${fontB64}) format('truetype');}
body{margin:0;background:#888;font-family:'LibClone';}
.page{position:relative;width:${PAGE_W}pt;height:${mm(297)}pt;background:#fff;margin:16px auto;box-shadow:0 2px 8px rgba(0,0,0,.4);}
.line{position:absolute;white-space:pre;font-size:${SIZE}pt;}
.col{position:absolute;left:${MARGIN}pt;top:${MARGIN}pt;width:${COL_W}pt;height:1px;border-top:1px dashed #c00;}
.colr{position:absolute;left:${MARGIN + COL_W}pt;top:${MARGIN}pt;height:200pt;border-left:1px dashed #c00;}
</style>
<div class=page><div class=col></div><div class=colr></div>`;
let y = MARGIN;
for (const lr of lineReport) {
  // build positioned line: place each word, advance by word width + computed space
  let x = MARGIN;
  let spans = '';
  lr.wordsArr.forEach((w, k) => {
    spans += `<span style="position:absolute;left:${x.toFixed(2)}pt;">${w.replace(/&/g, '&amp;').replace(/</g, '&lt;')}</span>`;
    x += wordWidth(w) + (k < lr.wordsArr.length - 1 ? lr.spaceW : 0);
  });
  html += `<div class=line style="top:${y.toFixed(2)}pt;left:0;width:${PAGE_W}pt;">${spans}</div>`;
  y += singleLineHeight;
}
html += `</div>`;
writeFileSync(join(outDir, 'spikeA-overlay.html'), html);

// ── report ───────────────────────────────────────────────────────────────────
const fmt = (n) => (typeof n === 'number' ? n.toFixed(2) : n);
console.log('═══ SPIKE A — thesis test ═══');
console.log(`A4 column width: ${COL_W.toFixed(2)}pt  body size: ${SIZE}pt  space advance: ${SPACE_W.toFixed(3)}pt`);
console.log('\n(1) METRIC-CLONE IDENTITY  Liberation Sans vs Arial (per-word width @11pt, kern/liga off)');
if (arial) {
  for (const r of rows) console.log(`   ${r.w.padEnd(18)} clone=${r.wl.padStart(8)}  arial=${r.wa.padStart(8)}  Δ=${r.dPct}%`);
  console.log(`   → max Δ ${maxDeltaPct.toFixed(3)}%, avg Δ ${(sumDeltaPct / rows.length).toFixed(3)}%  ${maxDeltaPct < 0.5 ? '✓ advance-identical' : '✗ DIVERGENT'}`);
} else {
  console.log('   (Arial reference not present — clone-identity check skipped)');
}
console.log('\n(2) usWin LINE-HEIGHT');
console.log(`   upem=${lm.upem} USE_TYPO_METRICS=${lm.useTypo} winAsc=${lm.winAsc} winDesc=${lm.winDesc}`);
console.log(`   winSingle=${lm.winSingle.toFixed(3)}pt  typoSingle=${lm.typoSingle.toFixed(3)}pt  → single line height = ${singleLineHeight.toFixed(3)}pt`);
console.log('\n(3) GREEDY + JUSTIFICATION  (paragraph → lines)');
for (const l of lineReport) {
  console.log(`   L${l.idx}: ${String(l.words).padStart(2)}w  natural→final ${fmt(l.wordSum + (l.gaps) * SPACE_W).padStart(7)}→${fmt(l.finalWidth).padStart(7)}pt  ${l.mode.padEnd(16)}  "${l.text.slice(0, 48)}${l.text.length > 48 ? '…' : ''}"`);
}
console.log(`   lines: ${lineReport.length}  overflow lines: ${overflow.length}  ${overflow.length === 0 ? '✓ no overflow' : '✗ OVERFLOW'}`);
console.log(`   max inter-word shrink used: ${Number.isFinite(maxShrink) ? ((1 - maxShrink) * 100).toFixed(1) + '%' : 'none'} (cap 20%)`);
console.log(`\n(4) overlay artifact: spikes/render-core/out/spikeA-overlay.html (open in a browser; compare to Word)`);
console.log('\nROUNDING POLICY (fixed for the engine): advances summed in font units (integers),');
console.log('converted to pt once per measured span as (sumFu / upem) * sizePt; no per-glyph pt rounding.');
console.log(`\nVERDICT: ${arial && maxDeltaPct < 0.5 && overflow.length === 0 ? 'GO-eligible ✓ (clone-identical, no overflow, line-height computed)' : 'REVIEW'}`);
