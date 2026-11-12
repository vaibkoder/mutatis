// Spike C — @pdfme/pdf-lib + fontkit SUBSET font emission (THROWAWAY).
//
// Proves the PDF backend: subset-embed an OFL clone, paint positioned runs
// (vector) across 2 pages from harfbuzz advances, superscript a footnote ref,
// and verify the output is (a) a valid multi-page vector PDF, (b) selectable
// text (re-extracted with the project's own pdfjs-dist), (c) font embedded as a
// SUBSET (subset BaseFont tag + FontFile2 stream much smaller than the original),
// (d) sane file size.
//
// Run: node spikes/render-core/spikeC-pdf.mjs
import { readFileSync, writeFileSync, statSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { PDFDocument, rgb } from '@pdfme/pdf-lib';
import * as fontkit from 'fontkit';
import { loadFont, shapeRun } from './hb-util.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const asset = (f) => join(HERE, 'assets', f);
const outPath = join(HERE, 'out', 'spikeC.pdf');

const SIZE = 11, SUP = 7; // body + superscript size
const fontBytes = readFileSync(asset('LiberationSans-Regular.ttf'));
const origFontSize = fontBytes.length;
const { font: hbFont, upem } = loadFont(asset('LiberationSans-Regular.ttf'));
const advPt = (text, size) => (shapeRun(hbFont, text).widthFu / upem) * size;

const A4 = [595.28, 841.89]; // pt
const MARGIN = 72;

const doc = await PDFDocument.create();
doc.registerFontkit(fontkit);
const embedded = await doc.embedFont(fontBytes, { subset: true });

// Two pages of positioned words painted from harfbuzz advances (vector text).
const PARAS = [
  'This Agreement is made between the parties and shall be governed by the laws of Singapore.',
  'The Borrower shall pay all amounts due under Clause 8.2 on each Interest Payment Date.',
  'Confidential Information must not be disclosed save as permitted by this Agreement.',
];
for (let pg = 0; pg < 2; pg++) {
  const page = doc.addPage(A4);
  let y = A4[1] - MARGIN;
  for (const para of PARAS) {
    let x = MARGIN;
    const words = para.split(' ');
    words.forEach((w, k) => {
      page.drawText(w, { x, y, size: SIZE, font: embedded, color: rgb(0.1, 0.1, 0.1) });
      x += advPt(w, SIZE) + advPt(' ', SIZE);
      // superscript a footnote marker after the first word of the first para on page 1
      if (pg === 0 && para === PARAS[0] && k === 0) {
        page.drawText('1', { x: x - advPt(' ', SIZE), y: y + SIZE * 0.33, size: SUP, font: embedded, color: rgb(0.1, 0.1, 0.1) });
        x += advPt('1', SUP);
      }
    });
    y -= 16;
  }
  // footnote at bottom of page 1
  if (pg === 0) {
    const fy = MARGIN + 24;
    page.drawLine({ start: { x: MARGIN, y: fy + 10 }, end: { x: MARGIN + 144, y: fy + 10 }, thickness: 0.5, color: rgb(0.4, 0.4, 0.4) });
    page.drawText('1', { x: MARGIN, y: fy + 4, size: SUP, font: embedded });
    page.drawText('See the recitals above for defined terms.', { x: MARGIN + 8, y: fy, size: 9, font: embedded });
  }
}
doc.setTitle('RenderCore Spike C');
// Object streams OFF so font dictionaries (/BaseFont, /FontFile2, /ToUnicode)
// are inspectable in the spike. The engine keeps compression on (smaller files).
const bytes = await doc.save({ useObjectStreams: false });
writeFileSync(outPath, bytes);
const sizeKB = (statSync(outPath).size / 1024).toFixed(1);

// ── verify ───────────────────────────────────────────────────────────────────
const pdfText = Buffer.from(bytes).toString('latin1');
// pdf-lib/fontkit subsets are named "<Family>-<NNNN>" (composite Type0) or the
// classic "ABCDEF+<Family>" tag — accept either as a subset signature.
const baseFont = pdfText.match(/\/BaseFont\s*\/([A-Za-z0-9]+(?:-\d+|\+[A-Za-z0-9]+))/);
const isSubset = !!baseFont && (/-\d+$/.test(baseFont[1]) || /^[A-Z]{6}\+/.test(baseFont[1]));
const isType0 = /\/Subtype\s*\/Type0/.test(pdfText) && /\/Subtype\s*\/CIDFontType2/.test(pdfText);
const hasToUnicode = /\/ToUnicode/.test(pdfText);
const hasFontFile2 = /\/FontFile2/.test(pdfText);
// Find the FontFile2 stream object and read its /Length (the embedded subset size).
let subsetBytes = NaN;
const ffRef = pdfText.match(/\/FontFile2\s+(\d+)\s+0\s+R/);
if (ffRef) {
  const obj = new RegExp(`\\b${ffRef[1]}\\s+0\\s+obj[\\s\\S]{0,400}?/Length\\s+(\\d+)`);
  const m = pdfText.match(obj);
  if (m) subsetBytes = parseInt(m[1], 10);
}

// reparse for page count
const re = await PDFDocument.load(bytes);
const pageCount = re.getPageCount();

// selectable-text proof via the project's own pdfjs-dist (legacy build, Node)
let extracted = '';
try {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
  const task = pdfjs.getDocument({ data: new Uint8Array(bytes), useWorkerFetch: false, isEvalSupported: false, disableFontFace: true });
  const pdf = await task.promise;
  const p1 = await pdf.getPage(1);
  const tc = await p1.getTextContent();
  extracted = tc.items.map((it) => it.str).join(' ').replace(/\s+/g, ' ').trim();
} catch (e) {
  extracted = `(pdfjs extraction unavailable: ${e.message})`;
}

console.log('═══ SPIKE C — pdf-lib + fontkit subset emission ═══');
console.log(`output: spikes/render-core/out/spikeC.pdf  size=${sizeKB} KB  pages=${pageCount}`);
console.log(`vector text: drawText positioned runs (not raster) ✓`);
console.log(`font subset: BaseFont=${baseFont ? baseFont[1] : 'NONE'}  composite(Type0/CIDFontType2)=${isType0}  ${isSubset ? '✓ (subset-named)' : '✗'}`);
console.log(`embedded TrueType: FontFile2=${hasFontFile2} ToUnicode=${hasToUnicode} (ToUnicode ⇒ selectable/extractable)`);
console.log(`subset size: ${Number.isFinite(subsetBytes) ? subsetBytes : '?'} bytes  vs original ${origFontSize} bytes  ${Number.isFinite(subsetBytes) && subsetBytes < origFontSize * 0.5 ? '✓ subsetted (≪ original)' : '⚠ check'}`);
console.log(`selectable text (pdfjs getTextContent, page 1):`);
console.log(`   "${extracted.slice(0, 120)}${extracted.length > 120 ? '…' : ''}"`);
const textOk = /Agreement|Singapore|parties/.test(extracted);
console.log(`   ${textOk ? '✓ real words extracted (text is selectable)' : '⚠ extraction did not yield expected words'}`);
const ok = pageCount === 2 && isSubset && hasToUnicode && hasFontFile2 && Number.isFinite(subsetBytes) && subsetBytes < origFontSize * 0.5 && textOk;
console.log(`\nNOTE: WebKit/Safari + Acrobat/Preview open is a user-side check (Phase 2 full cross-browser run).`);
console.log(`VERDICT: ${ok ? 'GO-eligible ✓ (valid 2-page vector PDF, subset-embedded font, selectable text)' : 'REVIEW'}`);
