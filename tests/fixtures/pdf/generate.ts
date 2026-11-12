/**
 * Deterministic PDF fixture generator — DEV-ONLY.
 *
 * Emits the synthetic, confidential-data-free `.pdf` (and one `.docx`) fixtures
 * the PDF-vs-PDF acceptance suite (P1–P13) runs against. No real or client
 * document ever enters the repo; every byte here is generated from this script.
 *
 * `pdf-lib` is a **dev-dependency only** — authoring throwaway test fixtures is
 * not authoring product output, so the product-side "no PDF authoring" rule
 * (ADR-0017) is untouched. This file is never imported by `src/**`.
 *
 * Run:  node --experimental-strip-types tests/fixtures/pdf/generate.ts
 * Output is byte-stable: all dates/producer/creator are pinned, so re-running
 * regenerates identical bytes (the A0 acceptance check).
 *
 * Honest coverage note (see the plan's A0 row + DONE_WITH_CONCERNS):
 *  - Glyph-level cases that pdf-lib's StandardFonts (WinAnsi) cannot encode —
 *    the f-ligatures U+FB00–06 (P10) and a ToUnicode-garbled stream (P9) — are
 *    NOT faithfully producible with a standard embedded font. Their guarantees
 *    are pinned by DIRECT unit tests on the A2 encoding guard and the A4
 *    normaliser (fed synthetic text-item arrays), which is where the logic
 *    lives. We still emit a soft-hyphen + dash fixture (those ARE WinAnsi) and a
 *    "near-garbled" page so the integration paths have something to chew on.
 */
import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { deflateSync } from 'node:zlib';

import {
  PDFDocument,
  StandardFonts,
  TextRenderingMode,
  rgb,
  type PDFFont,
  type PDFPage,
} from 'pdf-lib';

import PizZip from 'pizzip';

const OUT_DIR = dirname(fileURLToPath(import.meta.url));

// Pinned metadata so output bytes are reproducible across runs/machines.
const EPOCH = new Date('2026-01-01T00:00:00.000Z');
const PRODUCER = 'Mutatis fixture generator (A0)';

const PAGE_W = 595.28; // A4 portrait points
const PAGE_H = 841.89;
const MARGIN = 72;
const BODY_SIZE = 11;
const LINE_GAP = 6; // extra leading between lines

// ---------------------------------------------------------------------------
// Low-level helpers
// ---------------------------------------------------------------------------

interface Doc {
  pdf: PDFDocument;
  font: PDFFont;
  bold: PDFFont;
}

async function newDoc(): Promise<Doc> {
  const pdf = await PDFDocument.create();
  pdf.setProducer(PRODUCER);
  pdf.setCreator(PRODUCER);
  pdf.setCreationDate(EPOCH);
  pdf.setModificationDate(EPOCH);
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  return { pdf, font, bold };
}

/** Draw an array of body lines top-to-bottom on a fresh page. Returns the page. */
function drawLines(
  doc: Doc,
  lines: Array<{ text: string; bold?: boolean; colour?: [number, number, number] }>,
): PDFPage {
  const page = doc.pdf.addPage([PAGE_W, PAGE_H]);
  let y = PAGE_H - MARGIN;
  for (const line of lines) {
    const f = line.bold ? doc.bold : doc.font;
    const [r, g, b] = line.colour ?? [0, 0, 0];
    page.drawText(line.text, {
      x: MARGIN,
      y,
      size: BODY_SIZE,
      font: f,
      color: rgb(r, g, b),
    });
    y -= BODY_SIZE + LINE_GAP;
  }
  return page;
}

async function save(doc: Doc, name: string): Promise<void> {
  const bytes = await doc.pdf.save({ useObjectStreams: false });
  writeFileSync(join(OUT_DIR, name), bytes);

  console.log(`  wrote ${name} (${bytes.length} bytes)`);
}

/**
 * Minimal deterministic solid-colour PNG (no real raster content — a plain
 * coloured square stands in for a signature/seal/logo image XObject). Built by
 * hand so the only dependency is node:zlib. `tag` varies the pixel colour so a
 * "signature" image differs in bytes from a "logo" image (dedupe testing).
 */
function solidPng(size: number, tag: number): Uint8Array {
  const sig = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const r = (tag * 53) & 0xff;
  const g = (tag * 97) & 0xff;
  const b = (tag * 29) & 0xff;
  // Raw image data: each scanline prefixed with filter byte 0, RGB pixels.
  const raw = new Uint8Array(size * (1 + size * 3));
  for (let yy = 0; yy < size; yy++) {
    const rowStart = yy * (1 + size * 3);
    raw[rowStart] = 0;
    for (let xx = 0; xx < size; xx++) {
      const p = rowStart + 1 + xx * 3;
      raw[p] = r;
      raw[p + 1] = g;
      raw[p + 2] = b;
    }
  }
  const ihdr = new Uint8Array(13);
  const dv = new DataView(ihdr.buffer);
  dv.setUint32(0, size);
  dv.setUint32(4, size);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 2; // colour type RGB
  // 10,11,12 = compression/filter/interlace = 0
  const idat = deflateSync(raw, { level: 9 });
  const chunks = [chunk('IHDR', ihdr), chunk('IDAT', idat), chunk('IEND', new Uint8Array(0))];
  let total = sig.length;
  for (const c of chunks) total += c.length;
  const out = new Uint8Array(total);
  let off = 0;
  out.set(sig, off);
  off += sig.length;
  for (const c of chunks) {
    out.set(c, off);
    off += c.length;
  }
  return out;
}

function chunk(type: string, data: Uint8Array): Uint8Array {
  const typeBytes = Uint8Array.from([...type].map((ch) => ch.charCodeAt(0)));
  const body = new Uint8Array(typeBytes.length + data.length);
  body.set(typeBytes, 0);
  body.set(data, typeBytes.length);
  const out = new Uint8Array(4 + body.length + 4);
  const dv = new DataView(out.buffer);
  dv.setUint32(0, data.length);
  out.set(body, 4);
  dv.setUint32(4 + body.length, crc32(body));
  return out;
}

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(buf: Uint8Array): number {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]!)! & 0xff]! ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

// ---------------------------------------------------------------------------
// Shared base contract body (SG/UK, confidential-data-free)
// ---------------------------------------------------------------------------

/** The canonical single-column body. Variants below mutate one line each. */
function baseBody(): Array<{ text: string; bold?: boolean }> {
  return [
    { text: 'SALE AND PURCHASE AGREEMENT', bold: true },
    { text: '1. Definitions and Interpretation' },
    { text: '1.1 In this Agreement the following terms have the meanings given below.' },
    { text: '1.2 The Seller shall deliver the Goods to the Buyer on the Completion Date.' },
    { text: '2. Consideration' },
    { text: '2.1 The Buyer shall pay the Seller the sum of S$1,250,000 on Completion.' },
    { text: '2.2 Payment shall be made in accordance with Clause 2.1 by telegraphic transfer.' },
    { text: '3. Governing Law' },
    { text: '3.1 This Agreement is governed by the laws of Singapore.' },
    { text: '3.2 The parties submit to the jurisdiction of the Singapore courts under Clause 3.1.' },
  ];
}

// ---------------------------------------------------------------------------
// Fixture builders
// ---------------------------------------------------------------------------

async function singleColumnBase(): Promise<void> {
  const doc = await newDoc();
  drawLines(doc, baseBody());
  await save(doc, 'single-column.base.pdf');
}

/** P1 — one inserted paragraph mid-document. */
async function insertVariant(): Promise<void> {
  const doc = await newDoc();
  const body = baseBody();
  body.splice(4, 0, {
    text: '1.3 The Goods shall conform to the specification set out in Schedule 1.',
  });
  drawLines(doc, body);
  await save(doc, 'single-column.insert.pdf');
}

/** P2 — identical text, one paragraph bold + red (visual-only change). */
async function fontVariant(): Promise<void> {
  const doc = await newDoc();
  const body = baseBody().map((l, i) =>
    i === 3 ? { ...l, bold: true, colour: [0.8, 0, 0] as [number, number, number] } : l,
  );
  drawLines(doc, body);
  await save(doc, 'single-column.fontchange.pdf');
}

/** P4 — obligation language: shall → may in clause 1.2. */
async function shallMayVariant(): Promise<void> {
  const doc = await newDoc();
  const body = baseBody();
  body[3] = { text: '1.2 The Seller may deliver the Goods to the Buyer on the Completion Date.' };
  drawLines(doc, body);
  await save(doc, 'single-column.shall-may.pdf');
}

/** P6 — numeric amount change S$1,250,000 → S$1,500,000. */
async function amountVariant(): Promise<void> {
  const doc = await newDoc();
  const body = baseBody();
  body[5] = { text: '2.1 The Buyer shall pay the Seller the sum of S$1,500,000 on Completion.' };
  drawLines(doc, body);
  await save(doc, 'single-column.amount.pdf');
}

/**
 * P3 — renumbering cascade. A clause inserted at the top shifts every later
 * clause number AND its cross-references uniformly (Clause 2.1 → 3.1, etc.).
 */
async function renumberCascade(): Promise<void> {
  // Insert one clause near the top → every later clause number AND its
  // cross-references shift uniformly by +1. Four cross-refs shift (Clause
  // 2→3, 3→4, 4→5, 5→6), which is ≥ renumberThreshold (3), so the pipeline
  // collapses them into a single CascadeAnnotation (P3).
  const original = await newDoc();
  drawLines(original, [
    { text: 'DEED OF VARIATION', bold: true },
    { text: '1. Interpretation' },
    { text: '1.1 Terms have the meanings given in the principal Agreement.' },
    { text: '2. Payment' },
    { text: '2.1 The sum is payable under Clause 2 of this Deed.' },
    { text: '3. Notices' },
    { text: '3.1 Notices are given in accordance with Clause 3 of this Deed.' },
    { text: '4. Governing Law' },
    { text: '4.1 The deed is governed as provided in Clause 4 of this Deed.' },
    { text: '5. Variation' },
    { text: '5.1 The deed is varied only as permitted by Clause 5 of this Deed.' },
  ]);
  await save(original, 'renumber.base.pdf');

  const revised = await newDoc();
  drawLines(revised, [
    { text: 'DEED OF VARIATION', bold: true },
    { text: '1. Interpretation' },
    { text: '1.1 Terms have the meanings given in the principal Agreement.' },
    { text: '2. Recitals' },
    { text: '2.1 The parties have agreed to vary the principal Agreement.' },
    { text: '3. Payment' },
    { text: '3.1 The sum is payable under Clause 3 of this Deed.' },
    { text: '4. Notices' },
    { text: '4.1 Notices are given in accordance with Clause 4 of this Deed.' },
    { text: '5. Governing Law' },
    { text: '5.1 The deed is governed as provided in Clause 5 of this Deed.' },
    { text: '6. Variation' },
    { text: '6.1 The deed is varied only as permitted by Clause 6 of this Deed.' },
  ]);
  await save(revised, 'renumber.revised.pdf');
}

/** P5 — two-column page. Columns must be reading-order separated. */
async function twoColumn(): Promise<void> {
  const doc = await newDoc();
  const page = doc.pdf.addPage([PAGE_W, PAGE_H]);
  const colWidth = (PAGE_W - 2 * MARGIN - 24) / 2;
  const leftX = MARGIN;
  const rightX = MARGIN + colWidth + 24;
  const left = [
    'COLUMN A',
    'The first column carries the',
    'recitals of the agreement and',
    'sets out the background to the',
    'transaction between the parties.',
  ];
  const right = [
    'COLUMN B',
    'The second column carries the',
    'operative provisions and the',
    'obligations of each party under',
    'the terms of this agreement.',
  ];
  let y = PAGE_H - MARGIN;
  for (let i = 0; i < left.length; i++) {
    page.drawText(left[i]!, { x: leftX, y, size: BODY_SIZE, font: doc.font, color: rgb(0, 0, 0) });
    page.drawText(right[i]!, { x: rightX, y, size: BODY_SIZE, font: doc.font, color: rgb(0, 0, 0) });
    y -= BODY_SIZE + LINE_GAP;
  }
  await save(doc, 'two-column.pdf');
}

/** P7 — image-only / no extractable text (vector graphics only, zero text). */
async function imageOnly(): Promise<void> {
  const doc = await newDoc();
  const page = doc.pdf.addPage([PAGE_W, PAGE_H]);
  // Only vector rectangles — getTextContent() must return zero items.
  page.drawRectangle({ x: 100, y: 500, width: 300, height: 200, color: rgb(0.2, 0.3, 0.6) });
  page.drawRectangle({ x: 150, y: 300, width: 200, height: 120, color: rgb(0.6, 0.2, 0.2) });
  await save(doc, 'image-only.pdf');
}

/**
 * P7b — invisible OCR-style text layer: a full-page raster image with a thin
 * layer of text drawn in the Invisible render mode over it. Text extracts, but
 * coverage/uniformity flags it as a suspected OCR layer.
 */
async function ocrLayer(): Promise<void> {
  const doc = await newDoc();
  const page = doc.pdf.addPage([PAGE_W, PAGE_H]);
  const img = await doc.pdf.embedPng(solidPng(16, 7));
  page.drawImage(img, { x: 0, y: 0, width: PAGE_W, height: PAGE_H });
  const lines = [
    'SCANNED EXECUTION COPY',
    'This page is an image with an OCR text layer beneath it.',
    'The text is present but rendered invisibly over the raster.',
  ];
  let y = PAGE_H - MARGIN;
  for (const t of lines) {
    page.drawText(t, {
      x: MARGIN,
      y,
      size: BODY_SIZE,
      font: doc.font,
      color: rgb(0, 0, 0),
      // `renderMode` is supported at runtime but absent from the installed typings.
      renderMode: TextRenderingMode.Invisible,
    } as Parameters<typeof page.drawText>[1]);
    y -= BODY_SIZE + LINE_GAP;
  }
  await save(doc, 'ocr-layer.pdf');
}

/** P9 — near-garbled: many private-use / non-printable code points on a page. */
async function nearGarbled(): Promise<void> {
  const doc = await newDoc();
  const page = doc.pdf.addPage([PAGE_W, PAGE_H]);
  // WinAnsi can encode a run of box/control-ish glyphs; we lean on chars that
  // map to a high out-of-printable-range ratio when extracted. (See the honest
  // note at the top: the *true* ToUnicode-garbled guarantee is unit-tested.)
  const garble = '¤¦¬¶·¿×÷'.repeat(8);
  page.drawText(garble, { x: MARGIN, y: PAGE_H - MARGIN, size: BODY_SIZE, font: doc.font });
  await save(doc, 'near-garbled.pdf');
}

/** P10 (partial) — soft hyphens + en/em dashes (WinAnsi-encodable). */
async function softHyphenDash(): Promise<void> {
  const doc = await newDoc();
  drawLines(doc, [
    { text: 'SCHEDULE 1', bold: true },
    // U+00AD soft hyphen inside "indemnification"; en-dash range; em-dash clause.
    { text: 'The indem­nifying party shall — without limitation — cover losses.' },
    { text: 'The period is 30–60 days from the Completion Date.' },
  ]);
  await save(doc, 'soft-hyphen-dash.pdf');
}

/**
 * P11 — justified/kerned line fragmentation: each word drawn as a separate text
 * item at an irregular x, reproducing the sub-pixel item fragmentation PDF.js
 * surfaces for justified text. Reconstruction must yield one block per line.
 */
async function justifiedKerned(): Promise<void> {
  const doc = await newDoc();
  const page = doc.pdf.addPage([PAGE_W, PAGE_H]);
  const words =
    'The parties hereby agree that each obligation set out herein shall be performed promptly'.split(
      ' ',
    );
  let x = MARGIN;
  const y = PAGE_H - MARGIN;
  for (const w of words) {
    page.drawText(w, { x, y, size: BODY_SIZE, font: doc.font, color: rgb(0, 0, 0) });
    // Irregular advance (word width + jittered inter-word gap) → fragmented items.
    x += doc.font.widthOfTextAtSize(w, BODY_SIZE) + 4 + ((words.indexOf(w) % 3) * 1.5);
  }
  await save(doc, 'justified-kerned.pdf');
}

/**
 * P12 — fee/payment table, base + revised with one changed cell value. Grid
 * lines + per-cell text so A3's x-band detection can recover a coarse grid.
 */
async function feeTable(amount: string, name: string): Promise<void> {
  const doc = await newDoc();
  const page = doc.pdf.addPage([PAGE_W, PAGE_H]);
  const cols = [MARGIN, MARGIN + 200, MARGIN + 360];
  const rows = [
    ['Milestone', 'Due Date', 'Fee'],
    ['Signing', '01 Jan 2026', 'S$250,000'],
    ['Completion', '01 Mar 2026', amount],
    ['Final', '01 Jun 2026', 'S$100,000'],
  ];
  let y = PAGE_H - MARGIN;
  for (const row of rows) {
    for (let c = 0; c < row.length; c++) {
      page.drawText(row[c]!, { x: cols[c]!, y, size: BODY_SIZE, font: doc.font, color: rgb(0, 0, 0) });
    }
    y -= 28;
  }
  await save(doc, name);
}

/** P13 — a signature-image page (text + an embedded raster image XObject). */
async function signatureImage(): Promise<void> {
  const doc = await newDoc();
  const page = drawLines(doc, [
    { text: 'EXECUTION PAGE', bold: true },
    { text: 'Signed by the Seller in the presence of a witness.' },
  ]);
  const sig = await doc.pdf.embedPng(solidPng(24, 13));
  page.drawImage(sig, { x: MARGIN, y: PAGE_H - 260, width: 160, height: 60 });
  await save(doc, 'signature-image.pdf');
}

/** P13 — repeated letterhead logo on every page (must dedupe, not flag each). */
async function repeatedLogo(): Promise<void> {
  const doc = await newDoc();
  const logo = await doc.pdf.embedPng(solidPng(20, 3));
  for (let p = 0; p < 3; p++) {
    const page = doc.pdf.addPage([PAGE_W, PAGE_H]);
    page.drawImage(logo, { x: MARGIN, y: PAGE_H - 120, width: 80, height: 40 });
    page.drawText(`Page ${p + 1} of the agreement body text.`, {
      x: MARGIN,
      y: PAGE_H - 180,
      size: BODY_SIZE,
      font: doc.font,
      color: rgb(0, 0, 0),
    });
  }
  await save(doc, 'repeated-logo.pdf');
}

/**
 * Word+PDF pair (Phase E) — a PDF and a minimal .docx carrying the SAME body
 * text verbatim, so anchor matching is exercised on identical content.
 */
async function wordPdfPair(): Promise<void> {
  const lines = [
    'SUPPLY AGREEMENT',
    'The Supplier shall provide the Services with reasonable skill and care.',
    'The Customer shall pay the Charges within 30 days of each invoice.',
    'This Agreement is governed by the laws of England and Wales.',
  ];
  const doc = await newDoc();
  drawLines(doc, lines.map((text) => ({ text })));
  await save(doc, 'word-pdf-pair.pdf');
  writeDocx(lines, 'word-pdf-pair.docx');
}

/** Minimal valid .docx (one <w:p> per line) for the Word+PDF pair. */
function writeDocx(lines: string[], name: string): void {
  const esc = (s: string) =>
    s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const paras = lines
    .map((l) => `<w:p><w:r><w:t xml:space="preserve">${esc(l)}</w:t></w:r></w:p>`)
    .join('');
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body>${paras}<w:sectPr/></w:body></w:document>`;
  const contentTypes =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
    `</Types>`;
  const rels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
    `</Relationships>`;
  const zip = new PizZip();
  zip.file('[Content_Types].xml', contentTypes);
  zip.file('_rels/.rels', rels);
  zip.file('word/document.xml', documentXml);
  // Pinned mtime via PizZip date option keeps the archive byte-stable.
  // PizZip's overloads do not narrow on a literal options object, so the call
  // resolves to `never` without help. The runtime behaviour is unchanged.
  const bytes: Uint8Array = zip.generate(
    { type: 'uint8array', compression: 'DEFLATE', date: EPOCH } as Parameters<typeof zip.generate>[0],
  );
  writeFileSync(join(OUT_DIR, name), bytes);

  console.log(`  wrote ${name} (${bytes.length} bytes)`);
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main(): Promise<void> {
  mkdirSync(OUT_DIR, { recursive: true });

  console.log('Generating PDF fixtures →', OUT_DIR);
  await singleColumnBase();
  await insertVariant();
  await fontVariant();
  await shallMayVariant();
  await amountVariant();
  await renumberCascade();
  await twoColumn();
  await imageOnly();
  await ocrLayer();
  await nearGarbled();
  await softHyphenDash();
  await justifiedKerned();
  await feeTable('S$500,000', 'fee-table.base.pdf');
  await feeTable('S$650,000', 'fee-table.revised.pdf');
  await signatureImage();
  await repeatedLogo();
  await wordPdfPair();

  console.log('Done.');
}

void main();
