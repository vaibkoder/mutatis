/**
 * PDF backend (feature 16 §2.1 Output B; ADR-0024) — host-free, DOM-free.
 *
 * Paints the ONE shared `LayoutResult` as a programmatic **vector** PDF with
 * `@pdfme/pdf-lib` + `fontkit` subset embedding (Spike C proved the path):
 *   - positioned runs drawn as `drawText` (vector glyphs, not raster);
 *   - tracked-change decorations drawn as vector ops (underline / strike /
 *     change bar / highlight rects / move-anchor glyphs);
 *   - footnote bodies + separator at the bottom of the page they belong to;
 *   - superscript footnote refs;
 *   - selectable text (ToUnicode) + subset-embedded fonts (FontFile2 ≪ original);
 *   - object-stream compression ON (smaller files; Spike C disabled it only to
 *     inspect the font dicts).
 *
 * Coordinate mapping: the layout result is in **content-area** coords with y
 * measured DOWN from the content-area top (the same space the preview paints).
 * PDF user space has its origin bottom-left with y UP, so every y is flipped
 * through the page geometry here. Because both painters consume the identical
 * `LayoutResult`, the PDF paginates identically to the preview by construction
 * (acceptance criterion 7).
 *
 * No DOM, no `fetch`/XHR — font BYTES arrive via the injected `FontResolver`
 * (zero network egress; ADR-0001/0003). Pure transform of bytes in → bytes out.
 */
import { PDFDocument, rgb, type PDFFont, type PDFPage, type RGB } from '@pdfme/pdf-lib';
import * as fontkit from 'fontkit';
import type {
  LayoutResult, LayoutPage, LayoutLine, PositionedRun, Rect, PageGeometry,
} from '../types.js';
import { PAGE_INK } from '../types.js';
import type { FontResolver, ResolvedFace } from '../fonts/registry.js';

export interface PdfMetadata {
  title?: string;
  author?: string;
  subject?: string;
}

/** Superscript/subscript draw scale + baseline shift fractions (of run size). */
const SUP_SCALE = 0.8;
const SUP_RAISE = 0.33;
const SUB_DROP = 0.12;
/** Change-bar geometry (left of the content area). */
const CHANGE_BAR_GAP_PT = 6;
const CHANGE_BAR_WIDTH_PT = 2;
const CHANGE_BAR_COLOR = '#808080';
/** Footnote separator rule width (fraction of content width) + colour. */
const FN_SEP_FRACTION = 0.32;
const FN_SEP_COLOR = '#666666';

interface Hsla { r: number; g: number; b: number; a: number; }

/** Parse `#rgb` / `#rrggbb` / `#rrggbbaa` → normalised channels (+ alpha). */
function parseHex(hex: string): Hsla {
  let h = hex.replace('#', '').trim();
  if (h.length === 3) h = h.split('').map((c) => c + c).join('');
  const r = parseInt(h.slice(0, 2), 16) / 255;
  const g = parseInt(h.slice(2, 4), 16) / 255;
  const b = parseInt(h.slice(4, 6), 16) / 255;
  const a = h.length >= 8 ? parseInt(h.slice(6, 8), 16) / 255 : 1;
  return {
    r: Number.isFinite(r) ? r : 0,
    g: Number.isFinite(g) ? g : 0,
    b: Number.isFinite(b) ? b : 0,
    a: Number.isFinite(a) ? a : 1,
  };
}

/** Hex → pdf-lib RGB, alpha pre-blended over white (pdf-lib fills are opaque). */
function pdfColor(hex: string): RGB {
  const { r, g, b, a } = parseHex(hex);
  return rgb(r * a + (1 - a), g * a + (1 - a), b * a + (1 - a));
}

const faceKey = (f: ResolvedFace): string => `${f.family}|${f.bold ? 'b' : ''}${f.italic ? 'i' : ''}`;

/** Collect the distinct resolved faces referenced by every run on every page. */
function collectFaces(layout: LayoutResult, fonts: FontResolver): Map<string, ResolvedFace> {
  const faces = new Map<string, ResolvedFace>();
  const scanLine = (line: LayoutLine): void => {
    for (const run of line.runs) {
      const face = fonts.resolve(run.rpr);
      faces.set(faceKey(face), face);
    }
  };
  for (const page of layout.pages) {
    for (const line of page.bodyLines) scanLine(line);
    if (page.footnotes) for (const line of page.footnotes.lines) scanLine(line);
  }
  return faces;
}

/** PDF y of a run's baseline within `line` (content-area coords → PDF user space). */
function baselineY(g: PageGeometry, line: LayoutLine): number {
  const contentTopY = g.heightPt - g.marginTopPt;
  return contentTopY - line.topPt - line.ascentPt;
}

function isSuperscript(run: PositionedRun): boolean {
  return run.rpr.superscript || run.footnoteId !== undefined;
}

/** Paint one line's runs + per-run decorations onto a PDF page. */
function paintLine(
  page: PDFPage,
  g: PageGeometry,
  line: LayoutLine,
  fontOf: (run: PositionedRun) => PDFFont,
): void {
  const baseY = baselineY(g, line);
  const xAt = (xPt: number): number => g.marginLeftPt + xPt;

  // Change bar (left margin) for changed lines.
  if (line.changeBar) {
    const topY = (g.heightPt - g.marginTopPt) - line.topPt;
    page.drawLine({
      start: { x: g.marginLeftPt - CHANGE_BAR_GAP_PT, y: topY },
      end: { x: g.marginLeftPt - CHANGE_BAR_GAP_PT, y: topY - line.heightPt },
      thickness: CHANGE_BAR_WIDTH_PT,
      color: pdfColor(CHANGE_BAR_COLOR),
    });
  }

  for (const run of line.runs) {
    if (run.text.length === 0) continue;
    const size = run.rpr.sizePt;
    const sup = isSuperscript(run);
    const sub = run.rpr.subscript;
    const drawSize = sup || sub ? size * SUP_SCALE : size;
    const yShift = sup ? size * SUP_RAISE : sub ? -size * SUB_DROP : 0;
    const y = baseY + yShift;
    const x = xAt(run.xPt);
    // Text colour: tracked-change runs take their decoration colour; everything
    // else paints in the page's own ink (PAGE_INK) — the page is paper, so it
    // never inherits a host theme foreground.
    const textColor = run.decorations?.[0]?.color ?? PAGE_INK;
    page.drawText(run.text, { x, y, size: drawSize, font: fontOf(run), color: pdfColor(textColor) });

    // Inline anchor glyphs stay per-run; underline/strike are drawn once per
    // merged segment below (continuous across word gaps). 'highlight' is a page
    // Rect (behind text); 'changeBar' is line-level.
    for (const dec of run.decorations ?? []) {
      if ((dec.kind === 'moveAnchor' || dec.kind === 'changeAnchor') && dec.glyph) {
        const x2 = x + run.widthPt;
        page.drawText(dec.glyph, { x: x2, y: y + size * SUP_RAISE, size: size * SUP_SCALE, font: fontOf(run), color: pdfColor(dec.color) });
      }
    }
  }

  // Continuous underline/strike rules (merged spans — conventional redline look).
  for (const seg of line.decoSegs ?? []) {
    const x1 = xAt(seg.x1Pt);
    const x2 = xAt(seg.x2Pt);
    const thick = Math.max(0.5, seg.sizePt * 0.055);
    const segY = seg.kind === 'insUnderline' ? baseY - seg.sizePt * 0.12 : baseY + seg.sizePt * 0.28;
    page.drawLine({ start: { x: x1, y: segY }, end: { x: x2, y: segY }, thickness: thick, color: pdfColor(seg.color) });
  }
}

/** Paint the page's filled rects (insertion highlights, table shading) behind text. */
function paintRects(page: PDFPage, g: PageGeometry, rects: Rect[]): void {
  const contentTopY = g.heightPt - g.marginTopPt;
  for (const r of rects) {
    const opts: Parameters<PDFPage['drawRectangle']>[0] = {
      x: g.marginLeftPt + r.xPt,
      y: contentTopY - r.yPt - r.heightPt,
      width: r.widthPt,
      height: r.heightPt,
    };
    if (r.fill) opts.color = pdfColor(r.fill);
    if (r.stroke) { opts.borderColor = pdfColor(r.stroke); opts.borderWidth = r.strokeWidthPt ?? 0.5; }
    page.drawRectangle(opts);
  }
}

function paintPage(
  pdf: PDFPage,
  page: LayoutPage,
  fontOf: (run: PositionedRun) => PDFFont,
): void {
  const g = page.geometry;
  paintRects(pdf, g, page.rects);
  for (const line of page.bodyLines) paintLine(pdf, g, line, fontOf);

  if (page.footnotes) {
    const contentTopY = g.heightPt - g.marginTopPt;
    const sepY = contentTopY - page.footnotes.separatorTopPt;
    pdf.drawLine({
      start: { x: g.marginLeftPt, y: sepY },
      end: { x: g.marginLeftPt + (g.widthPt - g.marginLeftPt - g.marginRightPt) * FN_SEP_FRACTION, y: sepY },
      thickness: 0.5,
      color: pdfColor(FN_SEP_COLOR),
    });
    for (const line of page.footnotes.lines) paintLine(pdf, g, line, fontOf);
  }
}

/**
 * Render the shared layout result to a vector PDF (subset-embedded fonts,
 * selectable text). Returns the serialised PDF bytes. Page count == layout
 * pages (parity with the preview by construction).
 */
export async function renderPdf(
  layout: LayoutResult,
  fonts: FontResolver,
  meta: PdfMetadata = {},
): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  doc.registerFontkit(fontkit);

  // Embed each distinct face once, as a SUBSET (Spike C policy).
  const faces = collectFaces(layout, fonts);
  const embedded = new Map<string, PDFFont>();
  for (const [key, face] of faces) {
    embedded.set(key, await doc.embedFont(face.bytes, { subset: true }));
  }
  const fallback = embedded.values().next().value as PDFFont | undefined;
  const fontOf = (run: PositionedRun): PDFFont => {
    const face = fonts.resolve(run.rpr);
    return embedded.get(faceKey(face)) ?? fallback!;
  };

  for (const page of layout.pages) {
    const pdfPage = doc.addPage([page.geometry.widthPt, page.geometry.heightPt]);
    paintPage(pdfPage, page, fontOf);
  }

  doc.setTitle(meta.title ?? 'Mutatis rendered redline');
  doc.setCreator('Mutatis RenderCore');
  doc.setProducer('Mutatis RenderCore (@pdfme/pdf-lib)');
  if (meta.author) doc.setAuthor(meta.author);
  if (meta.subject) doc.setSubject(meta.subject);

  // Object streams ON → smaller files (fixed engine policy, gate finding 6).
  return doc.save();
}
