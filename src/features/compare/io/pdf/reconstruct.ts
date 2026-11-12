/**
 * Geometry paragraph reconstruction — turn positioned
 * glyph runs into logical paragraphs in reading order.
 *
 * PDFs store positioned glyphs, NOT paragraphs (ADR-0017 reliability note), so
 * this is a heuristic geometry pass:
 *   1. drop rotated items (watermarks/diagonal stamps);
 *   2. detect column gutters by merging every item's horizontal span and finding
 *      interior gaps empty across the WHOLE page → partition items into columns
 *      BEFORE any y-clustering, so a left-column line and a right-column line at
 *      the same y can never merge (no cross-column token leak — P5);
 *   3. within each column, cluster items into lines by y (primary), using hasEOL
 *      only as a secondary tie-breaker;
 *   4. join line items (kerned fragments rejoin, real spaces stay — P11);
 *   5. group lines into paragraphs by vertical gap;
 *   6. tag headings (per-page median font size + ALL-CAPS) and confidence.
 *
 * Reading order: page → column (left→right) → paragraph (top→bottom). Multi
 * column pages flag every paragraph low-confidence (a flagged region, never
 * silently mis-ordered). All thresholds come from reconstruct.config.ts. Pure.
 */
import type { PdfExtraction, PdfPageText } from './extract.js';
import type { PdfTextItem } from './pdfjs-types.js';
import { RECONSTRUCT_CONFIG, type ReconstructConfig } from './reconstruct.config.js';

/** One reconstructed logical paragraph, in reading order. */
export interface ReconstructedParagraph {
  text: string;
  pageNumber: number;
  /** Column band index on its page (0 = leftmost / single column). */
  column: number;
  /** y of the paragraph's first (top) line, for stable ordering/debug. */
  yTop: number;
  /** Median font size of the paragraph's items. */
  fontSize: number;
  /** Heuristic heading flag (per-page median font size + ALL-CAPS). */
  isHeading: boolean;
  /** Reconstruction confidence is reduced (e.g. multi-column page). */
  lowConfidence: boolean;
}

/** A positioned item with its derived geometry. */
interface Geo {
  item: PdfTextItem;
  x0: number;
  x1: number;
  y: number;
  fontSize: number;
}

/** Reconstruct every trustworthy page of an extraction into paragraphs. */
export function reconstructDocument(
  extraction: PdfExtraction,
  config: ReconstructConfig = RECONSTRUCT_CONFIG,
): ReconstructedParagraph[] {
  const out: ReconstructedParagraph[] = [];
  for (const page of extraction.pages) {
    out.push(...reconstructPage(page, config));
  }
  return out;
}

/** Reconstruct a single page into ordered paragraphs. */
export function reconstructPage(
  page: PdfPageText,
  config: ReconstructConfig = RECONSTRUCT_CONFIG,
): ReconstructedParagraph[] {
  // 1. Keep only non-empty, non-rotated items with geometry.
  const geos: Geo[] = [];
  for (const item of page.items) {
    if (item.str.trim() === '') continue;
    if (rotationDegrees(item.transform) > config.ROTATION_MAX_DEGREES) continue;
    const fontSize = itemFontSize(item.transform);
    const x0 = item.transform[4] ?? 0;
    const y = item.transform[5] ?? 0;
    geos.push({ item, x0, x1: x0 + (item.width || 0), y, fontSize });
  }
  if (geos.length === 0) return [];

  const median = medianFontSize(geos, config);

  // 2. Partition into columns BEFORE y-clustering (prevents cross-column leak).
  const boundaries = detectColumnGutters(geos, median, config);
  const multiColumn = boundaries.length > 0;
  const columns = partitionColumns(geos, boundaries);

  const paragraphs: ReconstructedParagraph[] = [];
  columns.forEach((colGeos, colIndex) => {
    const lines = clusterLines(colGeos, median, config);
    const paras = groupParagraphs(lines, median, config);
    for (const para of paras) {
      const text = para.lines.map((ln) => ln.text).join(' ').replace(/\s+/g, ' ').trim();
      if (text === '') continue;
      paragraphs.push({
        text,
        pageNumber: page.pageNumber,
        column: colIndex,
        yTop: para.yTop,
        fontSize: para.fontSize,
        isHeading: isHeadingLine(text, para.fontSize, median, config),
        lowConfidence: multiColumn,
      });
    }
  });
  return paragraphs;
}

// ---------------------------------------------------------------------------
// Geometry helpers
// ---------------------------------------------------------------------------

/** Font size = length of the text-space x-basis vector [a, b]. */
function itemFontSize(t: number[]): number {
  return Math.hypot(t[0] ?? 0, t[1] ?? 0);
}

/** Rotation of the text baseline from horizontal, in degrees (absolute). */
function rotationDegrees(t: number[]): number {
  const deg = Math.abs((Math.atan2(t[1] ?? 0, t[0] ?? 1) * 180) / Math.PI);
  return deg > 90 ? 180 - deg : deg; // fold 180° (upside-down) to 0 distance
}

function medianFontSize(geos: Geo[], config: ReconstructConfig): number {
  const sizes = geos.map((g) => g.fontSize).filter((s) => s > 0).sort((a, b) => a - b);
  if (sizes.length === 0) return config.DEFAULT_FONT_SIZE;
  const mid = Math.floor(sizes.length / 2);
  return sizes.length % 2 === 0 ? (sizes[mid - 1]! + sizes[mid]!) / 2 : sizes[mid]!;
}

// ---------------------------------------------------------------------------
// Column detection (merge all x-spans → interior empty gutters)
// ---------------------------------------------------------------------------

/** x positions that split the page into columns (sorted ascending). */
function detectColumnGutters(geos: Geo[], median: number, config: ReconstructConfig): number[] {
  const minGutter = config.COLUMN_GUTTER_MIN_FACTOR * median;
  const spans = geos.map((g) => [g.x0, g.x1] as [number, number]).sort((a, b) => a[0] - b[0]);
  // Merge overlapping spans into coverage intervals.
  const merged: Array<[number, number]> = [];
  for (const [s, e] of spans) {
    const last = merged[merged.length - 1];
    if (last && s <= last[1]) {
      last[1] = Math.max(last[1], e);
    } else {
      merged.push([s, e]);
    }
  }
  // Interior gaps between coverage intervals wider than minGutter are gutters.
  const boundaries: number[] = [];
  for (let i = 1; i < merged.length; i++) {
    const gap = merged[i]![0] - merged[i - 1]![1];
    if (gap > minGutter) boundaries.push((merged[i - 1]![1] + merged[i]![0]) / 2);
  }
  return boundaries;
}

/** Split items into columns by the gutter boundaries (left→right order). */
function partitionColumns(geos: Geo[], boundaries: number[]): Geo[][] {
  if (boundaries.length === 0) return [geos];
  const columns: Geo[][] = Array.from({ length: boundaries.length + 1 }, () => []);
  for (const g of geos) {
    // Column index = number of boundaries the item's start is to the right of.
    let col = 0;
    while (col < boundaries.length && g.x0 >= boundaries[col]!) col++;
    columns[col]!.push(g);
  }
  return columns;
}

// ---------------------------------------------------------------------------
// Line clustering + paragraph grouping
// ---------------------------------------------------------------------------

interface Line {
  text: string;
  yTop: number;
  fontSize: number;
}

/** Cluster a column's items into lines by y (descending = top to bottom). */
function clusterLines(geos: Geo[], median: number, config: ReconstructConfig): Line[] {
  const yTol = config.LINE_Y_TOLERANCE_FACTOR * median;
  const sorted = [...geos].sort((a, b) => b.y - a.y); // top first (PDF y grows upward)

  const lines: Line[] = [];
  let current: Geo[] = [];
  let currentY = Number.NaN;
  const flush = (): void => {
    if (current.length === 0) return;
    lines.push(buildLine(current, config));
    current = [];
  };
  for (const g of sorted) {
    if (current.length === 0 || Math.abs(g.y - currentY) <= yTol) {
      current.push(g);
      currentY = current.length === 1 ? g.y : currentY; // anchor on the first item's y
    } else {
      flush();
      current.push(g);
      currentY = g.y;
    }
  }
  flush();
  return lines;
}

/** Build one line's text from its items: x-sort, rejoin kerned fragments. */
function buildLine(items: Geo[], config: ReconstructConfig): Line {
  const sorted = [...items].sort((a, b) => a.x0 - b.x0);
  const fontSize = sorted[0]!.fontSize || config.DEFAULT_FONT_SIZE;
  const spaceGap = config.INTRA_LINE_SPACE_FACTOR * fontSize;
  let text = '';
  let prevX1 = Number.NaN;
  for (const g of sorted) {
    const s = g.item.str;
    if (text === '') {
      text = s;
    } else {
      const gap = g.x0 - prevX1;
      const alreadySpaced = /\s$/.test(text) || /^\s/.test(s);
      text += gap > spaceGap && !alreadySpaced ? ` ${s}` : s;
    }
    prevX1 = g.x1;
  }
  return { text, yTop: Math.max(...sorted.map((g) => g.y)), fontSize };
}

interface Para {
  lines: Line[];
  yTop: number;
  fontSize: number;
}

/** Group consecutive lines into paragraphs by vertical gap. */
function groupParagraphs(lines: Line[], median: number, config: ReconstructConfig): Para[] {
  const lineHeight = median > 0 ? median : config.DEFAULT_FONT_SIZE;
  const paraGap = config.PARAGRAPH_GAP_FACTOR * lineHeight;
  const paras: Para[] = [];
  let current: Line[] = [];
  const flush = (): void => {
    if (current.length === 0) return;
    paras.push({
      lines: current,
      yTop: current[0]!.yTop,
      fontSize: medianOf(current.map((l) => l.fontSize), config),
    });
    current = [];
  };
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i]!;
    if (current.length === 0) {
      current.push(line);
      continue;
    }
    const prev = current[current.length - 1]!;
    const gap = prev.yTop - line.yTop; // positive going down the page
    // An outsized vertical gap starts a new paragraph. (Heading detection is a
    // post-hoc per-paragraph flag — see isHeading — not a split signal here.)
    if (gap > paraGap) {
      flush();
    }
    current.push(line);
  }
  flush();
  return paras;
}

function medianOf(values: number[], config: ReconstructConfig): number {
  const s = values.filter((v) => v > 0).sort((a, b) => a - b);
  if (s.length === 0) return config.DEFAULT_FONT_SIZE;
  const mid = Math.floor(s.length / 2);
  return s.length % 2 === 0 ? (s[mid - 1]! + s[mid]!) / 2 : s[mid]!;
}

// ---------------------------------------------------------------------------
// Heading heuristic
// ---------------------------------------------------------------------------

function isHeadingLine(
  text: string,
  fontSize: number,
  median: number,
  config: ReconstructConfig,
): boolean {
  if (median > 0 && fontSize > median * config.HEADING_FONT_RATIO) return true;
  const letters = text.replace(/[^A-Za-z]/g, '');
  const words = text.trim().split(/\s+/).length;
  const allCaps = letters.length >= 2 && letters === letters.toUpperCase();
  return allCaps && words <= config.HEADING_MAX_WORDS;
}
