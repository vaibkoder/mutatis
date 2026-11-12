/**
 * Table layout (feature 16 §4 hard-part #4; ADR-0024) — host-free, pure, DOM-free.
 *
 * Turns an `RcTable` into laid-out rows whose cells sit side-by-side at their
 * `tblGrid` column offsets, each cell's text wrapped by the SAME line breaker the
 * body uses (`layoutParagraph`). Output is consumed by the paginator, which
 * places whole rows, splits tall splittable rows across pages, and keeps
 * `cantSplit` rows together. Graceful degradation is the rule (ADR-0024 §2.2):
 *  - missing/zero `tblGrid` → even columns;
 *  - natural width wider than the content area → scaled to fit;
 *  - nested tables (>1 level, deferred) → that cell's inner table flattened to its
 *    cell paragraphs, linearly — content is rendered, never crashed or dropped.
 *
 * Coordinates: each cell line's run `xPt` is already shifted to the absolute
 * content-area x (column-left + cell padding); `relTopPt` is the line's top within
 * the row. The paginator assigns the row's absolute `topPt`.
 */
import type {
  RcTable, RcRow, RcBlock, LayoutLine, Rect,
} from '../types.js';
import type { Shaper } from '../layout/shape.js';
import type { FontResolver } from '../fonts/registry.js';
import { layoutParagraph } from '../layout/paragraph.js';

type Line = Omit<LayoutLine, 'topPt'>;

const CELL_PAD_X = 4; // pt, left/right cell inset
const CELL_PAD_Y = 2; // pt, top/bottom cell inset
const BORDER_COLOR = '#999999';
const BORDER_W = 0.5;

/** A cell line placed within a row (run x already absolute; relTop within the row). */
export interface CellLine {
  line: Line;
  relTopPt: number;
}

export interface TableRowLayout {
  cellLines: CellLine[];
  heightPt: number;
  cantSplit: boolean;
  /** Border/shading rects, y relative to the row top. */
  rects: Rect[];
}

export interface TableLayout {
  rows: TableRowLayout[];
}

/** Shift every run in a line by `dx` (cell-left + padding). Returns a fresh line. */
function shiftLine(line: Line, dx: number): Line {
  const runs = line.runs.map((r) => ({ ...r, xPt: r.xPt + dx }));
  const out: Line = { heightPt: line.heightPt, ascentPt: line.ascentPt, runs };
  if (line.changeBar) out.changeBar = true;
  return out;
}

/** Lay out a cell's blocks into a vertical stack of lines (nested tables degraded). */
function cellLines(blocks: RcBlock[], innerWidthPt: number, shaper: Shaper, fonts: FontResolver): Line[] {
  const lines: Line[] = [];
  for (const b of blocks) {
    if (b.kind === 'paragraph') {
      lines.push(...layoutParagraph(b.paragraph, innerWidthPt, shaper, fonts).lines);
    } else {
      // Nested table (deferred >1 level): degrade to its cell paragraphs, linearly.
      for (const row of b.table.rows) for (const c of row.cells) lines.push(...cellLines(c.blocks, innerWidthPt, shaper, fonts));
    }
  }
  return lines;
}

/** Resolve effective column widths (pt) for the table within the content width. */
function columnWidths(table: RcTable, contentWidthPt: number): number[] {
  let cols = table.gridColsPt.filter((w) => Number.isFinite(w) && w > 0);
  const maxCells = table.rows.reduce((m, r) => Math.max(m, r.cells.reduce((s, c) => s + Math.max(1, c.gridSpan), 0)), 0);
  if (cols.length === 0) {
    const n = Math.max(1, maxCells);
    cols = new Array(n).fill(contentWidthPt / n);
  }
  // Pad/truncate to cover the widest row's column count.
  while (cols.length < maxCells) cols.push(cols[cols.length - 1] ?? contentWidthPt / Math.max(1, maxCells));
  const total = cols.reduce((a, b) => a + b, 0);
  if (total > contentWidthPt && total > 0) {
    const scale = contentWidthPt / total;
    cols = cols.map((w) => w * scale);
  }
  return cols;
}

const prefix = (cols: number[]): number[] => {
  const out = [0];
  for (const c of cols) out.push(out[out.length - 1]! + c);
  return out;
};

function layoutRow(row: RcRow, cols: number[], left: number[], shaper: Shaper, fonts: FontResolver): TableRowLayout {
  const cellLinesOut: CellLine[] = [];
  const rects: Rect[] = [];
  let cellHeights = 0;
  let ci = 0; // running column index
  for (const cell of row.cells) {
    const span = Math.max(1, cell.gridSpan);
    const colLeft = left[Math.min(ci, left.length - 1)]!;
    const colRight = left[Math.min(ci + span, left.length - 1)]!;
    const outerW = Math.max(0, colRight - colLeft);
    const innerW = Math.max(1, outerW - 2 * CELL_PAD_X);
    const lines = cellLines(cell.blocks, innerW, shaper, fonts);

    let y = CELL_PAD_Y;
    for (const ln of lines) {
      cellLinesOut.push({ line: shiftLine(ln, colLeft + CELL_PAD_X), relTopPt: y });
      y += ln.heightPt;
    }
    const cellH = y + CELL_PAD_Y;
    cellHeights = Math.max(cellHeights, cellH);
    // Basic border around the cell (drawn at row height once known, below).
    rects.push({ xPt: colLeft, yPt: 0, widthPt: outerW, heightPt: 0, stroke: BORDER_COLOR, strokeWidthPt: BORDER_W });
    ci += span;
  }
  const heightPt = Math.max(cellHeights, 8); // never a zero-height row
  for (const r of rects) r.heightPt = heightPt; // borders span the full (max) row height
  return { cellLines: cellLinesOut, heightPt, cantSplit: row.cantSplit, rects };
}

/** Lay out a whole table into placeable rows. */
export function layoutTable(table: RcTable, contentWidthPt: number, shaper: Shaper, fonts: FontResolver): TableLayout {
  const cols = columnWidths(table, contentWidthPt);
  const left = prefix(cols);
  const rows = table.rows.map((row) => layoutRow(row, cols, left, shaper, fonts));
  return { rows };
}

/**
 * Split a row at a vertical offset (pt from row top): lines whose box fits within
 * `availPt` stay; the rest become a continuation row rebased to relTop 0. At least
 * one line always stays (never an infinite split). Returns [head, tail|null].
 */
export function splitRow(row: TableRowLayout, availPt: number): [TableRowLayout, TableRowLayout | null] {
  const head: CellLine[] = [];
  const tail: CellLine[] = [];
  for (const cl of row.cellLines) {
    const bottom = cl.relTopPt + cl.line.heightPt + CELL_PAD_Y;
    if (bottom <= availPt || head.length === 0) head.push(cl);
    else tail.push(cl);
  }
  if (tail.length === 0) return [row, null];
  const headH = Math.max(...head.map((c) => c.relTopPt + c.line.heightPt)) + CELL_PAD_Y;
  const tailBase = Math.min(...tail.map((c) => c.relTopPt)) - CELL_PAD_Y;
  const tailLines = tail.map((c) => ({ line: c.line, relTopPt: c.relTopPt - tailBase }));
  const tailH = Math.max(...tailLines.map((c) => c.relTopPt + c.line.heightPt)) + CELL_PAD_Y;
  const border = (h: number): Rect[] => row.rects.map((r) => ({ ...r, heightPt: h }));
  return [
    { cellLines: head, heightPt: headH, cantSplit: row.cantSplit, rects: border(headH) },
    { cellLines: tailLines, heightPt: tailH, cantSplit: row.cantSplit, rects: border(tailH) },
  ];
}
