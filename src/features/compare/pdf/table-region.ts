/**
 * PDF approximate table-region detection + diff.
 *
 * A PDF has no `<w:tbl>` structure — a table is just text positioned in aligned
 * x-bands. This pass reconstructs a grid from raw positioned items by:
 *   1. clustering items into rows by y (as reconstruct.ts does for lines);
 *   2. splitting each row into cells at large intra-row x-gaps;
 *   3. clustering cell-start x across ALL rows into consistent COLUMN BANDS;
 *   4. requiring ≥{@link TABLE_REGION_CONFIG.MIN_COLS} bands and a strong
 *      majority of rows that map cleanly onto them — otherwise the region is NOT
 *      confidently a table and we return `null` so the caller treats it as prose.
 *
 * A successfully-gridded region is diffed by the SHARED structure-agnostic core
 * ({@link ../diff/tables-core.ts}, also used by the Word OOXML path) with every
 * op flagged `lowConfidence` — the geometry is a heuristic, never authoritative.
 * Per ADR-0008 low-confidence is a BADGE, never a suppression: a changed cell
 * surfaces as a precise low-confidence cell change instead of whole-row prose
 * churn, and an un-griddable region still surfaces (as prose), never hidden.
 *
 * Scope note: this detects ONE table region spanning the items it is given (the
 * P12 fixture is a table-only page). Segmenting a table embedded amid prose on a
 * mixed page is a documented follow-up, not silently approximated here.
 *
 * NOT auto-wired (a deliberate decision): `runPipelinePdf` does NOT call
 * this — it is a tested standalone capability. The shipped prose path already
 * isolates value-only cell changes cleanly (a one-cell fee-table edit → a single
 * numericalValue op, no churn); B4's gain is the structural-table case (inserted/
 * reordered rows), which needs validation against real-world tables before
 * auto table-vs-prose segmentation is trusted.
 *
 * Pure — author/date injected, no Office JS, no Date.now().
 */
import type { DiffOp } from '../model.js';
import { fnv1a } from '../preprocess/hash.js';
import { normaliseText } from '../preprocess/normalise.js';
import type { PdfTextItem } from '../io/pdf/pdfjs-types.js';
import { RECONSTRUCT_CONFIG, type ReconstructConfig } from '../io/pdf/reconstruct.config.js';
import type { CellModel, GridDiffOptions, RowModel } from '../diff/tables-core.js';
import { diffGrid } from '../diff/tables-core.js';

/**
 * Thresholds for table-region gridding. Named constants only (no magic literals),
 * matching reconstruct.config.ts — the heuristic stays auditable in one place.
 * Factors multiply the region's median font size unless noted.
 */
export const TABLE_REGION_CONFIG = {
  /** An intra-row x-gap wider than this × font size starts a new CELL. */
  CELL_GAP_FACTOR: 1.5,
  /** A cell-start x within this × font size of a band centre joins that band. */
  COLUMN_BAND_TOLERANCE_FACTOR: 1.0,
  /** Two row baselines within this × font size are the SAME row. */
  ROW_Y_TOLERANCE_FACTOR: 0.3,
  /** Minimum rows for a region to be considered a table. */
  MIN_ROWS: 2,
  /** Minimum consistent column bands for a region to be considered a table. */
  MIN_COLS: 2,
  /** Fraction of rows that must map cleanly onto the bands to grid confidently. */
  ROW_CONSISTENCY_RATIO: 0.6,
  /** Fallback font size when no item has a measurable size. */
  DEFAULT_FONT_SIZE: 11,
} as const;

export type TableRegionConfig = typeof TABLE_REGION_CONFIG;

/** A positioned item reduced to the geometry the gridder needs. */
interface Cellish {
  text: string;
  x0: number;
  x1: number;
  y: number;
  fontSize: number;
}

/** A row of detected cell-clusters (pre-band-mapping). */
interface RawRow {
  y: number;
  cells: Array<{ text: string; x0: number }>;
}

/**
 * Reconstruct a confident table grid from raw positioned items, or `null` if the
 * region is not confidently a table (caller falls back to flagged prose).
 */
export function detectTableRegion(
  items: PdfTextItem[],
  config: ReconstructConfig = RECONSTRUCT_CONFIG,
  regionConfig: TableRegionConfig = TABLE_REGION_CONFIG,
): RowModel[] | null {
  const geos = toCellish(items, config);
  if (geos.length === 0) return null;

  const median = medianFontSize(geos, regionConfig);
  const rawRows = clusterRows(geos, median, regionConfig);
  if (rawRows.length < regionConfig.MIN_ROWS) return null;

  const bands = deriveColumnBands(rawRows, median, regionConfig);
  if (bands.length < regionConfig.MIN_COLS) return null;

  const tol = regionConfig.COLUMN_BAND_TOLERANCE_FACTOR * median;
  const rows: RowModel[] = [];
  let cleanRows = 0;
  for (const raw of rawRows) {
    const mapped = mapRowToBands(raw, bands, tol);
    if (mapped === null) {
      // A row whose cells collide on a band is ambiguous — skip from the clean
      // count but still emit it (an empty-padded row), never drop content.
      rows.push(buildRow(emptyCells(bands.length)));
      continue;
    }
    cleanRows++;
    rows.push(buildRow(mapped));
  }

  // Require a strong majority of rows to map cleanly; otherwise it is not a table.
  if (cleanRows / rawRows.length < regionConfig.ROW_CONSISTENCY_RATIO) return null;

  return rows;
}

/**
 * Diff two detected table grids via the shared core, flagging every op
 * low-confidence and namespacing ids under 'pdf-table-op'. Never suppresses.
 */
export function diffPdfTableRegion(
  origGrid: RowModel[],
  revGrid: RowModel[],
  author: string,
  date: string,
  counter?: { next: number },
): DiffOp[] {
  const opts: GridDiffOptions = { author, date, lowConfidence: true, idPrefix: 'pdf-table-op' };
  if (counter) opts.counter = counter;
  return diffGrid(origGrid, revGrid, opts);
}

// ---------------------------------------------------------------------------
// Geometry → rows → cells
// ---------------------------------------------------------------------------

/** Drop empty/rotated items; reduce the rest to positioned text. */
function toCellish(items: PdfTextItem[], config: ReconstructConfig): Cellish[] {
  const out: Cellish[] = [];
  for (const item of items) {
    if (item.str.trim() === '') continue;
    if (rotationDegrees(item.transform) > config.ROTATION_MAX_DEGREES) continue;
    const x0 = item.transform[4] ?? 0;
    const y = item.transform[5] ?? 0;
    out.push({
      text: item.str,
      x0,
      x1: x0 + (item.width || 0),
      y,
      fontSize: Math.hypot(item.transform[0] ?? 0, item.transform[1] ?? 0),
    });
  }
  return out;
}

/** Cluster items into rows by y (top→bottom), then into cells by intra-row gaps. */
function clusterRows(geos: Cellish[], median: number, cfg: TableRegionConfig): RawRow[] {
  const yTol = cfg.ROW_Y_TOLERANCE_FACTOR * median;
  const sorted = [...geos].sort((a, b) => b.y - a.y); // top first (y grows upward)

  const rows: RawRow[] = [];
  let current: Cellish[] = [];
  let currentY = Number.NaN;
  const flush = (): void => {
    if (current.length === 0) return;
    rows.push(buildRawRow(current, median, cfg));
    current = [];
  };
  for (const g of sorted) {
    if (current.length === 0 || Math.abs(g.y - currentY) <= yTol) {
      if (current.length === 0) currentY = g.y;
      current.push(g);
    } else {
      flush();
      current.push(g);
      currentY = g.y;
    }
  }
  flush();
  return rows;
}

/** Split one row's items into cell-clusters at large x-gaps. */
function buildRawRow(items: Cellish[], median: number, cfg: TableRegionConfig): RawRow {
  const cellGap = cfg.CELL_GAP_FACTOR * (median > 0 ? median : cfg.DEFAULT_FONT_SIZE);
  const sorted = [...items].sort((a, b) => a.x0 - b.x0);
  const cells: Array<{ text: string; x0: number }> = [];
  let text = '';
  let startX = sorted[0]!.x0;
  let prevX1 = Number.NaN;
  for (const g of sorted) {
    if (text === '') {
      text = g.text;
      startX = g.x0;
    } else if (g.x0 - prevX1 > cellGap) {
      cells.push({ text, x0: startX });
      text = g.text;
      startX = g.x0;
    } else {
      const alreadySpaced = /\s$/.test(text) || /^\s/.test(g.text);
      text += alreadySpaced ? g.text : ` ${g.text}`;
    }
    prevX1 = g.x1;
  }
  if (text !== '') cells.push({ text, x0: startX });
  return { y: sorted[0]!.y, cells };
}

// ---------------------------------------------------------------------------
// Column bands
// ---------------------------------------------------------------------------

/** Cluster cell-start x across all rows into consistent column-band centres. */
function deriveColumnBands(rawRows: RawRow[], median: number, cfg: TableRegionConfig): number[] {
  const tol = cfg.COLUMN_BAND_TOLERANCE_FACTOR * median;
  const starts: number[] = [];
  for (const row of rawRows) for (const c of row.cells) starts.push(c.x0);
  starts.sort((a, b) => a - b);

  // Agglomerate near-equal starts into bands; band centre = running mean.
  const bands: Array<{ sum: number; count: number; centre: number }> = [];
  for (const x of starts) {
    const last = bands[bands.length - 1];
    if (last && x - last.centre <= tol) {
      last.sum += x;
      last.count++;
      last.centre = last.sum / last.count;
    } else {
      bands.push({ sum: x, count: 1, centre: x });
    }
  }
  return bands.map((b) => b.centre);
}

/**
 * Map a raw row's cells onto the column bands. Returns one text per band (empty
 * for an unfilled band), or `null` if two cells collide on the same band.
 */
function mapRowToBands(
  row: RawRow,
  bands: number[],
  tol: number,
): string[] | null {
  const out = new Array<string>(bands.length).fill('');
  const filled = new Array<boolean>(bands.length).fill(false);
  for (const cell of row.cells) {
    let best = -1;
    let bestDist = Infinity;
    for (let b = 0; b < bands.length; b++) {
      const d = Math.abs(cell.x0 - bands[b]!);
      if (d < bestDist) {
        bestDist = d;
        best = b;
      }
    }
    if (best < 0 || bestDist > tol) return null; // a cell that fits no band → not a clean grid row
    if (filled[best]) return null; // two cells in one band → ambiguous
    filled[best] = true;
    out[best] = cell.text;
  }
  return out;
}

// ---------------------------------------------------------------------------
// RowModel construction
// ---------------------------------------------------------------------------

function buildRow(cellTexts: string[]): RowModel {
  const cells: CellModel[] = cellTexts.map((t) => {
    const text = normaliseText(t);
    return { text, hash: fnv1a(text) };
  });
  const rowHash = fnv1a(cells.map((c) => c.hash).join('|'));
  return { cells, hash: rowHash };
}

function emptyCells(n: number): string[] {
  return new Array<string>(n).fill('');
}

// ---------------------------------------------------------------------------
// Geometry helpers (mirrors reconstruct.ts)
// ---------------------------------------------------------------------------

function rotationDegrees(t: number[]): number {
  const deg = Math.abs((Math.atan2(t[1] ?? 0, t[0] ?? 1) * 180) / Math.PI);
  return deg > 90 ? 180 - deg : deg;
}

function medianFontSize(geos: Cellish[], cfg: TableRegionConfig): number {
  const sizes = geos.map((g) => g.fontSize).filter((s) => s > 0).sort((a, b) => a - b);
  if (sizes.length === 0) return cfg.DEFAULT_FONT_SIZE;
  const mid = Math.floor(sizes.length / 2);
  return sizes.length % 2 === 0 ? (sizes[mid - 1]! + sizes[mid]!) / 2 : sizes[mid]!;
}
