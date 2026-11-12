/**
 * Structure-agnostic table-grid diff core.
 * Spec: spec/features/08-compare.md §9.3, §9.6 (Word path); the PDF approximate table-region path.
 *
 * This is the cell-LCS alignment decoupled from any source format. It knows only
 * about a grid of {@link RowModel}s (rows of {@link CellModel}s with content
 * hashes) — it does NOT know whether those rows came from an OOXML `<w:tbl>` or
 * from geometry-reconstructed PDF text bands. The two callers are:
 *   - {@link ../diff/tables.ts} (Word OOXML): exact, full-confidence.
 *   - {@link ../pdf/table-region.ts} (PDF): approximate, every op flagged
 *     `lowConfidence` via {@link GridDiffOptions.lowConfidence}.
 *
 * Adding a column shifts every cell's index, so naive position diffing reports
 * the whole table as changed. Instead we align rows first (row-hash LCS), then
 * cells within aligned rows (cell-hash LCS): a genuinely added column surfaces as
 * `structural` insert ops, not a wholesale rewrite. Pure — author/date injected,
 * never sourced from Date.now(). Never suppresses (ADR-0008): low-confidence ops
 * are badged, never hidden.
 */
import type { DiffOp, ChangeClass } from '../model.js';

import { computeLCS } from './lcs.js';

/** A parsed table cell: its normalised text + content hash. */
export interface CellModel {
  text: string;
  hash: string;
}

/** A parsed table row: its cells + a row hash (concatenation of cell hashes). */
export interface RowModel {
  cells: CellModel[];
  hash: string;
}

/** Knobs shared by every op the grid diff emits. */
export interface GridDiffOptions {
  author: string;
  date: string;
  /**
   * Mark every produced op low-confidence — set by the PDF approximate
   * table-region path (B4). The UI badges it; never a suppression (ADR-0008).
   */
  lowConfidence?: boolean;
  /** Shared op-id counter so ids stay unique across many grids; created if omitted. */
  counter?: { next: number };
  /** Op-id prefix (default 'table-op'; the PDF path uses 'pdf-table-op'). */
  idPrefix?: string;
}

/** Internal: the resolved (defaults-applied) options threaded through the diff. */
interface ResolvedOptions {
  author: string;
  date: string;
  lowConfidence: boolean;
  counter: { next: number };
  idPrefix: string;
}

function resolve(opts: GridDiffOptions): ResolvedOptions {
  return {
    author: opts.author,
    date: opts.date,
    lowConfidence: opts.lowConfidence ?? false,
    counter: opts.counter ?? { next: 0 },
    idPrefix: opts.idPrefix ?? 'table-op',
  };
}

/**
 * Diff one aligned grid pair: row-LCS, then cell-LCS within aligned rows. Returns
 * the ops in document order. Callers that have multiple tables drive this once
 * per aligned pair (sharing `opts.counter`) and handle whole-table add/remove
 * themselves via {@link rowOp}.
 */
export function diffGrid(
  origRows: RowModel[],
  revRows: RowModel[],
  opts: GridDiffOptions,
): DiffOp[] {
  const resolved = resolve(opts);
  const ops: DiffOp[] = [];
  diffGridInto(origRows, revRows, resolved, ops);
  return ops;
}

/** Row-LCS then per-row cell-LCS, appending ops to `ops`. */
function diffGridInto(
  origRows: RowModel[],
  revRows: RowModel[],
  opts: ResolvedOptions,
  ops: DiffOp[],
): void {
  const rowMatches = computeLCS(origRows, revRows, (r) => r.hash);

  let oi = 0;
  let ri = 0;
  const handleRowGap = (oEnd: number, rEnd: number): void => {
    // Index-align gap rows; surplus rows become whole-row insert/delete.
    const oRemain = oEnd - oi;
    const rRemain = rEnd - ri;
    const paired = Math.min(oRemain, rRemain);
    for (let k = 0; k < paired; k++) {
      diffRowCells(oi + k, ri + k, origRows[oi + k]!, revRows[ri + k]!, opts, ops);
    }
    for (let k = paired; k < oRemain; k++) {
      ops.push(rowOp('delete', oi + k, origRows[oi + k]!, opts));
    }
    for (let k = paired; k < rRemain; k++) {
      ops.push(rowOp('insert', ri + k, revRows[ri + k]!, opts));
    }
    oi = oEnd;
    ri = rEnd;
  };

  for (const [mo, mr] of rowMatches) {
    handleRowGap(mo, mr);
    // Hash-equal row → cells identical; emit one equal op per cell for completeness.
    diffRowCells(mo, mr, origRows[mo]!, revRows[mr]!, opts, ops);
    oi = mo + 1;
    ri = mr + 1;
  }
  handleRowGap(origRows.length, revRows.length);
}

/**
 * Align the cells of an aligned row pair via cell-hash LCS. Added/removed cells
 * (e.g. a new column) → structural insert/delete; changed cells → replace;
 * identical cells → equal.
 */
function diffRowCells(
  rowO: number,
  rowR: number,
  origRow: RowModel,
  revRow: RowModel,
  opts: ResolvedOptions,
  ops: DiffOp[],
): void {
  const cellMatches = computeLCS(origRow.cells, revRow.cells, (c) => c.hash);

  let oi = 0;
  let ri = 0;
  const handleCellGap = (oEnd: number, rEnd: number): void => {
    const oRemain = oEnd - oi;
    const rRemain = rEnd - ri;
    const paired = Math.min(oRemain, rRemain);
    for (let k = 0; k < paired; k++) {
      // Same column position, different content → replace (word diff deferred).
      ops.push(cellOp('replace', rowR, ri + k, opts, ['substantiveText']));
    }
    // Surplus cells: a removed column (orig only) / added column (rev only).
    for (let k = paired; k < oRemain; k++) {
      ops.push(cellOp('delete', rowO, oi + k, opts, ['structural']));
    }
    for (let k = paired; k < rRemain; k++) {
      ops.push(cellOp('insert', rowR, ri + k, opts, ['structural']));
    }
    oi = oEnd;
    ri = rEnd;
  };

  for (const [mo, mr] of cellMatches) {
    handleCellGap(mo, mr);
    ops.push(cellOp('equal', rowR, mr, opts, ['substantiveText']));
    oi = mo + 1;
    ri = mr + 1;
  }
  handleCellGap(origRow.cells.length, revRow.cells.length);
}

// ---------------------------------------------------------------------------
// Op constructors
// ---------------------------------------------------------------------------

/** A per-cell DiffOp with partId 'tableCell:rXcY' and isTableCell semantics. */
export function cellOp(
  type: DiffOp['type'],
  row: number,
  col: number,
  opts: ResolvedOptions,
  classes: [ChangeClass, ...ChangeClass[]],
): DiffOp {
  const op: DiffOp = {
    id: `${opts.idPrefix}-${opts.counter.next++}`,
    type,
    origRange: type === 'insert' ? null : { start: row, end: row },
    revRange: type === 'delete' ? null : { start: row, end: row },
    partId: `tableCell:r${row}c${col}`,
    classification: classes[0],
    classes,
    author: opts.author,
    date: opts.date,
  };
  if (opts.lowConfidence) op.lowConfidence = true;
  return op;
}

/** A whole-row insert/delete represented as a structural op spanning its cells. */
export function rowOp(
  type: DiffOp['type'],
  row: number,
  rowModel: RowModel,
  opts: ResolvedOptions,
): DiffOp {
  void rowModel;
  return cellOp(type, row, 0, opts, ['structural']);
}

/** Resolve caller options once for direct {@link rowOp}/{@link cellOp} use. */
export function resolveGridOptions(opts: GridDiffOptions): ResolvedOptions {
  return resolve(opts);
}
