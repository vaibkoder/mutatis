/**
 * Structured table redline author — Text-vs-text Markdown path (ADR-0027).
 * Spec: spec/features/08-compare.md §9.3/§9.6.
 *
 * Authors a tracked-changes `<w:tbl>` from a parsed table grid (or a pair, for a
 * change). It activates the row-then-cell LCS the spec describes (here via the
 * shared `computeLCS` primitive) and emits native Word revision markup:
 *   - **cell value change** → clean whole-cell `<w:del>` + `<w:ins>` runs
 *     (deterministic, conventional redline style — a changed amount reads as old-struck /
 *     new-inserted, not digit-level salad; cf. ADR-0026);
 *   - **added / removed column** → `<w:tcPr><w:cellIns>/<w:cellDel>` + wrapped runs;
 *   - **added / removed row** → `<w:trPr><w:ins>/<w:del>` row-mark + wrapped runs;
 *   - **whole table added / removed** → every row/cell marked.
 *
 * Column alignment: when both sides have the SAME column count, columns align by
 * index (so a renamed header is a cell replace, not a column move); when counts
 * differ, the header row's cell-hash LCS locates the structural column add/remove
 * and that map is applied uniformly to every row. Faithful, never-suppress
 * (ADR-0008): every original-or-revised cell appears in the output exactly once.
 *
 * Pure: author/date/ids come from the injected {@link AuthorContext}.
 */
import type { RowModel } from '../diff/tables-core.js';
import { computeLCS } from '../diff/lcs.js';
import type { AuthorContext } from './author.js';
import { wrapIns, wrapDel } from './revision-markup.js';
import {
  cellRunXml,
  cleanCellParagraph,
  columnWidths,
  tableCellXml,
  tableRowXml,
  wrapTable,
} from '../io/text/synthesise-table.js';

// ---------------------------------------------------------------------------
// Revision marks (cellIns/cellDel/row-ins/row-del) — small local builders
// ---------------------------------------------------------------------------

function escapeAttr(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** A self-closing revision mark element, e.g. `<w:cellIns w:id=… w:author=… w:date=…/>`. */
function revMark(tag: string, ctx: AuthorContext): string {
  return `<w:${tag} w:id="${ctx.nextId()}" w:author="${escapeAttr(ctx.author)}" w:date="${escapeAttr(ctx.date)}"/>`;
}

/** `<w:trPr>` carrying a row insertion/deletion mark. */
function rowMarkTrPr(kind: 'ins' | 'del', ctx: AuthorContext): string {
  return `<w:trPr>${revMark(kind, ctx)}</w:trPr>`;
}

// ---------------------------------------------------------------------------
// Cell paragraph builders
// ---------------------------------------------------------------------------

/** A cell paragraph whose (single) run is wrapped in `<w:ins>`; empty → `<w:p/>`. */
function insCellParagraph(text: string, ctx: AuthorContext): string {
  const run = cellRunXml(text);
  if (run === '') return '<w:p/>';
  return `<w:p>${wrapIns(run, ctx.nextId(), ctx.author, ctx.date)}</w:p>`;
}

/** A cell paragraph whose (single) run is wrapped in `<w:del>` (→ `<w:delText>`); empty → `<w:p/>`. */
function delCellParagraph(text: string, ctx: AuthorContext): string {
  const run = cellRunXml(text);
  if (run === '') return '<w:p/>';
  return `<w:p>${wrapDel(run, ctx.nextId(), ctx.author, ctx.date)}</w:p>`;
}

/** A changed cell: old value struck (`<w:del>`) then new value inserted (`<w:ins>`). */
function changedCellParagraph(oldText: string, newText: string, ctx: AuthorContext): string {
  const delRun = cellRunXml(oldText);
  const insRun = cellRunXml(newText);
  const del = delRun === '' ? '' : wrapDel(delRun, ctx.nextId(), ctx.author, ctx.date);
  const ins = insRun === '' ? '' : wrapIns(insRun, ctx.nextId(), ctx.author, ctx.date);
  return `<w:p>${del}${ins}</w:p>`;
}

// ---------------------------------------------------------------------------
// Column alignment
// ---------------------------------------------------------------------------

type ColSlot =
  | { kind: 'match'; o: number; r: number }
  | { kind: 'ins'; r: number }
  | { kind: 'del'; o: number };

/**
 * Align the columns of two tables using their header rows. Equal width → align by
 * index (renames are cell replaces). Differing width → header-hash LCS isolates
 * the structural column add/remove, applied uniformly to every row.
 */
function alignColumns(origHeader: RowModel, revHeader: RowModel): ColSlot[] {
  const wo = origHeader.cells.length;
  const wr = revHeader.cells.length;
  if (wo === wr) {
    return Array.from({ length: wr }, (_v, i): ColSlot => ({ kind: 'match', o: i, r: i }));
  }
  const matches = computeLCS(origHeader.cells, revHeader.cells, (c) => c.hash);
  const slots: ColSlot[] = [];
  let oi = 0;
  let ri = 0;
  for (const [mo, mr] of matches) {
    while (ri < mr) slots.push({ kind: 'ins', r: ri++ });
    while (oi < mo) slots.push({ kind: 'del', o: oi++ });
    slots.push({ kind: 'match', o: mo, r: mr });
    oi = mo + 1;
    ri = mr + 1;
  }
  while (ri < wr) slots.push({ kind: 'ins', r: ri++ });
  while (oi < wo) slots.push({ kind: 'del', o: oi++ });
  return slots;
}

// ---------------------------------------------------------------------------
// Row emitters
// ---------------------------------------------------------------------------

/** A row present on both sides — per-column compare via the shared slot map. */
function emitAlignedRow(
  origRow: RowModel,
  revRow: RowModel,
  slots: ColSlot[],
  ctx: AuthorContext,
): string {
  const widths = columnWidths(slots.length);
  let cells = '';
  slots.forEach((slot, idx) => {
    const w = widths[idx] ?? 0;
    if (slot.kind === 'match') {
      const o = origRow.cells[slot.o]?.text ?? '';
      const r = revRow.cells[slot.r]?.text ?? '';
      cells += o === r ? tableCellXml(w, cleanCellParagraph(r)) : tableCellXml(w, changedCellParagraph(o, r, ctx));
    } else if (slot.kind === 'ins') {
      const r = revRow.cells[slot.r]?.text ?? '';
      cells += tableCellXml(w, insCellParagraph(r, ctx), revMark('cellIns', ctx));
    } else {
      const o = origRow.cells[slot.o]?.text ?? '';
      cells += tableCellXml(w, delCellParagraph(o, ctx), revMark('cellDel', ctx));
    }
  });
  return tableRowXml(cells);
}

/**
 * A wholly inserted (or deleted) row. The `<w:trPr>` row-mark records the row
 * insertion/deletion; each cell's runs are wrapped in `<w:ins>`/`<w:del>`. No
 * per-cell `<w:cellIns>/<w:cellDel>` here — those mark a cell/column added to an
 * *existing* row, which would be redundant inside an already-inserted row.
 */
function emitWholeRow(row: RowModel, kind: 'ins' | 'del', ctx: AuthorContext): string {
  const widths = columnWidths(row.cells.length);
  const cells = row.cells
    .map((c, i) =>
      tableCellXml(widths[i] ?? 0, kind === 'ins' ? insCellParagraph(c.text, ctx) : delCellParagraph(c.text, ctx)),
    )
    .join('');
  return tableRowXml(cells, rowMarkTrPr(kind, ctx));
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/** Author a wholly-inserted table — every row/cell wrapped in insert markup. */
export function authorTableInsert(grid: RowModel[], ctx: AuthorContext): string {
  const ncols = grid[0]?.cells.length ?? 0;
  const rowsXml = grid.map((row) => emitWholeRow(row, 'ins', ctx)).join('');
  return wrapTable(columnWidths(ncols), rowsXml);
}

/** Author a wholly-deleted table — every row/cell wrapped in delete markup. */
export function authorTableDelete(grid: RowModel[], ctx: AuthorContext): string {
  const ncols = grid[0]?.cells.length ?? 0;
  const rowsXml = grid.map((row) => emitWholeRow(row, 'del', ctx)).join('');
  return wrapTable(columnWidths(ncols), rowsXml);
}

/**
 * Author a changed table: row-hash LCS, then per aligned row pair a column-aware
 * cell compare; surplus rows become whole-row insert/delete. Identical tables
 * produce zero revision markup (every cell clean).
 */
export function authorTableReplace(orig: RowModel[], rev: RowModel[], ctx: AuthorContext): string {
  if (orig.length === 0) return authorTableInsert(rev, ctx);
  if (rev.length === 0) return authorTableDelete(orig, ctx);

  const slots = alignColumns(orig[0]!, rev[0]!);
  const rowMatches = computeLCS(orig, rev, (r) => r.hash);

  let rowsXml = '';
  let oi = 0;
  let ri = 0;
  const handleGap = (oEnd: number, rEnd: number): void => {
    const oRemain = oEnd - oi;
    const rRemain = rEnd - ri;
    const paired = Math.min(oRemain, rRemain);
    for (let k = 0; k < paired; k++) rowsXml += emitAlignedRow(orig[oi + k]!, rev[ri + k]!, slots, ctx);
    for (let k = paired; k < oRemain; k++) rowsXml += emitWholeRow(orig[oi + k]!, 'del', ctx);
    for (let k = paired; k < rRemain; k++) rowsXml += emitWholeRow(rev[ri + k]!, 'ins', ctx);
    oi = oEnd;
    ri = rEnd;
  };

  for (const [mo, mr] of rowMatches) {
    handleGap(mo, mr);
    rowsXml += emitAlignedRow(orig[mo]!, rev[mr]!, slots, ctx);
    oi = mo + 1;
    ri = mr + 1;
  }
  handleGap(orig.length, rev.length);

  return wrapTable(columnWidths(slots.length), rowsXml);
}
