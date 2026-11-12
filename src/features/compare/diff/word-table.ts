/**
 * Word table analysis — the row and cell alignment of a Word table pair
 * (ADR-0046). Spec: spec/features/08-compare.md §9.3, §9.6.
 *
 * A Word body table is ONE alignment unit (`preprocess/table-units.ts`), so the
 * body's Level-1 diff pairs table with table. This stage looks inside every such
 * pair:
 *
 *  - **Rows** are matched on their hash (text AND layout) by LCS. What lies
 *    between two matches is a gap: with the same number of rows on each side
 *    the rows were edited in place and pair by position; otherwise they pair
 *    order-preserving by similarity (the ADR-0039 rule at row scale), so one
 *    inserted row does not shift the pairing of the rows after it. A row left
 *    over is a whole-row insert or delete.
 *  - **Cells** of a paired row with the same layout (cell count, `gridSpan`,
 *    `vMerge`) pair one to one, and each cell's content is analysed by the
 *    analyser passed in — `analyseRecords`, the body's own stages, so a cell is
 *    compared by exactly the rules a paragraph is (ADR-0043 mark order
 *    included). A paired row whose layout changed (a merge, a split, a column)
 *    is a delete of the original row plus an insert of the revised one.
 *
 * Never suppresses (ADR-0008): every original and every revised row appears in
 * exactly one slot, and every cell of a `match` row in exactly one cell slot.
 */
import type { CompareConfig } from '@shared/config/index.js';

import type {
  DiffOp, ParaRecord, TableCellSlot, TableColumnSlot, TableDiff, TableRowSlot, WordTableCell, WordTableModel,
  WordTableRow,
} from '../model.js';

import { computeLCS } from './lcs.js';
import { bagDice, bagOf, contentTokens } from './replace-threshold.js';

/** Analyses one cell's content: the two cells' records → that content's ops. */
export type CellAnalyser = (original: ParaRecord[], revised: ParaRecord[]) => Promise<DiffOp[]>;

/** Above this many cells in one row gap, pair positionally instead of by similarity. */
const MAX_GAP_CELLS = 250_000;

/**
 * Stage 2.8 — attach a {@link TableDiff} to every `equal` or `replace` op that
 * pairs two Word table units, and turn an `equal` whose cells changed in a way
 * text identity cannot see (a style, a number, a footnote alone) into a
 * `property` op, as ADR-0042 does for a paragraph. Returns a new list.
 */
export async function analyseWordTables(
  ops: DiffOp[],
  original: ParaRecord[],
  revised: ParaRecord[],
  config: CompareConfig,
  analyse: CellAnalyser,
): Promise<DiffOp[]> {
  const out: DiffOp[] = [];
  for (const op of ops) {
    const o = op.origRange ? original[op.origRange.start] : undefined;
    const r = op.revRange ? revised[op.revRange.start] : undefined;
    const pairs = (op.type === 'equal' || op.type === 'replace') && o?.wordTable && r?.wordTable;
    if (!pairs || o === undefined || r === undefined || o.wordTable === undefined || r.wordTable === undefined) {
      out.push(op);
      continue;
    }
    // A byte-identical table carries no change of any kind: nothing to look at.
    if (op.type === 'equal' && o.rawXml === r.rawXml) {
      out.push(op);
      continue;
    }
    const tableDiff = await diffWordTable(o.wordTable, r.wordTable, config, analyse);
    if (op.type === 'replace') {
      out.push({ ...op, tableDiff });
      continue;
    }
    // Equal text and shape. A cell may still have changed what text identity
    // does not anchor on — surface it, exactly as a paragraph's would be.
    const first = firstNestedChange(tableDiff);
    if (first === undefined) {
      out.push(op);
      continue;
    }
    const property: DiffOp = { ...op, type: 'property', tableDiff };
    if (first.propertyChange !== undefined) {
      // The cascade marker is dropped: a table op is never collapsed into a
      // §6.1 cascade, which would take every other cell change with it.
      const { cascade: _cascade, ...change } = first.propertyChange;
      void _cascade;
      property.propertyChange = change;
    }
    out.push(property);
  }
  return out;
}

/** Row then cell alignment of one table pair. */
export async function diffWordTable(
  orig: WordTableModel,
  rev: WordTableModel,
  config: CompareConfig,
  analyse: CellAnalyser,
): Promise<TableDiff> {
  const plan = alignRows(orig.rows, rev.rows, config, projectedRowKeys(orig, rev));
  const columns = columnMapFor(orig, rev, plan);
  const rows: TableRowSlot[] = [];
  for (const slot of plan) {
    if (slot.kind === 'delete') {
      rows.push({ kind: 'delete', o: slot.i });
      continue;
    }
    if (slot.kind === 'insert') {
      rows.push({ kind: 'insert', r: slot.i });
      continue;
    }
    const or = orig.rows[slot.o]!;
    const rr = rev.rows[slot.r]!;
    if (sameLayout(or, rr)) {
      const cells: TableCellSlot[] = [];
      for (let k = 0; k < rr.cells.length; k++) {
        cells.push({ kind: 'match', o: k, r: k, ops: await cellOps(or.cells[k]!, rr.cells[k]!, analyse) });
      }
      rows.push({ kind: 'match', o: slot.o, r: slot.r, cells });
    } else if (columns !== undefined && fitsColumns(or, rr, columns)) {
      const cells: TableCellSlot[] = [];
      for (const col of columns) {
        if (col.kind === 'insert') cells.push({ kind: 'insert', r: col.r });
        else if (col.kind === 'delete') cells.push({ kind: 'delete', o: col.o });
        else cells.push({ kind: 'match', o: col.o, r: col.r, ops: await cellOps(or.cells[col.o]!, rr.cells[col.r]!, analyse) });
      }
      rows.push({ kind: 'match', o: slot.o, r: slot.r, cells });
    } else {
      rows.push({ kind: 'delete', o: slot.o }, { kind: 'insert', r: slot.r });
    }
  }
  return columns === undefined ? { rows } : { rows, columns };
}

/**
 * LCS on row keys — the row hashes, or {@link projectedRowKeys} when a column
 * was added or removed — then each gap paired by {@link pairGap}, in document
 * order.
 */
function alignRows(
  orig: readonly WordTableRow[],
  rev: readonly WordTableRow[],
  config: CompareConfig,
  keys: RowKeys | undefined,
): GapSlot[] {
  const o = keys?.o ?? orig.map((row) => row.hash);
  const r = keys?.r ?? rev.map((row) => row.hash);
  const out: GapSlot[] = [];
  let oi = 0;
  let ri = 0;
  for (const [mo, mr] of computeLCS(o, r, (key) => key)) {
    out.push(...pairGap(orig, oi, mo, rev, ri, mr, config), { kind: 'pair', o: mo, r: mr });
    oi = mo + 1;
    ri = mr + 1;
  }
  out.push(...pairGap(orig, oi, orig.length, rev, ri, rev.length, config));
  return out;
}

/** One alignment key per row on each side. */
interface RowKeys {
  o: string[];
  r: string[];
}

/** Column subsets tried when a column was added or removed; past this, row hashes alone. */
const MAX_COLUMN_SUBSETS = 64;

/**
 * Row keys that see past a column added or removed. A new column changes every
 * row's hash, so the row LCS finds nothing, the whole table becomes one gap of
 * equal length and is paired by POSITION — and a row deleted plus a row
 * inserted elsewhere then pairs every row in between with its neighbour. When
 * the grid's width changed, each full-width row (one unmerged cell per grid
 * column — see {@link fullWidth}) is keyed by its cells with the wider side's
 * surplus columns left out: the subset left out is the one under which the
 * most rows agree. Any other row keeps its own hash. Undefined where the width
 * did not change, too many subsets exist, or no row agrees under any.
 */
function projectedRowKeys(orig: WordTableModel, rev: WordTableModel): RowKeys | undefined {
  const wo = gridWidth(orig);
  const wr = gridWidth(rev);
  if (wo === wr) return undefined;
  const revWider = wr > wo;
  const [narrow, wide] = revWider ? [orig, rev] : [rev, orig];
  const [wn, ww] = revWider ? [wo, wr] : [wr, wo];
  const subsets = columnSubsets(ww, ww - wn);
  if (subsets === undefined) return undefined;
  const key = (row: WordTableRow, width: number, drop: ReadonlySet<number>): string =>
    fullWidth(row, width)
      ? row.cells.filter((_c, i) => !drop.has(i)).map((c) => `${c.vMerge ?? ''}:${c.hash}`).join('|')
      : `row:${row.hash}`;
  const narrowKeys = narrow.rows.map((row) => key(row, wn, new Set()));
  let best: { keys: string[]; agree: number } | undefined;
  for (const drop of subsets) {
    const keys = wide.rows.map((row) => key(row, ww, drop));
    const agree = computeLCS(narrowKeys, keys, (k) => k).length;
    if (best === undefined || agree > best.agree) best = { keys, agree };
  }
  if (best === undefined || best.agree === 0) return undefined;
  return revWider ? { o: narrowKeys, r: best.keys } : { o: best.keys, r: narrowKeys };
}

/** Every `k`-element subset of `0…n-1`, or undefined past {@link MAX_COLUMN_SUBSETS}. */
function columnSubsets(n: number, k: number): Set<number>[] | undefined {
  const out: Set<number>[] = [];
  const pick = (from: number, chosen: number[]): boolean => {
    if (chosen.length === k) {
      out.push(new Set(chosen));
      return out.length <= MAX_COLUMN_SUBSETS;
    }
    for (let i = from; i < n; i++) if (!pick(i + 1, [...chosen, i])) return false;
    return true;
  };
  return pick(0, []) ? out : undefined;
}

/** A table's width in grid columns: its `w:tblGrid`, else its widest row. */
function gridWidth(t: WordTableModel): number {
  const declared = (t.tblGridXml.match(/<w:gridCol[\s/>]/g) ?? []).length;
  if (declared > 0) return declared;
  return Math.max(0, ...t.rows.map((row) => row.cells.reduce((n, c) => n + c.gridSpan, 0)));
}

/**
 * One cell per grid column, none merged across columns: the rows a column map
 * can be laid over. (A cell may still be merged DOWN — `vMerge` — which the map
 * carries cell by cell; see {@link fitsColumns}.)
 */
function fullWidth(row: WordTableRow, width: number): boolean {
  return row.cells.length === width && row.cells.every((c) => c.gridSpan === 1);
}

/**
 * Whether a paired row pair can be revised through `columns`: both rows full
 * width for their side's grid, and every cell the map pairs merged down the
 * same way on both sides — a cell whose vertical merge changed is a layout
 * change, and its row reads as out and in.
 */
function fitsColumns(or: WordTableRow, rr: WordTableRow, columns: readonly TableColumnSlot[]): boolean {
  const wo = columns.filter((c) => c.kind !== 'insert').length;
  const wr = columns.filter((c) => c.kind !== 'delete').length;
  return (
    fullWidth(or, wo) &&
    fullWidth(rr, wr) &&
    columns.every((c) => c.kind !== 'match' || or.cells[c.o]!.vMerge === rr.cells[c.r]!.vMerge)
  );
}

/**
 * A column added or removed across the whole table, as a column map — or
 * undefined. For a grid whose width changed, read from the paired rows that are
 * {@link fullWidth} on both sides (at least one is needed); a merged row — a
 * heading or total spanning the table — cannot take the map and stays a row
 * out and a row in, which is always valid. The columns pair order-preserving,
 * the pairing that keeps the most paired rows' cells identical; the surplus
 * columns on the wider side are the ones added or removed.
 */
function columnMapFor(
  orig: WordTableModel,
  rev: WordTableModel,
  plan: readonly GapSlot[],
): TableColumnSlot[] | undefined {
  const wo = gridWidth(orig);
  const wr = gridWidth(rev);
  const pairs = plan.filter(
    (s): s is Extract<GapSlot, { kind: 'pair' }> =>
      s.kind === 'pair' && fullWidth(orig.rows[s.o]!, wo) && fullWidth(rev.rows[s.r]!, wr),
  );
  if (wo === wr || wo === 0 || wr === 0 || pairs.length === 0) return undefined;
  // agree[i][j]: paired rows whose original column i and revised column j hold identical content.
  const agree = Array.from({ length: wo }, (_v, i) =>
    Array.from({ length: wr }, (_w, j) =>
      pairs.filter((p) => orig.rows[p.o]!.cells[i]!.hash === rev.rows[p.r]!.cells[j]!.hash).length,
    ),
  );
  const keep = Math.min(wo, wr);
  // best[i][j] = the best total agreement pairing EVERY remaining column of the
  // narrower side (from orig i.., rev j..) with one of the wider side's, in
  // order; NEG where that is impossible. The wider side's columns left over are
  // the ones added or removed.
  const NEG = -1;
  const best: number[][] = Array.from({ length: wo + 1 }, () => new Array<number>(wr + 1).fill(NEG));
  for (let i = wo; i >= 0; i--) {
    for (let j = wr; j >= 0; j--) {
      const leftO = wo - i;
      const leftR = wr - j;
      if ((wo <= wr && leftO === 0) || (wr < wo && leftR === 0)) {
        best[i]![j] = 0;
        continue;
      }
      let b = NEG;
      if (i < wo && j < wr && best[i + 1]![j + 1]! >= 0) b = agree[i]![j]! + best[i + 1]![j + 1]!;
      // Skip a column on the wider side only (the narrower side pairs every column).
      if (wr > wo && j < wr && best[i]![j + 1]! > b) b = best[i]![j + 1]!;
      if (wo > wr && i < wo && best[i + 1]![j]! > b) b = best[i + 1]![j]!;
      best[i]![j] = b;
    }
  }
  if (best[0]![0]! < 0) return undefined;
  const slots: TableColumnSlot[] = [];
  let i = 0;
  let j = 0;
  while (i < wo || j < wr) {
    const pairScore = i < wo && j < wr && best[i + 1]![j + 1]! >= 0 ? agree[i]![j]! + best[i + 1]![j + 1]! : NEG;
    if (pairScore >= 0 && pairScore === best[i]![j]!) {
      slots.push({ kind: 'match', o: i++, r: j++ });
    } else if (wr > wo && j < wr) {
      slots.push({ kind: 'insert', r: j++ });
    } else if (i < wo) {
      slots.push({ kind: 'delete', o: i++ });
    } else {
      slots.push({ kind: 'insert', r: j++ });
    }
  }
  return slots.filter((s) => s.kind === 'match').length === keep ? slots : undefined;
}

/**
 * A cell pair's content ops. Content whose XML is identical paragraph for
 * paragraph has no change of any kind and gets no ops (the author re-emits it
 * verbatim); anything else goes through the body's analysis.
 */
async function cellOps(o: WordTableCell, r: WordTableCell, analyse: CellAnalyser): Promise<DiffOp[]> {
  const identical = o.content.length === r.content.length && o.content.every((rec, i) => rec.rawXml === r.content[i]!.rawXml);
  return identical ? [] : analyse(o.content, r.content);
}

/** Same cell count, and every cell spans and merges the same way. */
function sameLayout(a: WordTableRow, b: WordTableRow): boolean {
  return (
    a.cells.length === b.cells.length &&
    a.cells.every((c, i) => c.gridSpan === b.cells[i]!.gridSpan && c.vMerge === b.cells[i]!.vMerge)
  );
}

type GapSlot = { kind: 'pair'; o: number; r: number } | { kind: 'delete'; i: number } | { kind: 'insert'; i: number };

/**
 * Pair the rows of one gap. Equal counts → by position (edited in place).
 * Otherwise the order-preserving pairing of maximal total similarity among
 * pairs at or above `replaceSimilarityThreshold`; leftovers are emitted in
 * document order, a deletion before the insertion it sits beside.
 */
function pairGap(
  orig: readonly WordTableRow[], oStart: number, oEnd: number,
  rev: readonly WordTableRow[], rStart: number, rEnd: number,
  config: CompareConfig,
): GapSlot[] {
  const m = oEnd - oStart;
  const n = rEnd - rStart;
  const slots: GapSlot[] = [];
  if (m === n || m * n > MAX_GAP_CELLS) {
    const paired = Math.min(m, n);
    for (let k = 0; k < paired; k++) slots.push({ kind: 'pair', o: oStart + k, r: rStart + k });
    for (let k = paired; k < m; k++) slots.push({ kind: 'delete', i: oStart + k });
    for (let k = paired; k < n; k++) slots.push({ kind: 'insert', i: rStart + k });
    return slots;
  }
  const threshold = config.replaceSimilarityThreshold;
  const sim = (i: number, j: number): number => rowSimilarity(orig[oStart + i]!, rev[rStart + j]!, config);
  // best[i][j] = best total over orig rows i.. and rev rows j..
  const best: number[][] = Array.from({ length: m + 1 }, () => new Array<number>(n + 1).fill(0));
  const take: boolean[][] = Array.from({ length: m + 1 }, () => new Array<boolean>(n + 1).fill(false));
  const score: number[][] = Array.from({ length: m }, (_v, i) => Array.from({ length: n }, (_w, j) => sim(i, j)));
  for (let i = m - 1; i >= 0; i--) {
    for (let j = n - 1; j >= 0; j--) {
      const skip = Math.max(best[i + 1]![j]!, best[i]![j + 1]!);
      const s = score[i]![j]!;
      const pair = s >= threshold ? s + best[i + 1]![j + 1]! : -1;
      if (pair > skip) {
        best[i]![j] = pair;
        take[i]![j] = true;
      } else {
        best[i]![j] = skip;
      }
    }
  }
  let i = 0;
  let j = 0;
  while (i < m && j < n) {
    if (take[i]![j]) {
      slots.push({ kind: 'pair', o: oStart + i, r: rStart + j });
      i++;
      j++;
    } else if (best[i + 1]![j]! >= best[i]![j + 1]!) {
      slots.push({ kind: 'delete', i: oStart + i++ });
    } else {
      slots.push({ kind: 'insert', i: rStart + j++ });
    }
  }
  while (i < m) slots.push({ kind: 'delete', i: oStart + i++ });
  while (j < n) slots.push({ kind: 'insert', i: rStart + j++ });
  return slots;
}

/**
 * How alike two rows are: their text's Dice (ADR-0026's arithmetic), or — when
 * they have the same cell count — the share of cells whose content is
 * identical, whichever is higher. A row's key cell ("Agency fee") staying put
 * while its amount changes is the same row, even where the words are too few
 * for Dice to say so.
 */
function rowSimilarity(a: WordTableRow, b: WordTableRow, config: CompareConfig): number {
  const text = (row: WordTableRow): string =>
    row.cells.map((c) => c.content.map((p) => p.normText).join(' ')).join(' ');
  const dice = bagDice(bagOf(contentTokens(text(a), config)), bagOf(contentTokens(text(b), config)));
  if (a.cells.length !== b.cells.length || a.cells.length === 0) return dice;
  const same = a.cells.filter((c, k) => c.hash === b.cells[k]!.hash).length;
  return Math.max(dice, same / a.cells.length);
}

/** The first non-equal op inside a table diff, depth first, or undefined. */
export function firstNestedChange(diff: TableDiff): DiffOp | undefined {
  for (const row of diff.rows) {
    if (row.kind !== 'match') continue;
    for (const cell of row.cells) {
      if (cell.kind !== 'match') continue;
      for (const op of cell.ops) {
        if (op.type !== 'equal') return op;
      }
    }
  }
  return undefined;
}

/** True when a table diff carries any change at all: a row or cell slot, or a nested op. */
export function tableDiffChanges(diff: TableDiff): boolean {
  return diff.rows.some(
    (row) =>
      row.kind !== 'match' ||
      row.cells.some((cell) => cell.kind !== 'match' || cell.ops.some((op) => op.type !== 'equal')),
  );
}
