/**
 * What a Word table change (ADR-0046) looks like to the surfaces that print or
 * navigate to one: the change-list row and *Go to change*.
 * Spec: spec/features/08-compare.md §8.3.
 *
 * A table op has no `wordDiffs` of its own — its changes live in its cells'
 * ops — so every surface that reads `wordDiffs` would show a blank row and do
 * nothing when clicked. And a table's own text (`normText`) runs across cell
 * boundaries, which Word's search cannot cross. So both surfaces look INSIDE:
 * the first cell that changed.
 *
 * Pure, and free of the XML stack, so a UI component can import it.
 */
import type { DiffOp, ParaRecord, TableDiff, WordDiff, WordTableModel } from '../model.js';

/**
 * The first changed op inside a table op's cells, depth first, nested tables
 * followed down to the change itself. Undefined when the table changed only by
 * whole rows or columns.
 */
export function firstCellChange(op: DiffOp): DiffOp | undefined {
  const diff = op.tableDiff;
  if (diff === undefined) return undefined;
  for (const row of diff.rows) {
    if (row.kind !== 'match') continue;
    for (const cell of row.cells) {
      if (cell.kind !== 'match') continue;
      for (const inner of cell.ops) {
        if (inner.type === 'equal') continue;
        if (inner.tableDiff !== undefined) {
          const deeper = firstCellChange(inner);
          if (deeper !== undefined) return deeper;
          continue;
        }
        return inner;
      }
    }
  }
  return undefined;
}

/**
 * How many separate changes a table op carries: every changed op in a matched
 * cell (a nested table's counted through), every inserted or deleted cell, and
 * every inserted or deleted row. A preview that shows ONE of them says how many
 * more there are, so a row never reads as though one cell were the whole story.
 */
export function tableChangeCount(diff: TableDiff): number {
  let n = 0;
  for (const row of diff.rows) {
    if (row.kind !== 'match') {
      n++;
      continue;
    }
    for (const cell of row.cells) {
      if (cell.kind !== 'match') {
        n++;
        continue;
      }
      for (const inner of cell.ops) {
        if (inner.type === 'equal') continue;
        n += inner.tableDiff !== undefined ? Math.max(1, tableChangeCount(inner.tableDiff)) : 1;
      }
    }
  }
  return n;
}

/**
 * Whether any cell op (nested tables included) adds, removes or moves a
 * footnote, picture or object — the one property change the redline marks
 * (ADR-0045 § Amendment). A table whose cells changed only a style or a number
 * is otherwise authored as written, as a paragraph's would be.
 */
export function tableMarksObjects(diff: TableDiff): boolean {
  return diff.rows.some(
    (row) =>
      row.kind === 'match' &&
      row.cells.some(
        (cell) =>
          cell.kind === 'match' &&
          cell.ops.some((op) =>
            op.tableDiff !== undefined ? tableMarksObjects(op.tableDiff) : op.propertyChange?.objects === true,
          ),
      ),
  );
}

/** "row inserted", "2 rows deleted, column added" — the shape of a table change. */
export function tableShapeSummary(diff: TableDiff): string {
  const count = (n: number, what: string): string => `${n} ${what}${n === 1 ? '' : 's'}`;
  const rowsIn = diff.rows.filter((r) => r.kind === 'insert').length;
  const rowsOut = diff.rows.filter((r) => r.kind === 'delete').length;
  const colsIn = diff.columns?.filter((c) => c.kind === 'insert').length ?? 0;
  const colsOut = diff.columns?.filter((c) => c.kind === 'delete').length ?? 0;
  const parts: string[] = [];
  if (colsIn > 0) parts.push(`${count(colsIn, 'column')} added`);
  if (colsOut > 0) parts.push(`${count(colsOut, 'column')} removed`);
  // A row out and a row in at one place is a row whose cells were merged or split.
  if (rowsIn > 0) parts.push(`${count(rowsIn, 'row')} inserted`);
  if (rowsOut > 0) parts.push(`${count(rowsOut, 'row')} deleted`);
  return parts.length === 0 ? 'Table changed' : `Table: ${parts.join(', ')}`;
}

/**
 * Word diffs to aim *Go to change* at, each line one cell's words: the first
 * changed cell's own diff, else an inserted row's first non-empty cell, else —
 * for a deleted row, whose words are no longer in the document — the nearest
 * surviving row's first cell. Never the table's whole text.
 */
export function tableAnchorDiffs(op: DiffOp, original: readonly ParaRecord[], revised: readonly ParaRecord[]): WordDiff[] {
  const orig = op.origRange ? original[op.origRange.start]?.wordTable : undefined;
  const rev = op.revRange ? revised[op.revRange.start]?.wordTable : undefined;
  if (op.tableDiff === undefined || orig === undefined || rev === undefined) return [];
  return anchorIn(op.tableDiff, orig, rev);
}

function anchorIn(diff: TableDiff, orig: WordTableModel, rev: WordTableModel): WordDiff[] {
  const firstCellText = (table: WordTableModel, row: number): string =>
    table.rows[row]?.cells.map((c) => c.content.map((p) => p.normText).join(' ').trim()).find((t) => t !== '') ?? '';
  for (const row of diff.rows) {
    if (row.kind !== 'match') continue;
    for (const cell of row.cells) {
      if (cell.kind !== 'match') continue;
      const oc = orig.rows[row.o]!.cells[cell.o]!;
      const rc = rev.rows[row.r]!.cells[cell.r]!;
      for (const inner of cell.ops) {
        if (inner.type === 'equal') continue;
        const o = inner.origRange ? oc.content[inner.origRange.start] : undefined;
        const r = inner.revRange ? rc.content[inner.revRange.start] : undefined;
        if (inner.tableDiff !== undefined && o?.wordTable && r?.wordTable) {
          const deeper = anchorIn(inner.tableDiff, o.wordTable, r.wordTable);
          if (deeper.length > 0) return deeper;
          continue;
        }
        if (inner.wordDiffs !== undefined && inner.wordDiffs.length > 0) return inner.wordDiffs;
        if (r !== undefined) return [{ op: inner.type === 'insert' ? 'ins' : 'eq', text: r.normText }];
        // A deleted paragraph: its words are gone; aim at the cell that held it.
        const kept = rc.content.map((p) => p.normText).join(' ').trim();
        if (kept !== '') return [{ op: 'eq', text: kept }];
      }
    }
  }
  const inserted = diff.rows.find((r) => r.kind === 'insert');
  if (inserted?.kind === 'insert') {
    const text = firstCellText(rev, inserted.r);
    if (text !== '') return [{ op: 'ins', text }];
  }
  // A deleted row, or a column: aim at the first surviving row that has words.
  for (let r = 0; r < rev.rows.length; r++) {
    const text = firstCellText(rev, r);
    if (text !== '') return [{ op: 'eq', text }];
  }
  return [];
}
