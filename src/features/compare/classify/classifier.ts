/**
 * Feature 08 — Compare · classifier entry points (Task B1).
 *
 * Pure functions that label DiffOps with their §6 classification. They NEVER
 * mutate the input op's `type`, `wordDiffs`, ranges, or any other field beyond
 * `classification`/`classes`, and they NEVER drop an op (ADR-0008).
 *
 * Spec: spec/features/08-compare.md §6.
 */

import type { CompareConfig } from '@shared/config/index.js';
import type { ChangeClass, DiffOp, ParaRecord, TableDiff } from '../model.js';
import { classifyClasses, PRIORITY_ORDER } from './taxonomy.js';

/**
 * ADR-0027: classify a text-vs-text Markdown TABLE-unit op. The base taxonomy is
 * run/word-diff based and does not model tables, so a table op is labelled here:
 * a whole-table insert/delete or a shape change (row/column count differs) is
 * `structural`; a same-shape content change is `substantiveText`. Either way it
 * is never `formatting`/`whitespacePunct`, so a table change is never filtered
 * (§7.3, ADR-0008). Reached only when a record carries `.table` (text path only).
 */
function classifyTableOp(op: DiffOp, origRec?: ParaRecord, revRec?: ParaRecord): DiffOp {
  if (op.tableDiff !== undefined) return classifyWordTableOp(op, op.tableDiff);
  let classes: ChangeClass[];
  if (op.type === 'insert' || op.type === 'delete' || op.type === 'move') {
    classes = ['structural', 'substantiveText'];
  } else {
    const o = origRec?.table;
    const r = revRec?.table;
    const shapeChanged =
      !!o && !!r && (o.length !== r.length || (o[0]?.cells.length ?? 0) !== (r[0]?.cells.length ?? 0));
    classes = shapeChanged ? ['structural', 'substantiveText'] : ['substantiveText'];
  }
  return { ...op, classification: classes[0]!, classes };
}

/**
 * ADR-0046: a Word table pair carries its cells' own ops, each classified by
 * the body's stages (§7.3's table-cell floor included, since every cell
 * paragraph is `isTableCell`). The table op takes the UNION of those classes,
 * plus `structural` when a row or a cell was inserted, deleted or re-laid —
 * so it is filtered only when every change inside it would be.
 */
function classifyWordTableOp(op: DiffOp, diff: TableDiff): DiffOp {
  const found = new Set<ChangeClass>();
  for (const row of diff.rows) {
    if (row.kind !== 'match') {
      found.add('structural');
      continue;
    }
    for (const cell of row.cells) {
      if (cell.kind !== 'match') {
        found.add('structural');
        continue;
      }
      for (const inner of cell.ops) {
        if (inner.type === 'equal') continue;
        for (const c of inner.classes) found.add(c);
      }
    }
  }
  if (found.size === 0) found.add('substantiveText');
  const classes = PRIORITY_ORDER.filter((c) => found.has(c));
  return { ...op, classification: classes[0]!, classes };
}

/**
 * Classify a single DiffOp. Returns a NEW op with `classification` (primary,
 * highest priority) and `classes` (all matching, priority-ordered) populated.
 * Does not modify the input op. Pure — no side effects.
 *
 * `config` is accepted for signature symmetry with the SG/UK rules layer (B2);
 * the base taxonomy itself is config-independent.
 */
export function classifyOp(
  op: DiffOp,
  original: ParaRecord[],
  revised: ParaRecord[],
  _config: CompareConfig,
): DiffOp {
  // ADR-0027: table-unit ops (text-vs-text Markdown) are classified structurally;
  // the base taxonomy below is run/word-diff based and does not model tables.
  const origRec = op.origRange ? original[op.origRange.start] : undefined;
  const revRec = op.revRange ? revised[op.revRange.start] : undefined;
  if (origRec?.table || revRec?.table) {
    return classifyTableOp(op, origRec, revRec);
  }

  const classes = classifyClasses(op, original, revised);
  // classifyClasses always returns a non-empty list; the fallback satisfies the
  // type checker without ever being hit (and never drops the op — ADR-0008).
  const primary = classes[0] ?? 'substantiveText';
  return {
    ...op,
    classification: primary,
    classes,
  };
}

/**
 * Classify every op in the list. Returns a NEW array of the same length — one
 * classified op per input op (never adds or drops ops). Does not mutate the
 * input ops.
 */
export function classifyAll(
  ops: DiffOp[],
  original: ParaRecord[],
  revised: ParaRecord[],
  config: CompareConfig,
): DiffOp[] {
  return ops.map((op) => classifyOp(op, original, revised, config));
}
