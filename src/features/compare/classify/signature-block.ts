/**
 * Feature 08 — Compare · signature/execution-block forcing rule (Task B2).
 *
 * Any change inside a paragraph flagged `isSignatureBlock` carries legal weight
 * disproportionate to its text size (UK deed execution blocks; SG common-seal /
 * director-secretary signing). Such a change is ALWAYS forced to substantive
 * and is NEVER filterable as formatting / whitespace (§7.2).
 *
 * Spec: spec/features/08-compare.md §7.2.
 */

import type { ChangeClass, DiffOp, ParaRecord } from '../model.js';

/** Whether any paragraph this op touches is a signature/execution block. */
function touchesSignatureBlock(
  op: DiffOp,
  original: ParaRecord[],
  revised: ParaRecord[],
): boolean {
  if (op.origRange) {
    for (let i = op.origRange.start; i <= op.origRange.end; i++) {
      if (original[i]?.isSignatureBlock) return true;
    }
  }
  if (op.revRange) {
    for (let i = op.revRange.start; i <= op.revRange.end; i++) {
      if (revised[i]?.isSignatureBlock) return true;
    }
  }
  return false;
}

/**
 * Force a signature/execution-block op to substantive. Returns a NEW op (does
 * not mutate the input). If the op touches no signature block, it is returned
 * unchanged (a fresh copy).
 *
 * - `classification` is forced to `structural` (for insert/delete/move) or
 *   `substantiveText` otherwise.
 * - `substantiveText` is guaranteed present in `classes`.
 * - `spellingVariant` is cleared — a signature block is never reduced to a
 *   filterable house-style/whitespace/formatting item.
 */
export function applySignatureBlockRule(
  op: DiffOp,
  original: ParaRecord[],
  revised: ParaRecord[],
): DiffOp {
  if (!touchesSignatureBlock(op, original, revised)) {
    return { ...op, classes: [...op.classes] };
  }

  const isStructural = op.type === 'insert' || op.type === 'delete' || op.type === 'move';
  const forced: ChangeClass = isStructural ? 'structural' : 'substantiveText';

  const classes = new Set<ChangeClass>(op.classes);
  classes.add('substantiveText');
  if (isStructural) classes.add('structural');
  // Never filterable as noise: drop formatting / whitespace labels.
  classes.delete('formatting');
  classes.delete('whitespacePunct');

  // Drop `spellingVariant` entirely (never a filterable house-style item).
  // Under exactOptionalPropertyTypes the key must be omitted, not set to
  // undefined — so rebuild the op without it.
  const { spellingVariant: _dropped, ...rest } = op;
  void _dropped;
  return {
    ...rest,
    classification: forced,
    classes: orderByPrimary(forced, classes),
  };
}

/** Put the forced primary class first; keep the remainder in a stable order. */
function orderByPrimary(primary: ChangeClass, classes: Set<ChangeClass>): ChangeClass[] {
  const rest = [...classes].filter((c) => c !== primary);
  return [primary, ...rest];
}
