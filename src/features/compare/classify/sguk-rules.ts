/**
 * Feature 08 — Compare · SG/UK forcing rules (Task B2).
 *
 * These rules OVERRIDE the base taxonomy (B1) for specific legal-significance
 * signals. They only ever ESCALATE a change toward substantive / never-filter;
 * they never relabel a change as filterable noise, and they never drop an op
 * (ADR-0008).
 *
 * Order of application (§7):
 *   1. en-US→en-GB spelling variant  → force substantiveText + spellingVariant.
 *   2. Signature / execution block    → force substantive, never filterable.
 *   3. Table-cell value change        → force numericalValue / substantiveText.
 *
 * Spec: spec/features/08-compare.md §7.1–§7.3.
 */

import type { CompareConfig } from '@shared/config/index.js';
import type { ChangeClass, DiffOp, ParaRecord, WordDiff } from '../model.js';
import { PRIORITY_ORDER } from './taxonomy.js';
import { detectSpellingVariant } from './spelling-variant.js';
import { applySignatureBlockRule } from './signature-block.js';

const NUMERIC_ATOMICS = new Set<string>(['amount', 'date', 'percent', 'duration']);

/** Non-equal (ins/del) word diffs of an op. */
function changedDiffs(op: DiffOp): WordDiff[] {
  return (op.wordDiffs ?? []).filter((wd) => wd.op !== 'eq');
}

/** Whether the op has any non-equal text delta. */
function hasTextDelta(op: DiffOp): boolean {
  return changedDiffs(op).some((wd) => wd.text.length > 0);
}

/** Whether any paragraph this op touches is a table cell. */
function touchesTableCell(
  op: DiffOp,
  original: ParaRecord[],
  revised: ParaRecord[],
): boolean {
  if (op.origRange) {
    for (let i = op.origRange.start; i <= op.origRange.end; i++) {
      if (original[i]?.isTableCell) return true;
    }
  }
  if (op.revRange) {
    for (let i = op.revRange.start; i <= op.revRange.end; i++) {
      if (revised[i]?.isTableCell) return true;
    }
  }
  return false;
}

/** Whether any changed atomic token is a numerical value. */
function hasNumericAtomicChange(op: DiffOp): boolean {
  return changedDiffs(op).some(
    (wd) => wd.atomic !== undefined && NUMERIC_ATOMICS.has(wd.atomic),
  );
}

/** Build a priority-ordered class list with `primary` guaranteed first. */
function withPrimary(primary: ChangeClass, classes: Iterable<ChangeClass>): ChangeClass[] {
  const set = new Set<ChangeClass>(classes);
  set.add(primary);
  // Primary first, then the rest in canonical priority order.
  const rest = PRIORITY_ORDER.filter((c) => c !== primary && set.has(c));
  return [primary, ...rest];
}

/**
 * Apply the SG/UK forcing rules to an already-classified op (B1 output).
 * Returns a NEW op. Pure — no side effects.
 */
export function applySgUkRules(
  op: DiffOp,
  original: ParaRecord[],
  revised: ParaRecord[],
  config: CompareConfig,
): DiffOp {
  let result: DiffOp = { ...op, classes: [...op.classes] };

  // 1. en-US→en-GB spelling variant — force substantiveText, tag the sub-flag,
  //    never allow formatting/whitespace to win.
  const variant = detectSpellingVariant(result, config);
  if (variant) {
    const classes = new Set<ChangeClass>(result.classes);
    classes.delete('formatting');
    classes.delete('whitespacePunct');
    result = {
      ...result,
      classification: 'substantiveText',
      classes: withPrimary('substantiveText', classes),
      spellingVariant: { usSpelling: variant.usSpelling, gbSpelling: variant.gbSpelling },
    };
  }

  // 2. Signature / execution block — force substantive, never filterable.
  result = applySignatureBlockRule(result, original, revised);

  // 3. Table-cell value change with a real text delta — force at least
  //    numericalValue (atomic value change) or substantiveText; never
  //    formatting / whitespace.
  if (touchesTableCell(result, original, revised) && hasTextDelta(result)) {
    const forced: ChangeClass = hasNumericAtomicChange(result)
      ? 'numericalValue'
      : 'substantiveText';
    // Only escalate: do not demote a higher-priority primary (e.g. structural
    // or obligationLanguage) that already outranks the forced floor.
    const classes = new Set<ChangeClass>(result.classes);
    classes.delete('formatting');
    classes.delete('whitespacePunct');
    classes.add(forced);
    const currentRank = PRIORITY_ORDER.indexOf(result.classification);
    const forcedRank = PRIORITY_ORDER.indexOf(forced);
    const primary =
      currentRank >= 0 && currentRank < forcedRank ? result.classification : forced;
    result = {
      ...result,
      classification: primary,
      classes: withPrimary(primary, classes),
    };
  }

  return result;
}
