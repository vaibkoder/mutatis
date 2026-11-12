/**
 * Compare at scale — the batch index builder (ADR-0037; spec §5.2).
 *
 * The direct analogue of `export/report-model.ts → buildCompareReport`, one
 * level up, and it inherits that module's two rules verbatim:
 *
 *  - **It takes no filter flags.** On-screen filters are navigation. A count
 *    here is always the true count (ADR-0008). There is no parameter you could
 *    pass to make this builder report less than happened.
 *  - **The date is injected**, never read from the clock inside, so the builder
 *    stays pure and its output is testable byte-for-byte.
 *
 * One further rule of its own: `rows.length` always equals the number of pairs
 * the batch was asked to run. A pair that failed, was cancelled, or cannot be
 * compared is a ROW WITH A REASON, never a missing row.
 *
 * Pure — no Office JS, no DOM, no network.
 */
import type {
  BatchFlatten,
  BatchIndex,
  BatchMode,
  BatchRow,
  BatchTotals,
} from './model.js';

export interface BuildBatchIndexInput {
  mode: BatchMode;
  /** The template's label. One-to-many only. */
  templateLabel?: string;
  /** Every pair the batch was asked to run, in run order. */
  rows: readonly BatchRow[];
  /** ISO-8601, injected by the caller. */
  generatedAt: string;
  /**
   * The flatten choice the batch ran under. Required, not optional: spec §8
   * case 7 says it is stated ON the index, and an optional field is one a
   * caller can forget.
   */
  flatten: BatchFlatten;
  disclosures?: readonly string[];
}

/** Tally the rows. Derived from the rows themselves so it cannot disagree with them. */
export function tallyBatch(rows: readonly BatchRow[]): BatchTotals {
  const totals: BatchTotals = {
    pairs: rows.length,
    done: 0,
    failed: 0,
    cancelled: 0,
    unrunnable: 0,
    changes: 0,
  };
  for (const row of rows) {
    switch (row.status) {
      case 'done':
        totals.done += 1;
        totals.changes += row.changeCount;
        break;
      case 'failed':
        totals.failed += 1;
        break;
      case 'cancelled':
        totals.cancelled += 1;
        break;
      case 'unrunnable':
        totals.unrunnable += 1;
        break;
      // 'queued' and 'running' are in-flight states; they are counted in
      // `pairs` and nowhere else — an index built mid-run is still honest
      // about how many pairs exist.
      default:
        break;
    }
  }
  return totals;
}

/**
 * Build the batch index from the rows a run produced.
 *
 * Note the signature: `rows` in, index out, and **no filter parameter**. That
 * absence is the ADR-0008 wall for this module, the same way
 * `buildCompareReport` has no filter parameter.
 */
export function buildBatchIndex(input: BuildBatchIndexInput): BatchIndex {
  const index: BatchIndex = {
    kind: 'compare-batch',
    mode: input.mode,
    generatedAt: input.generatedAt,
    rows: input.rows,
    totals: tallyBatch(input.rows),
    flatten: input.flatten,
    disclosures: input.disclosures ?? [],
  };
  if (input.templateLabel !== undefined) index.templateLabel = input.templateLabel;
  return index;
}
