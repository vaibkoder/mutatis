/**
 * Compare at scale — data model (ADR-0037; spec §5.2).
 *
 * Plain data only: every type here must survive `structuredClone`, because rows
 * cross a worker boundary and artefacts are browser-managed `Blob`s rather than
 * JS-heap byte arrays (spec §4.3).
 *
 * The load-bearing property is that a pair the batch was ASKED to run always
 * produces a row. Failed, cancelled and unrunnable are ROWS WITH REASONS, never
 * absences — a batch index that quietly omits what went wrong is the batch-scale
 * version of hiding a change (ADR-0008).
 *
 * Pure — no Office JS, no DOM, no network.
 */
import type { FlattenPolicy } from '@shared/config/index.js';
import type { ChangeClass } from '../model.js';

/** Which product surface produced this batch. */
export type BatchMode = 'one-to-many' | 'many-to-many';

/**
 * Where a pair got to.
 *
 * - `failed` — this pair errored (a corrupt file, a bad package). The batch
 *   continues; one bad file must not cost the user ninety-nine good redlines.
 * - `cancelled` — never ran, or was in flight, when the user stopped the batch.
 * - `unrunnable` — the pair is real and identified, but the comparison is out of
 *   scope (today: a cross-format pair — ADR-0017's Phase-E alignment gate was
 *   never measured). Surfaced, never silently skipped.
 */
export type BatchRowStatus =
  | 'queued'
  | 'running'
  | 'done'
  | 'failed'
  | 'cancelled'
  | 'unrunnable';

/**
 * A materialised artefact. The bytes live in a browser-managed `Blob` — which
 * the engine may spill to disk — so retaining N of these is sound in a way that
 * retaining N `PipelineResult`s is not.
 */
export interface ArtefactHandle {
  kind: 'redlineDocx' | 'snapshotPdf';
  /** Suggested download name, already sanitised and de-duplicated. */
  fileName: string;
  byteLength: number;
  blob: Blob;
}

/** One comparison in a batch. */
export interface BatchRow {
  /** Stable within the batch. Rows are keyed by this, never by file name. */
  id: string;
  /** Display label: the variant's file name (1:N) or `"A ▸ B"` (M:N). */
  label: string;
  status: BatchRowStatus;
  /** Required when status is `failed` or `unrunnable`; says what and why. */
  reason?: string;
  /**
   * The pair's TOTAL change count — `summary.total`, never a filtered count.
   * On-screen filters are navigation; they do not reach this number (ADR-0008).
   */
  changeCount: number;
  byClassification: Readonly<Record<ChangeClass, number>>;
  quarantinedOriginal: number;
  quarantinedRevised: number;
  artefacts: readonly ArtefactHandle[];
  /**
   * Set when the cumulative artefact ceiling was reached before this row, so
   * its artefacts are produced by re-running the pair on download. Always
   * accompanied by a batch-level disclosure (spec §4.3 step 6).
   */
  artefactOnDemand: boolean;
  durationMs: number;
}

/** The per-batch flatten choice, one per side. */
export interface BatchFlatten {
  template: FlattenPolicy;
  variant: FlattenPolicy;
}

/** Batch-level tallies. Derived, never authored by hand. */
export interface BatchTotals {
  pairs: number;
  done: number;
  failed: number;
  cancelled: number;
  unrunnable: number;
  /** Summed `changeCount` across completed pairs. */
  changes: number;
}

/**
 * The flat, filter-independent index of a batch — the analogue of
 * `export/report-model.ts → CompareReport`, one level up.
 */
export interface BatchIndex {
  kind: 'compare-batch';
  mode: BatchMode;
  /** The template's label. One-to-many only. */
  templateLabel?: string;
  /** Injected by the caller — NEVER `Date.now()` inside the builder. */
  generatedAt: string;
  rows: readonly BatchRow[];
  totals: BatchTotals;
  /**
   * The flatten choice every row in this batch was produced under.
   *
   * Per batch, not per pair (spec §8 case 7). It is on the index because
   * "accept all tracked changes first" materially changes what a redline says,
   * and a reader who cannot see which setting produced a hundred redlines
   * cannot judge any of them.
   */
  flatten: BatchFlatten;
  /** Blind-spot notices carried with the batch, exactly as the per-pair reports do. */
  disclosures: readonly string[];
}

/** Every `ChangeClass`, zeroed — the starting point for a row's tally. */
export function emptyClassificationTally(): Record<ChangeClass, number> {
  return {
    structural: 0,
    obligationLanguage: 0,
    numericalValue: 0,
    definedTerm: 0,
    crossRefRenumber: 0,
    substantiveText: 0,
    whitespacePunct: 0,
    formatting: 0,
  };
}
