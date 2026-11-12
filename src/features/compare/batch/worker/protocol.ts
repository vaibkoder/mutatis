/**
 * Batch-compare worker protocol (ADR-0037; spec §4.2) — host-free.
 *
 * The same shape as the RenderCore transport: two discriminated unions, every
 * message carrying a job `id`, everything structured-clone-safe. `ParaRecord`
 * is plain data (strings and numbers) and `CompareConfig` is resolved
 * configuration, so a job crosses the boundary by clone. Nothing here is
 * transferable — `ParaRecord[]` is an object graph, not a buffer — so unlike
 * RenderCore's font bytes there is no transfer path to reason about.
 *
 * One difference from RenderCore worth stating, because it changes what
 * cancellation means: `runComparePipeline` awaits between word-diff chunks, so
 * a cancel that arrives mid-run is genuinely observed part-way through the
 * heaviest stage. The render handler can only check at stage boundaries.
 *
 * No `Office.*`/`Word.*`, no `fetch` — records in, diff out (ADR-0001/0003).
 */
import type { CompareConfig } from '@shared/config/index.js';
import type { ChangeClass, ChangeSummary, DiffOp, ParaRecord } from '../../model.js';

/** Run one pair through the shipped Compare pipeline. */
export interface CompareJobRequest {
  type: 'compare';
  id: number;
  /**
   * The template/original side. In a one-to-many batch this is the SAME array
   * for every job — preprocessed once and reused (spec §4.3 step 2). Structured
   * clone copies it per post, which is the price of the worker boundary; the
   * saving that matters is not re-parsing and re-normalising the package.
   */
  originalParas: ParaRecord[];
  revisedParas: ParaRecord[];
  config: CompareConfig;
  author: string;
  /** ISO-8601, injected by the caller so a batch's redlines share one date. */
  date: string;
  quarantinedOriginal?: number;
  quarantinedRevised?: number;
}

export interface CancelRequest {
  type: 'cancel';
  id: number;
}

export type ClientToWorker = CompareJobRequest | CancelRequest;

/**
 * What comes back from a completed pair.
 *
 * Deliberately NOT the whole `PipelineResult`. `originalParas` and
 * `revisedParas` are document-sized on both sides and the runner drops them
 * immediately anyway (spec §4.3 step 3), so shipping them across the boundary
 * would pay a full structured-clone cost for data with a lifetime of
 * microseconds. Drill-down re-runs the pair instead.
 */
export interface CompareJobOutcome {
  ops: DiffOp[];
  summary: ChangeSummary;
  authoredBodyXml: string;
  classByRevId: ReadonlyMap<string, readonly ChangeClass[]>;
  quarantinedOriginal: number;
  quarantinedRevised: number;
}

export interface ProgressMessage { type: 'progress'; id: number; pct: number; }
export interface ResultMessage { type: 'result'; id: number; outcome: CompareJobOutcome; }
export interface ErrorMessage { type: 'error'; id: number; message: string; }
export interface CancelledMessage { type: 'cancelled'; id: number; }

export type WorkerToClient =
  | ProgressMessage
  | ResultMessage
  | ErrorMessage
  | CancelledMessage;
