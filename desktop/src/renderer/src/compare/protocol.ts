import type { RenderModel } from '@shared/render-core/types.js';
import type { ChangeSummary, DiffOp } from '@features/compare/model.js';
import type { PairingResult } from '../pairing/build-pairing.js';
import type { RedlineContext } from '../ai/context.js';
import type { StructureOutline } from '../structure/outline.js';
import type { AppliedEdit, ProposedEdit, RefusedEdit } from '../edits/apply-edits.js';

/**
 * The compare worker protocol.
 *
 * Conventions borrowed from the engine's own workers: a monotonic numeric `id`,
 * an explicit `cancel` message, and a four-member reply union.
 *
 * Cancellation is a MESSAGE, not an AbortSignal. The pipeline takes a duck-typed
 * `{ aborted: boolean }` rather than a real AbortSignal, which is lucky — a real
 * one is not structured-cloneable and could not cross this boundary at all. The
 * worker owns a mutable flag and the cancel message flips it.
 */

export interface CompareRequest {
  readonly type: 'compare';
  readonly id: number;
  /** Transferred, not copied — the buffers are detached on the sending side. */
  readonly originalBytes: ArrayBuffer;
  readonly revisedBytes: ArrayBuffer;
  readonly jurisdiction: 'UK' | 'SG';
  readonly author: string;
  /** ISO-8601. Injected by the caller so results are reproducible in tests. */
  readonly date: string;
  readonly originalName: string;
  readonly revisedName: string;
}

export interface CancelRequest {
  readonly type: 'cancel';
  readonly id: number;
}

/** Build a pairing proposal for a many-to-many bundle. */
export interface PairingRequest {
  readonly type: 'pair';
  readonly id: number;
  readonly docs: ReadonlyArray<{
    readonly docId: string;
    readonly side: 'a' | 'b';
    readonly fileName: string;
    readonly bytes: ArrayBuffer;
  }>;
  readonly jurisdiction: 'UK' | 'SG';
  /** ISO-8601, injected by the caller — never Date.now() inside the builder. */
  readonly generatedAt: string;
}

/**
 * Apply the assistant's edits to one document and redline the result against it.
 *
 * The two halves are ONE request because they are one idea: an AI edit is not
 * finished until the user can see it as a tracked change. Splitting them would
 * make it possible to produce an edited document that never got redlined.
 */
export interface SelfRedlineRequest {
  readonly type: 'self-redline';
  readonly id: number;
  /** The document being edited. Transferred. */
  readonly targetBytes: ArrayBuffer;
  readonly targetName: string;
  readonly edits: readonly ProposedEdit[];
  readonly jurisdiction: 'UK' | 'SG';
  /** The w:author stamped on every tracked change. Says who made the edit. */
  readonly author: string;
  readonly date: string;
}

export interface SelfRedlineOutcome {
  readonly outcome: CompareOutcome;
  readonly applied: readonly AppliedEdit[];
  readonly refused: readonly RefusedEdit[];
  /** The edited document, clean. The user can save it; the redline shows the changes. */
  readonly editedBytes: ArrayBuffer;
}

export interface SelfRedlineResultMessage {
  readonly type: 'self-redline-result';
  readonly id: number;
  readonly result: SelfRedlineOutcome;
}

export type ClientToWorker =
  | CompareRequest
  | PairingRequest
  | SelfRedlineRequest
  | CancelRequest;

export interface CompareOutcome {
  /** Flowed-prose render model. Plain data: arrays and Maps, clone-safe. */
  readonly model: RenderModel;
  readonly summary: ChangeSummary;
  readonly ops: readonly DiffOp[];
  /**
   * The merged body with native w:ins / w:del / w:moveFrom / w:moveTo. Carried
   * back because the .docx export and the phase-7 self-redline both need it,
   * and re-running the pipeline to recover it would be absurd.
   */
  readonly authoredBodyXml: string;
  readonly quarantinedOriginal: number;
  readonly quarantinedRevised: number;
  /**
   * The clause structure of each document, for the Structure panel and the
   * assistant's citations. null for a side the analyser could not read. Plain
   * JSON, clone-safe. Never used to filter or hide any change.
   */
  readonly structure: { readonly original: StructureOutline | null; readonly revised: StructureOutline | null };
  /**
   * What the assistant is given to read. Built in the worker beside the
   * pipeline; see ai/context.ts. Plain JSON, clone-safe.
   */
  readonly context: RedlineContext;
}

export interface ProgressMessage {
  readonly type: 'progress';
  readonly id: number;
  readonly pct: number;
  readonly step: string;
}
export interface ResultMessage {
  readonly type: 'result';
  readonly id: number;
  readonly outcome: CompareOutcome;
}
export interface ErrorMessage {
  readonly type: 'error';
  readonly id: number;
  readonly message: string;
}
export interface CancelledMessage {
  readonly type: 'cancelled';
  readonly id: number;
}

export interface PairingResultMessage {
  readonly type: 'pairing';
  readonly id: number;
  readonly result: PairingResult;
}

export type WorkerToClient =
  | ProgressMessage
  | ResultMessage
  | PairingResultMessage
  | SelfRedlineResultMessage
  | ErrorMessage
  | CancelledMessage;

/**
 * Human labels for the pipeline's 0-100 progress.
 *
 * The engine reports a bare percentage; a long batch needs words. The bands are
 * approximate by design — they describe what the user is waiting for, not an
 * exact internal stage boundary.
 */
export function stepFor(pct: number): string {
  if (pct < 10) return 'Reading documents';
  if (pct < 25) return 'Extracting paragraphs';
  if (pct < 55) return 'Aligning and diffing';
  if (pct < 75) return 'Classifying changes';
  if (pct < 90) return 'Authoring tracked changes';
  return 'Building the redline';
}
