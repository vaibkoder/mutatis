/**
 * UI contract for the batch-compare panel (ADR-0037; spec §5.4).
 *
 * Same shape as `compare/ui-types.ts`: signals in, actions out, with a stub
 * factory so a panel can be rendered in isolation. The stub exists for the same
 * reason the single-pair one does — and carries the same hazard, recorded here
 * because the project has already been bitten by it once
 * (`web/WebCompare.tsx:15-22`): a panel rendered WITHOUT an injected live
 * driver falls back to no-ops and silently does nothing, and unit tests that
 * render the panel do not notice. Any surface mounting this panel must inject
 * the real driver, and must have a test asserting that it did.
 */
import type { Signal } from '@preact/signals';

import type { ChangeSummary, DiffOp } from '../model.js';
import type { CompareStatus, FlattenChoice } from '../ui-types.js';
import type { BatchIndex, BatchRow } from './model.js';
import { MAX_BATCH_VARIANTS } from './limits.js';

/**
 * One pair's full change set, recovered on demand.
 *
 * The batch deliberately drops this the moment a pair completes (spec §4.3
 * step 3) — two document-sized `ParaRecord[]` per pair is the thing that breaks
 * first at N. Opening a row re-runs that single pair to get it back, which is
 * bounded, fast, and correct.
 */
export interface PairDetail {
  rowId: string;
  label: string;
  ops: DiffOp[];
  summary: ChangeSummary;
}

/** One file queued for a batch, before it is read. */
export interface BatchSlot {
  id: string;
  label: string;
  byteLength: number;
}

export interface BatchDriverState {
  status: Signal<CompareStatus>;
  /** 0–100 across the whole batch (spec §5.5). */
  progress: Signal<number>;
  errorMessage: Signal<string | null>;

  /** The template slot. `null` until one is chosen. */
  template: Signal<BatchSlot | null>;
  /** Queued variants, in the order they will run. */
  variants: Signal<readonly BatchSlot[]>;

  /**
   * Live rows, updated as pairs finish. The full set of pairs the batch was
   * asked to run — a failure is a row with a reason, never a missing row.
   */
  rows: Signal<readonly BatchRow[]>;
  /** The finished index. `null` until a run completes or is cancelled. */
  index: Signal<BatchIndex | null>;

  /** Per-side flatten choice, applied to the whole batch (spec §8 case 7). */
  flattenTemplate: Signal<FlattenChoice>;
  flattenVariant: Signal<FlattenChoice>;
  /**
   * The variant cap this driver will actually enforce, resolved from
   * `compare.batch.maxVariants`. On the state rather than imported by the panel
   * so a profile that tightens the cap cannot leave the panel counting up to a
   * different number than the runner refuses at.
   */
  maxVariants: number;

  setTemplate(file: File): void;
  addVariants(files: readonly File[]): void;
  removeVariant(id: string): void;
  clearVariants(): void;

  runBatch(): Promise<void>;
  cancelBatch(): void;

  /**
   * Re-run one pair to recover its full change set.
   *
   * Deliberately NOT `compareDriverState.runCompare()`: that path materialises
   * the redline, which on the web target means triggering a file download — an
   * absurd side effect for "show me what changed". This re-runs the pipeline
   * only. Resolves `null` if the row cannot be re-run.
   */
  openPair(id: string): Promise<PairDetail | null>;

  /** Download one row's artefact. */
  downloadRow(id: string): Promise<void>;
  /** Download every completed row's artefact as one archive. */
  downloadAll(): Promise<void>;
}

/**
 * A no-op driver for rendering the panel in isolation.
 *
 * Every action is deliberately inert. See the module comment: injecting the
 * real driver is the mounting surface's job, and forgetting to is a silent
 * failure, not a loud one.
 */
export function makeStubBatchDriverState(
  signal: <T>(value: T) => Signal<T>,
): BatchDriverState {
  return {
    status: signal<CompareStatus>('idle'),
    progress: signal(0),
    errorMessage: signal<string | null>(null),
    template: signal<BatchSlot | null>(null),
    variants: signal<readonly BatchSlot[]>([]),
    rows: signal<readonly BatchRow[]>([]),
    index: signal<BatchIndex | null>(null),
    flattenTemplate: signal<FlattenChoice>('acceptAll'),
    flattenVariant: signal<FlattenChoice>('acceptAll'),
    maxVariants: MAX_BATCH_VARIANTS,
    setTemplate: () => {},
    addVariants: () => {},
    removeVariant: () => {},
    clearVariants: () => {},
    runBatch: () => Promise.resolve(),
    cancelBatch: () => {},
    openPair: () => Promise.resolve(null),
    downloadRow: () => Promise.resolve(),
    downloadAll: () => Promise.resolve(),
  };
}
