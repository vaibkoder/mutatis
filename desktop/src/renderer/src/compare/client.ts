import type {
  CompareOutcome,
  ClientToWorker,
  SelfRedlineOutcome,
  WorkerToClient,
} from './protocol.js';
import type { ProposedEdit } from '../edits/apply-edits.js';
import type { PairingResult } from '../pairing/build-pairing.js';
import { createCompareWorker } from './create-worker.js';

export interface CompareJobOptions {
  readonly originalBytes: ArrayBuffer;
  readonly revisedBytes: ArrayBuffer;
  readonly jurisdiction?: 'UK' | 'SG';
  readonly author?: string;
  readonly date?: string;
  /** Shown to the model as the baseline/revised labels in the context payload. */
  readonly originalName?: string;
  readonly revisedName?: string;
  readonly onProgress?: (pct: number, step: string) => void;
}

export interface CompareJob {
  readonly id: number;
  readonly outcome: Promise<CompareOutcome>;
}

export interface PairingJobOptions {
  readonly docs: ReadonlyArray<{
    readonly docId: string;
    readonly side: 'a' | 'b';
    readonly fileName: string;
    readonly bytes: ArrayBuffer;
  }>;
  readonly jurisdiction?: 'UK' | 'SG';
  readonly generatedAt?: string;
}

export interface SelfRedlineOptions {
  readonly targetBytes: ArrayBuffer;
  readonly targetName: string;
  readonly edits: readonly ProposedEdit[];
  readonly jurisdiction?: 'UK' | 'SG';
  readonly author?: string;
  readonly date?: string;
  readonly onProgress?: (pct: number, step: string) => void;
}

export interface SelfRedlineJob {
  readonly id: number;
  readonly result: Promise<SelfRedlineOutcome>;
}

export interface CompareClient {
  compare(options: CompareJobOptions): CompareJob;
  proposePairing(options: PairingJobOptions): Promise<PairingResult>;
  selfRedline(options: SelfRedlineOptions): SelfRedlineJob;
  cancel(id: number): void;
  terminate(): void;
}

/** Thrown when a job is cancelled, so callers can tell it from a real failure. */
export class CompareCancelledError extends Error {
  constructor() {
    super('Comparison cancelled');
    this.name = 'CompareCancelledError';
  }
}

interface Pending {
  resolve: (value: never) => void;
  reject: (error: Error) => void;
  onProgress: ((pct: number, step: string) => void) | undefined;
}

/**
 * Client for the compare worker.
 *
 * One worker serves every job; ids keep them apart. That is deliberate — a
 * worker per comparison would re-pay module init (and the XML parser) on every
 * run, which dominates the cost of a small document.
 */
export function createCompareClient(): CompareClient {
  const worker = createCompareWorker();

  const pending = new Map<number, Pending>();
  let nextId = 0;

  worker.onmessage = (event: MessageEvent<WorkerToClient>) => {
    const message = event.data;
    const entry = pending.get(message.id);
    if (!entry) return;

    switch (message.type) {
      case 'progress':
        entry.onProgress?.(message.pct, message.step);
        break;
      case 'result':
        pending.delete(message.id);
        (entry.resolve as (v: CompareOutcome) => void)(message.outcome);
        break;
      case 'pairing':
        pending.delete(message.id);
        (entry.resolve as (v: PairingResult) => void)(message.result);
        break;
      case 'self-redline-result':
        pending.delete(message.id);
        (entry.resolve as (v: SelfRedlineOutcome) => void)(message.result);
        break;
      case 'error':
        pending.delete(message.id);
        entry.reject(new Error(message.message));
        break;
      case 'cancelled':
        pending.delete(message.id);
        entry.reject(new CompareCancelledError());
        break;
    }
  };

  // A worker-level error never lands on a specific job, so it fails all of them
  // rather than hanging every caller forever.
  worker.onerror = (event) => {
    const error = new Error(
      `Compare worker failed: ${event.message || 'unknown error'}`,
    );
    for (const [, entry] of pending) entry.reject(error);
    pending.clear();
  };

  function send(message: ClientToWorker, transfer?: Transferable[]): void {
    if (transfer) worker.postMessage(message, transfer);
    else worker.postMessage(message);
  }

  return {
    compare(options) {
      const id = ++nextId;
      const outcome = new Promise<CompareOutcome>((resolve, reject) => {
        pending.set(id, { resolve, reject, onProgress: options.onProgress });
      });

      send(
        {
          type: 'compare',
          id,
          originalBytes: options.originalBytes,
          revisedBytes: options.revisedBytes,
          jurisdiction: options.jurisdiction ?? 'UK',
          author: options.author ?? 'Mutatis Desktop',
          date: options.date ?? new Date().toISOString(),
          originalName: options.originalName ?? 'Original',
          revisedName: options.revisedName ?? 'Modified',
        },
        // Transfer rather than copy. The buffers are detached here afterwards,
        // which is why callers must read each File once per job.
        [options.originalBytes, options.revisedBytes],
      );

      return { id, outcome };
    },

    selfRedline(options) {
      const id = ++nextId;
      const result = new Promise<SelfRedlineOutcome>((resolve, reject) => {
        pending.set(id, {
          resolve: resolve as (v: never) => void,
          reject,
          onProgress: options.onProgress,
        });
      });

      send(
        {
          type: 'self-redline',
          id,
          targetBytes: options.targetBytes,
          targetName: options.targetName,
          edits: options.edits,
          jurisdiction: options.jurisdiction ?? 'UK',
          // The author is stamped into every w:ins and w:del, so a reviewer
          // opening the .docx in Word sees who proposed the change.
          author: options.author ?? 'Mutatis assistant',
          date: options.date ?? new Date().toISOString(),
        },
        [options.targetBytes],
      );

      return { id, result };
    },

    proposePairing(options) {
      const id = ++nextId;
      const result = new Promise<PairingResult>((resolve, reject) => {
        pending.set(id, {
          resolve: resolve as (v: never) => void,
          reject,
          onProgress: undefined,
        });
      });

      send(
        {
          type: 'pair',
          id,
          docs: options.docs,
          jurisdiction: options.jurisdiction ?? 'UK',
          generatedAt: options.generatedAt ?? new Date().toISOString(),
        },
        options.docs.map((d) => d.bytes),
      );

      return result;
    },

    cancel(id) {
      send({ type: 'cancel', id });
    },

    terminate() {
      for (const [, entry] of pending) entry.reject(new CompareCancelledError());
      pending.clear();
      worker.terminate();
    },
  };
}
