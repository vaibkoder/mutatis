/**
 * Main-thread batch-compare pool client (ADR-0037; spec §4.2) — host-free.
 *
 * RenderCore ships one shared worker addressed by job id; there is no pool
 * anywhere in the repo. A batch needs one, because its jobs are independent and
 * CPU-bound and there are up to a hundred of them.
 *
 * The `WorkerLike` structural type is carried over deliberately: it is why the
 * worker path is unit-testable without jsdom `Worker` support, and a pool with
 * no test seam would be a pool nobody can prove correct.
 *
 * Pool discipline, and the reason for each rule:
 *
 *  - **Never more than one in-flight job per worker.** Each job holds two
 *    document-sized `ParaRecord[]`; two jobs on one worker would double a
 *    worker's peak footprint for no throughput, since the work is CPU-bound.
 *  - **A queued job that is cancelled is never posted.** It settles locally, so
 *    a cancelled batch stops costing anything immediately rather than draining
 *    its queue through the workers first.
 *  - **A worker is returned to the free list on every terminal outcome** —
 *    result, error and cancel alike. A pool that leaks a worker on the error
 *    path deadlocks on the batch after the one that failed.
 *
 * The client touches no DOM and no Office API; it only marshals messages.
 */
import type { CompareConfig } from '@shared/config/index.js';
import type { ParaRecord } from '../../model.js';
import type {
  CancelledMessage,
  ClientToWorker,
  CompareJobOutcome,
  CompareJobRequest,
  ErrorMessage,
  ProgressMessage,
  ResultMessage,
  WorkerToClient,
} from './protocol.js';

/** The subset of `Worker` the pool needs (the real `Worker` satisfies it). */
export interface WorkerLike {
  postMessage(message: ClientToWorker, transfer?: Transferable[]): void;
  onmessage: ((ev: { data: WorkerToClient }) => void) | null;
  terminate?(): void;
}

/** One pair to compare. `originalParas` is shared across a one-to-many batch. */
export interface CompareJobInput {
  originalParas: ParaRecord[];
  revisedParas: ParaRecord[];
  config: CompareConfig;
  author: string;
  date: string;
  quarantinedOriginal?: number;
  quarantinedRevised?: number;
}

export interface CompareJobOptions {
  onProgress?: (pct: number) => void;
}

export interface CompareJob {
  id: number;
  /** Resolves with the outcome; rejects with an `AbortError` if cancelled. */
  outcome: Promise<CompareJobOutcome>;
}

export interface BatchPool {
  /** Enqueue a pair. Runs as soon as a worker is free. */
  submit(input: CompareJobInput, opts?: CompareJobOptions): CompareJob;
  /** Cancel one job, queued or running. */
  cancel(id: number): void;
  /** Cancel every queued and running job. */
  cancelAll(): void;
  /** How many workers back this pool. 1 for the in-process fallback. */
  readonly size: number;
  terminate(): void;
}

interface Pending {
  input: CompareJobInput;
  onProgress: ((pct: number) => void) | undefined;
  resolve(outcome: CompareJobOutcome): void;
  reject(err: Error): void;
  /** The worker running it, or null while queued. */
  worker: WorkerLike | null;
}

function abortError(): DOMException {
  return new DOMException('Compare cancelled', 'AbortError');
}

/**
 * Build a pool over the given workers.
 *
 * Passing one worker gives a correct, sequential pool — which is exactly what
 * the in-process fallback needs, so there is no second code path for it.
 */
export function createBatchPool(workers: readonly WorkerLike[]): BatchPool {
  if (workers.length === 0) throw new Error('A batch pool needs at least one worker.');

  let nextId = 1;
  const pending = new Map<number, Pending>();
  const queue: number[] = [];
  const free: WorkerLike[] = [...workers];

  const release = (worker: WorkerLike | null): void => {
    if (worker && !free.includes(worker)) free.push(worker);
    pump();
  };

  const settle = (id: number, apply: (job: Pending) => void): void => {
    const job = pending.get(id);
    if (!job) return;
    pending.delete(id);
    const { worker } = job;
    apply(job);
    release(worker);
  };

  function pump(): void {
    while (free.length > 0 && queue.length > 0) {
      const id = queue.shift() as number;
      const job = pending.get(id);
      // Cancelled while queued: it was settled locally and never posted.
      if (!job) continue;
      const worker = free.pop() as WorkerLike;
      job.worker = worker;
      const request: CompareJobRequest = {
        type: 'compare',
        id,
        originalParas: job.input.originalParas,
        revisedParas: job.input.revisedParas,
        config: job.input.config,
        author: job.input.author,
        date: job.input.date,
        quarantinedOriginal: job.input.quarantinedOriginal ?? 0,
        quarantinedRevised: job.input.quarantinedRevised ?? 0,
      };
      worker.postMessage(request);
    }
  }

  for (const worker of workers) {
    worker.onmessage = (ev): void => {
      const m: WorkerToClient = ev.data;
      switch (m.type) {
        case 'progress': {
          pending.get((m as ProgressMessage).id)?.onProgress?.((m as ProgressMessage).pct);
          break;
        }
        case 'result':
          settle(m.id, job => { job.resolve((m as ResultMessage).outcome); });
          break;
        case 'cancelled':
          settle((m as CancelledMessage).id, job => { job.reject(abortError()); });
          break;
        case 'error':
          settle(m.id, job => { job.reject(new Error((m as ErrorMessage).message)); });
          break;
      }
    };
  }

  /**
   * Cancel one job. A free function rather than a method so `cancelAll` does
   * not depend on `this` — the pool is passed around and destructured.
   */
  const cancelJob = (id: number): void => {
    const job = pending.get(id);
    if (!job) return;
    if (job.worker === null) {
      // Still queued — settle it here and never pay to post it.
      pending.delete(id);
      job.reject(abortError());
      return;
    }
    // Running — the worker flips its abort token and posts `cancelled`, which
    // is what actually settles the promise and frees the worker.
    job.worker.postMessage({ type: 'cancel', id });
  };

  return {
    size: workers.length,

    submit(input, opts = {}): CompareJob {
      const id = nextId++;
      const outcome = new Promise<CompareJobOutcome>((resolve, reject) => {
        pending.set(id, {
          input,
          onProgress: opts.onProgress,
          resolve,
          reject,
          worker: null,
        });
        queue.push(id);
      });
      pump();
      return { id, outcome };
    },

    cancel(id): void {
      cancelJob(id);
    },

    cancelAll(): void {
      // `cancelJob`, not `this.cancel` — a destructured `cancelAll` would throw.
      for (const id of [...pending.keys()]) cancelJob(id);
    },

    terminate(): void {
      for (const id of [...pending.keys()]) {
        const job = pending.get(id);
        pending.delete(id);
        job?.reject(abortError());
      }
      queue.length = 0;
      for (const worker of workers) worker.terminate?.();
    },
  };
}
