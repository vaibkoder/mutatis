/**
 * Batch pool construction and the fallback ladder (ADR-0037; spec §4.2).
 *
 * Mirrors `compare/render/index.ts`'s ladder, for the same three reasons it
 * exists there:
 *
 *  1. the standalone single-file build sets `VITE_DP_INPROCESS`, and the define
 *     makes the branch statically true so the Worker chunk tree-shakes out;
 *  2. a page opened over `file://` cannot construct a module Worker at all;
 *  3. a batch must still RUN — slower — wherever a single compare runs. Falling
 *     back is not a degraded mode with different results; it is the same pool
 *     with one in-process worker.
 *
 * That last point is why the in-process path is a `WorkerLike` rather than a
 * second implementation of `BatchPool`: one code path, one set of semantics, and
 * a pooled run and an in-process run are directly comparable (acceptance B9).
 */
import { createBatchWorkerHandler } from './handler.js';
import { createBatchPool, type BatchPool, type WorkerLike } from './client.js';
import type { ClientToWorker, WorkerToClient } from './protocol.js';
import { resolvePoolSize } from '../constants.js';

export type { BatchPool, WorkerLike, CompareJob, CompareJobInput, CompareJobOptions } from './client.js';
export { createBatchPool } from './client.js';
export { createBatchWorkerHandler } from './handler.js';
export type * from './protocol.js';

/**
 * A `WorkerLike` that runs the handler in-process.
 *
 * Messages are delivered asynchronously (a resolved-promise turn) rather than
 * by direct call, so the handler yields to the message loop exactly as a real
 * worker does — which is what lets a `cancel` posted during a run actually be
 * observed part-way through, instead of only after the job finishes.
 */
export function createInProcessWorker(): WorkerLike {
  const fake: WorkerLike = {
    onmessage: null,
    postMessage(message: ClientToWorker): void {
      void Promise.resolve().then(() => handler.handle(message));
    },
    terminate(): void { /* nothing to tear down */ },
  };
  const handler = createBatchWorkerHandler({
    post: (m: WorkerToClient) => { fake.onmessage?.({ data: m }); },
  });
  return fake;
}

/** Build one module Worker, or null where the host forbids it. */
function tryCreateWorker(): WorkerLike | null {
  if ((import.meta as { env?: Record<string, string | undefined> }).env?.VITE_DP_INPROCESS === '1') {
    return null;
  }
  try {
    const worker = new Worker(
      new URL('../../../../workers/compare.worker.ts', import.meta.url),
      { type: 'module' },
    );
    // The DOM `Worker`'s `onmessage` is typed against `MessageEvent`; `WorkerLike`
    // narrows the message data. Structurally compatible at runtime — the protocol
    // is the contract — so bridge them through the seam type.
    return worker as unknown as WorkerLike;
  } catch {
    return null;
  }
}

export interface CreatePoolOptions {
  /** Force a size, bounded by `BATCH_POOL_MAX`. `0` (default) auto-sizes. */
  poolSize?: number;
  /** Injected by tests. When given, no real Worker is constructed. */
  workers?: readonly WorkerLike[];
}

/**
 * Build the pool a batch run should use.
 *
 * Never throws and never returns null: if not one real Worker can be made, the
 * result is a one-worker in-process pool, which is slower and correct.
 */
export function createBatchPoolForHost(opts: CreatePoolOptions = {}): BatchPool {
  if (opts.workers && opts.workers.length > 0) return createBatchPool(opts.workers);

  const cores = typeof navigator === 'undefined' ? undefined : navigator.hardwareConcurrency;
  const size = resolvePoolSize(cores, opts.poolSize ?? 0);

  const workers: WorkerLike[] = [];
  for (let i = 0; i < size; i += 1) {
    const worker = tryCreateWorker();
    if (worker === null) break;
    workers.push(worker);
  }

  if (workers.length === 0) return createBatchPool([createInProcessWorker()]);
  return createBatchPool(workers);
}
