/**
 * Batch-compare Web Worker entry (ADR-0037; spec §4.2).
 *
 * Thin transport only: wires the host-free `createBatchWorkerHandler` to the
 * worker global so a batch's comparisons run off the main thread. Imports ONLY
 * host-free compare code — no `Office.*`/`Word.*`, no DOM, no `fetch`
 * (ADR-0001/0003: everything a job needs arrives in its message).
 */
import { createBatchWorkerHandler } from '@features/compare/batch/worker/handler.js';
import type { ClientToWorker } from '@features/compare/batch/worker/protocol.js';

/** Minimal view of the dedicated-worker global (the project's lib excludes WebWorker). */
interface WorkerScope {
  postMessage(message: unknown): void;
  onmessage: ((ev: { data: ClientToWorker }) => void) | null;
}

const scope = self as unknown as WorkerScope;
const handler = createBatchWorkerHandler({ post: (message) => scope.postMessage(message) });
scope.onmessage = (ev): void => { void handler.handle(ev.data); };
