/**
 * RenderCore Web Worker entry (feature 16 §3 "Worker boundary"; ADR-0024).
 *
 * Thin transport only: wires the host-free `createWorkerHandler` to the worker
 * global so the entire parse→layout→paginate→decorate pipeline runs off the main
 * thread. Imports ONLY host-free engine code — no `Office.*`/`Word.*`, no DOM, no
 * `fetch` (font bytes arrive in the render message; ADR-0001/0003 zero egress).
 */
import { createWorkerHandler } from '@shared/render-core/worker/handler.js';
import type { ClientToWorker } from '@shared/render-core/worker/protocol.js';

/** Minimal view of the dedicated-worker global (the project's lib excludes WebWorker). */
interface WorkerScope {
  postMessage(message: unknown): void;
  onmessage: ((ev: { data: ClientToWorker }) => void) | null;
}

const scope = self as unknown as WorkerScope;
const handler = createWorkerHandler({ post: (message) => scope.postMessage(message) });
scope.onmessage = (ev): void => { void handler.handle(ev.data); };
