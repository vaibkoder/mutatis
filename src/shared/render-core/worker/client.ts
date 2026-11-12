/**
 * Main-thread render client (feature 16 §3; ADR-0024) — host-free.
 *
 * Drives a RenderCore worker over the message protocol: posts render jobs, routes
 * progress to a per-job callback, resolves with the `LayoutResult`, and supports
 * cancellation by job id. Typed against a minimal `WorkerLike` so the real
 * `Worker` and a fake (for tests) both satisfy it. The client touches no DOM and
 * no Office API — it only marshals messages.
 */
import type { RenderInput } from '../input/index.js';
import type { FaceInput } from '../fonts/registry.js';
import type { LayoutResult, RenderTheme } from '../types.js';
import type { ClientToWorker, WorkerToClient } from './protocol.js';

/** The subset of `Worker` the client needs (real `Worker` satisfies it). */
export interface WorkerLike {
  postMessage(message: ClientToWorker, transfer?: Transferable[]): void;
  onmessage: ((ev: { data: WorkerToClient }) => void) | null;
  terminate?(): void;
}

export interface RenderJobOptions {
  theme?: RenderTheme;
  onProgress?: (pct: number) => void;
}

export interface RenderJob {
  id: number;
  /** Resolves with the layout, rejects on error or with an AbortError on cancel. */
  layout: Promise<LayoutResult>;
}

export interface RenderClient {
  render(input: RenderInput, faces: FaceInput[], opts?: RenderJobOptions): RenderJob;
  cancel(id: number): void;
  terminate(): void;
}

interface Pending {
  resolve(layout: LayoutResult): void;
  reject(err: Error): void;
  onProgress: ((pct: number) => void) | undefined;
}

export function createRenderClient(worker: WorkerLike): RenderClient {
  let nextId = 1;
  const pending = new Map<number, Pending>();

  worker.onmessage = (ev): void => {
    const m = ev.data;
    const job = pending.get(m.id);
    if (!job) return;
    switch (m.type) {
      case 'progress':
        job.onProgress?.(m.pct);
        break;
      case 'result':
        pending.delete(m.id);
        job.resolve(m.layout);
        break;
      case 'cancelled':
        pending.delete(m.id);
        job.reject(new DOMException('Render cancelled', 'AbortError'));
        break;
      case 'error':
        pending.delete(m.id);
        job.reject(new Error(m.message));
        break;
    }
  };

  return {
    render(input, faces, opts = {}): RenderJob {
      const id = nextId++;
      const layout = new Promise<LayoutResult>((resolve, reject) => {
        pending.set(id, { resolve, reject, onProgress: opts.onProgress });
        const request: ClientToWorker = opts.theme
          ? { type: 'render', id, input, faces, theme: opts.theme }
          : { type: 'render', id, input, faces };
        worker.postMessage(request);
      });
      return { id, layout };
    },
    cancel(id): void {
      worker.postMessage({ type: 'cancel', id });
    },
    terminate(): void {
      worker.terminate?.();
    },
  };
}
