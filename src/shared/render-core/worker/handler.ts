/**
 * Worker-side handler (feature 16 §3; ADR-0024) — host-free, DOM-free.
 *
 * The pure pipeline split into its stages so the worker can post coarse progress
 * (build → paginate → decorate) and honour cancellation at stage boundaries — the
 * only honest interruption points for a synchronous CPU task. Decoupled from the
 * real `WorkerGlobalScope` behind a `HandlerPort` so it is unit-tested against a
 * fake port with no actual Worker. Imports only host-free engine code.
 */
import { buildRenderModel } from '../input/index.js';
import { paginate } from '../paginate/paginator.js';
import { decorateLayout } from '../decorate/decorate.js';
import { createFontResolver } from '../fonts/registry.js';
import { createHarfBuzzShaper, type Shaper } from '../layout/shape.js';
import { DEFAULT_THEME } from '../types.js';
import type { ClientToWorker, WorkerToClient } from './protocol.js';

export interface HandlerPort {
  post(message: WorkerToClient): void;
}

export interface WorkerHandler {
  handle(message: ClientToWorker): Promise<void>;
}

/**
 * Build a handler over `port`. One HarfBuzz shaper is created lazily and reused
 * across render jobs (wasm init is the slow first step — and a natural async
 * cancellation point).
 */
export function createWorkerHandler(port: HandlerPort): WorkerHandler {
  const cancelled = new Set<number>();
  let shaperPromise: Promise<Shaper> | null = null;
  const getShaper = (): Promise<Shaper> => (shaperPromise ??= createHarfBuzzShaper());

  const isCancelled = (id: number): boolean => {
    if (!cancelled.has(id)) return false;
    cancelled.delete(id);
    port.post({ type: 'cancelled', id });
    return true;
  };

  return {
    async handle(message: ClientToWorker): Promise<void> {
      if (message.type === 'cancel') { cancelled.add(message.id); return; }
      const { id } = message;
      try {
        port.post({ type: 'progress', id, pct: 0 });
        const shaper = await getShaper(); // async ⇒ a cancel posted now is seen below
        if (isCancelled(id)) return;

        const fonts = createFontResolver(message.faces);
        const model = buildRenderModel(message.input);
        port.post({ type: 'progress', id, pct: 40 });

        const layout = paginate(model, shaper, fonts);
        port.post({ type: 'progress', id, pct: 90 });

        decorateLayout(layout, message.theme ?? DEFAULT_THEME);
        if (isCancelled(id)) return;

        port.post({ type: 'progress', id, pct: 100 });
        port.post({ type: 'result', id, layout });
      } catch (e) {
        port.post({ type: 'error', id, message: e instanceof Error ? e.message : String(e) });
      }
    },
  };
}
