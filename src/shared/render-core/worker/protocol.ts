/**
 * Worker message protocol (feature 16 §3 "Worker boundary"; ADR-0024) — host-free.
 *
 * The entire parse→layout→paginate→decorate pipeline is a PURE function, so the
 * worker is only transport: the main thread posts a render request carrying the
 * `RenderInput` + font BYTES; the worker runs the pure pipeline and posts back
 * the plain-data `LayoutResult` (structured-clone-safe), plus coarse progress and
 * cancellation. No `Office.*`/`Word.*`, no `fetch` — bytes in, layout out.
 *
 * Transferables: font-byte buffers MAY be transferred for Phase-2 perf, but the
 * main thread reuses the same bytes to paint the PDF/preview, so by default they
 * are copied (structured clone), not detached. The protocol is transfer-capable.
 */
import type { RenderInput } from '../input/index.js';
import type { FaceInput } from '../fonts/registry.js';
import type { LayoutResult, RenderTheme } from '../types.js';

export interface RenderRequest {
  type: 'render';
  id: number;
  input: RenderInput;
  /** Font faces (clone/embedded bytes) the resolver is built from, in the worker. */
  faces: FaceInput[];
  theme?: RenderTheme;
}

export interface CancelRequest {
  type: 'cancel';
  id: number;
}

export type ClientToWorker = RenderRequest | CancelRequest;

export interface ProgressMessage { type: 'progress'; id: number; pct: number; }
export interface ResultMessage { type: 'result'; id: number; layout: LayoutResult; }
export interface ErrorMessage { type: 'error'; id: number; message: string; }
export interface CancelledMessage { type: 'cancelled'; id: number; }

export type WorkerToClient = ProgressMessage | ResultMessage | ErrorMessage | CancelledMessage;
