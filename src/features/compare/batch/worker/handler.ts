/**
 * Batch-compare worker handler (ADR-0037; spec §4.2) — host-free, DOM-free.
 *
 * Follows `shared/render-core/worker/handler.ts`: decoupled from the real
 * `WorkerGlobalScope` behind a `HandlerPort` so it can be unit-tested against a
 * fake port with no actual Worker, which is what makes the worker path testable
 * under jsdom at all.
 *
 * Cancellation is finer-grained here than in RenderCore. `runComparePipeline`
 * awaits between word-diff chunks, so a `cancel` that arrives mid-run is
 * delivered by the message loop at one of those yields and observed on the very
 * next poll — the render handler, whose stages are synchronous, can only check
 * between them. The rejection contract is the pipeline's own: `AbortError`.
 *
 * Imports only host-free compare code — no `Office.*`/`Word.*`, no `fetch`.
 */
import { runComparePipeline } from '../../pipeline.js';
import type { ClientToWorker, WorkerToClient } from './protocol.js';

export interface HandlerPort {
  post(message: WorkerToClient): void;
}

export interface BatchWorkerHandler {
  handle(message: ClientToWorker): Promise<void>;
}

/**
 * Build a handler over `port`.
 *
 * Two pieces of state, and the reason for each:
 *
 *  - `live` maps a running job to the abort token the pipeline is polling, so a
 *    `cancel` flips that token directly. No polling loop, no interval.
 *  - `preCancelled` remembers a cancel that arrived *before* its job did, which
 *    happens whenever the pool cancels a job whose `compare` message is still in
 *    flight. Dropping it would let a cancelled batch keep burning a core.
 */
export function createBatchWorkerHandler(port: HandlerPort): BatchWorkerHandler {
  const live = new Map<number, { aborted: boolean }>();
  const preCancelled = new Set<number>();
  /**
   * The highest job id this worker has ever been asked to run.
   *
   * Job ids are minted in increasing order by one client, so an id at or below
   * this has already started — and therefore already finished, since `live` no
   * longer holds it. Remembering a cancel for such an id would keep it in
   * `preCancelled` forever, because the `compare` message that clears it can
   * never arrive again. That is the ordinary result/cancel crossing race, so
   * the leak would be routine rather than exotic.
   */
  let highestStarted = 0;

  return {
    async handle(message: ClientToWorker): Promise<void> {
      if (message.type === 'cancel') {
        const token = live.get(message.id);
        if (token) token.aborted = true;
        else if (message.id > highestStarted) preCancelled.add(message.id);
        // Otherwise the job has already run to completion and the cancel
        // crossed its result in flight: there is nothing left to stop.
        return;
      }

      const { id } = message;
      if (id > highestStarted) highestStarted = id;

      if (preCancelled.delete(id)) {
        port.post({ type: 'cancelled', id });
        return;
      }

      const abortSignal = { aborted: false };
      live.set(id, abortSignal);

      try {
        const result = await runComparePipeline({
          originalParas: message.originalParas,
          revisedParas: message.revisedParas,
          config: message.config,
          author: message.author,
          date: message.date,
          quarantinedOriginal: message.quarantinedOriginal ?? 0,
          quarantinedRevised: message.quarantinedRevised ?? 0,
          onProgress: pct => { port.post({ type: 'progress', id, pct }); },
          abortSignal,
        });

        // A cancel landing after the pipeline's last checkpoint would otherwise
        // deliver a result for a job the caller has already given up on.
        if (abortSignal.aborted) {
          port.post({ type: 'cancelled', id });
          return;
        }

        port.post({
          type: 'result',
          id,
          outcome: {
            ops: result.ops,
            summary: result.summary,
            authoredBodyXml: result.authoredBodyXml,
            classByRevId: result.classByRevId,
            quarantinedOriginal: result.quarantinedOriginal,
            quarantinedRevised: result.quarantinedRevised,
          },
        });
      } catch (e) {
        // The pipeline signals cancellation by throwing. That is not an error
        // the user should be shown as one.
        if (e instanceof DOMException && e.name === 'AbortError') {
          port.post({ type: 'cancelled', id });
          return;
        }
        port.post({ type: 'error', id, message: e instanceof Error ? e.message : String(e) });
      } finally {
        live.delete(id);
      }
    },
  };
}
