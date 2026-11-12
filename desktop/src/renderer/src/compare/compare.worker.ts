/// <reference lib="webworker" />
import { runComparison } from './run-comparison.js';
import { buildPairing } from '../pairing/build-pairing.js';
import { applyEdits } from '../edits/apply-edits.js';
import { stepFor } from './protocol.js';
import type {
  ClientToWorker,
  CompareRequest,
  PairingRequest,
  SelfRedlineRequest,
  WorkerToClient,
} from './protocol.js';

/**
 * Message wiring only. All the work lives in `run-comparison.ts` so it can be
 * tested without a worker — the same split the engine's own workers use.
 */

const ctx = self as unknown as DedicatedWorkerGlobalScope;

/** Cancellation flags by job id. The pipeline polls `.aborted`. */
const cancellations = new Map<number, { aborted: boolean }>();

function post(message: WorkerToClient): void {
  ctx.postMessage(message);
}

async function handleCompare(request: CompareRequest): Promise<void> {
  const { id } = request;
  const abortSignal = { aborted: false };
  cancellations.set(id, abortSignal);

  try {
    post({ type: 'progress', id, pct: 2, step: stepFor(2) });

    const outcome = await runComparison({
      originalBytes: request.originalBytes,
      revisedBytes: request.revisedBytes,
      jurisdiction: request.jurisdiction,
      author: request.author,
      date: request.date,
      originalName: request.originalName,
      revisedName: request.revisedName,
      onProgress: (pct) => post({ type: 'progress', id, pct, step: stepFor(pct) }),
      abortSignal,
    });

    if (abortSignal.aborted) {
      post({ type: 'cancelled', id });
      return;
    }
    post({ type: 'result', id, outcome });
  } catch (error) {
    // The pipeline signals cancellation by THROWING an AbortError, so an abort
    // usually arrives here rather than at the flag check above.
    if (error instanceof DOMException && error.name === 'AbortError') {
      post({ type: 'cancelled', id });
      return;
    }
    post({
      type: 'error',
      id,
      message: error instanceof Error ? error.message : String(error),
    });
  } finally {
    cancellations.delete(id);
  }
}

/**
 * Sketching reads and preprocesses every document on both sides, which is why
 * it belongs here and not on the UI thread: at the 50-document cap that is 100
 * documents parsed before anything is shown.
 */
function handlePairing(request: PairingRequest): void {
  const { id } = request;
  try {
    post({ type: 'progress', id, pct: 5, step: 'Reading documents' });
    const result = buildPairing(request.docs, request.jurisdiction, request.generatedAt);
    post({ type: 'pairing', id, result });
  } catch (error) {
    post({
      type: 'error',
      id,
      message: error instanceof Error ? error.message : String(error),
    });
  }
}

/**
 * Apply the assistant's edits, then redline the result against the document it
 * edited.
 *
 * The redline runs through the SAME pipeline as every other comparison, with the
 * assistant named as the author. That is not a convenience — it is what makes
 * "every AI edit lands as a tracked change the user reviews" true by
 * construction rather than by discipline. There is no second path that could
 * produce an edit the redline does not show.
 */
async function handleSelfRedline(request: SelfRedlineRequest): Promise<void> {
  const { id } = request;
  const abortSignal = { aborted: false };
  cancellations.set(id, abortSignal);

  try {
    post({ type: 'progress', id, pct: 2, step: 'Applying the proposed edits' });

    // The target buffer is transferred, so it is read once and then reused for
    // the comparison — slicing here because applyEdits does not consume it.
    const target = request.targetBytes;
    const { bytes, applied, refused } = applyEdits(target, request.edits);

    if (applied.length === 0) {
      post({
        type: 'error',
        id,
        message:
          refused.length > 0
            ? `No edit could be applied. ${refused[0]?.refusal ?? ''}`
            : 'The assistant proposed no edits.',
      });
      return;
    }

    const outcome = await runComparison({
      originalBytes: target,
      revisedBytes: bytes,
      jurisdiction: request.jurisdiction,
      author: request.author,
      date: request.date,
      originalName: request.targetName,
      revisedName: `${request.targetName} (assistant's edits)`,
      onProgress: (pct) => post({ type: 'progress', id, pct, step: stepFor(pct) }),
      abortSignal,
    });

    if (abortSignal.aborted) {
      post({ type: 'cancelled', id });
      return;
    }

    const editedBytes = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer;

    post({
      type: 'self-redline-result',
      id,
      result: { outcome, applied, refused, editedBytes },
    });
  } catch (error) {
    if (error instanceof DOMException && error.name === 'AbortError') {
      post({ type: 'cancelled', id });
      return;
    }
    post({
      type: 'error',
      id,
      message: error instanceof Error ? error.message : String(error),
    });
  } finally {
    cancellations.delete(id);
  }
}

ctx.onmessage = (event: MessageEvent<ClientToWorker>) => {
  const message = event.data;
  if (message.type === 'compare') {
    void handleCompare(message);
  } else if (message.type === 'pair') {
    handlePairing(message);
  } else if (message.type === 'self-redline') {
    void handleSelfRedline(message);
  } else if (message.type === 'cancel') {
    const flag = cancellations.get(message.id);
    if (flag) flag.aborted = true;
  }
};
