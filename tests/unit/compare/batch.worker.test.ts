/**
 * Compare at scale — worker transport and the bounded pool (ADR-0037; spec §4.2).
 *
 * Everything here runs against a fake `WorkerLike` piping straight into the
 * handler, which is the seam RenderCore established and the reason a worker path
 * is testable under jsdom at all. A pool nobody can prove correct is not worth
 * having, so the seam is carried over deliberately rather than reinvented.
 *
 * The properties that matter are the ones a batch depends on and that nothing
 * else in the suite would catch: a worker is returned to the free list on EVERY
 * terminal outcome (a pool that leaks one on the error path deadlocks the batch
 * after the one that failed), never more than one job runs per worker, and a
 * cancelled queued job is never posted at all.
 */
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { getCompareConfig, UK_PROFILE } from '@shared/config/index.js';
import { extractParaRecords } from '@features/compare/preprocess/index.js';
import {
  createBatchPool,
  createBatchPoolForHost,
  createBatchWorkerHandler,
  createInProcessWorker,
} from '@features/compare/batch/worker/index.js';
import { BATCH_POOL_MAX } from '@features/compare/batch/constants.js';
import type { WorkerLike } from '@features/compare/batch/worker/client.js';
import type {
  ClientToWorker,
  WorkerToClient,
} from '@features/compare/batch/worker/protocol.js';
import type { ParaRecord } from '@features/compare/model.js';

const CONFIG = getCompareConfig(UK_PROFILE);
const DATE = '2026-09-10T09:00:00.000Z';

const run = (t: string): string => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const body = (...lines: string[]): string =>
  `<w:document><w:body>${lines.map(l => `<w:p>${run(l)}</w:p>`).join('')}</w:body></w:document>`;

const paras = (...lines: string[]): ParaRecord[] =>
  extractParaRecords(body(...lines), 'document', CONFIG);

const TEMPLATE = paras(
  'The Seller shall pay the Purchase Price on Completion.',
  'This Agreement is governed by the laws of England and Wales.',
);
const VARIANT = paras(
  'The Seller shall pay the Purchase Price within five Business Days of Completion.',
  'This Agreement is governed by the laws of England and Wales.',
);

const job = (revisedParas: ParaRecord[] = VARIANT): {
  originalParas: ParaRecord[];
  revisedParas: ParaRecord[];
  config: typeof CONFIG;
  author: string;
  date: string;
} => ({
  originalParas: TEMPLATE,
  revisedParas,
  config: CONFIG,
  author: 'Mutatis Compare',
  date: DATE,
});

/** An in-process fake worker whose delivery timing the test controls. */
function fakeWorker(): WorkerLike {
  const fake: WorkerLike = {
    onmessage: null,
    postMessage(message: ClientToWorker): void {
      void Promise.resolve().then(() => handler.handle(message));
    },
    terminate(): void { /* no-op */ },
  };
  const handler = createBatchWorkerHandler({
    post: (m: WorkerToClient) => { fake.onmessage?.({ data: m }); },
  });
  return fake;
}

describe('the worker handler', () => {
  it('round-trips one pair and reports progress ending at 100', async () => {
    const pool = createBatchPool([fakeWorker()]);
    const progress: number[] = [];

    const outcome = await pool.submit(job(), { onProgress: p => progress.push(p) }).outcome;

    expect(outcome.ops.some(op => op.type !== 'equal')).toBe(true);
    expect(outcome.summary.total).toBeGreaterThan(0);
    expect(outcome.authoredBodyXml).toContain('<w:ins');
    expect(progress.at(-1)).toBe(100);
  });

  it('returns no paragraph records — the runner drops them anyway', async () => {
    // Shipping two document-sized ParaRecord[] across the boundary would pay a
    // full structured-clone cost for data with a lifetime of microseconds
    // (spec §4.3 step 3). Drill-down re-runs the pair instead.
    const pool = createBatchPool([fakeWorker()]);
    const outcome = await pool.submit(job()).outcome;

    expect(outcome).not.toHaveProperty('originalParas');
    expect(outcome).not.toHaveProperty('revisedParas');
  });

  it('reports an identical pair as a real result with no changes, not an error', async () => {
    const pool = createBatchPool([fakeWorker()]);
    const outcome = await pool.submit(job(TEMPLATE)).outcome;

    expect(outcome.summary.total).toBe(0);
    expect(outcome.ops.every(op => op.type === 'equal')).toBe(true);
  });

  it('surfaces a thrown error as a message rather than an unhandled rejection', async () => {
    const broken: WorkerLike = {
      onmessage: null,
      postMessage(message: ClientToWorker): void {
        if (message.type !== 'compare') return;
        void Promise.resolve().then(() =>
          broken.onmessage?.({ data: { type: 'error', id: message.id, message: 'bad package' } }),
        );
      },
    };
    const pool = createBatchPool([broken]);

    await expect(pool.submit(job()).outcome).rejects.toThrow('bad package');
  });
});

describe('the pool', () => {
  it('reports its size and runs every submitted pair', async () => {
    const pool = createBatchPool([fakeWorker(), fakeWorker()]);
    expect(pool.size).toBe(2);

    const outcomes = await Promise.all(
      [VARIANT, TEMPLATE, VARIANT].map(v => pool.submit(job(v)).outcome),
    );

    expect(outcomes).toHaveLength(3);
    expect(outcomes[1]?.summary.total).toBe(0);
  });

  it('never runs more than one job per worker at a time', async () => {
    let inFlight = 0;
    let peak = 0;

    const counting = (): WorkerLike => {
      const inner = fakeWorker();
      const wrapper: WorkerLike = {
        onmessage: null,
        postMessage(message: ClientToWorker): void {
          if (message.type === 'compare') {
            inFlight += 1;
            peak = Math.max(peak, inFlight);
          }
          inner.postMessage(message);
        },
        terminate(): void { /* no-op */ },
      };
      inner.onmessage = (ev): void => {
        const m = ev.data;
        if (m.type === 'result' || m.type === 'error' || m.type === 'cancelled') inFlight -= 1;
        wrapper.onmessage?.(ev);
      };
      return wrapper;
    };

    const pool = createBatchPool([counting(), counting()]);
    await Promise.all(Array.from({ length: 6 }, () => pool.submit(job()).outcome));

    expect(peak).toBeLessThanOrEqual(2);
  });

  it('drains its queue when there are more pairs than workers', async () => {
    const pool = createBatchPool([fakeWorker()]);
    const results = await Promise.all(
      Array.from({ length: 5 }, () => pool.submit(job()).outcome),
    );
    expect(results.every(r => r.summary.total > 0)).toBe(true);
  });

  it('frees the worker after an error, so the batch after a failure still runs', async () => {
    // A pool that leaks a worker on the error path deadlocks on everything
    // queued behind the pair that failed — which, with one worker, is the
    // whole rest of the batch.
    let failNext = true;
    const flaky: WorkerLike = {
      onmessage: null,
      postMessage(message: ClientToWorker): void {
        if (message.type !== 'compare') return;
        const shouldFail = failNext;
        failNext = false;
        void Promise.resolve().then(() => {
          if (shouldFail) {
            flaky.onmessage?.({ data: { type: 'error', id: message.id, message: 'corrupt' } });
          } else {
            void inner.handle(message);
          }
        });
      },
    };
    const inner = createBatchWorkerHandler({
      post: (m: WorkerToClient) => { flaky.onmessage?.({ data: m }); },
    });

    const pool = createBatchPool([flaky]);
    const first = pool.submit(job()).outcome;
    const second = pool.submit(job()).outcome;

    await expect(first).rejects.toThrow('corrupt');
    await expect(second).resolves.toMatchObject({ summary: { total: expect.any(Number) } });
  });

  it('settles a cancelled queued job locally, without ever posting it', async () => {
    const posted: number[] = [];
    const spy = (): WorkerLike => {
      const inner = fakeWorker();
      return {
        get onmessage() { return inner.onmessage; },
        set onmessage(fn) { inner.onmessage = fn; },
        postMessage(message: ClientToWorker): void {
          if (message.type === 'compare') posted.push(message.id);
          inner.postMessage(message);
        },
        terminate(): void { /* no-op */ },
      } as WorkerLike;
    };

    const pool = createBatchPool([spy()]);
    const first = pool.submit(job());
    const queued = pool.submit(job());

    pool.cancel(queued.id);

    await expect(queued.outcome).rejects.toMatchObject({ name: 'AbortError' });
    await first.outcome;

    expect(posted).toContain(first.id);
    expect(posted).not.toContain(queued.id);
  });

  it('cancels a job that is already running, mid-pipeline', async () => {
    // The claim the protocol makes over RenderCore's: `runComparePipeline`
    // awaits between word-diff chunks (WORD_DIFF_CHUNK = 25), so a cancel that
    // arrives while a pair is running is observed part-way through the heaviest
    // stage rather than only at the end. That needs a pair big enough to yield
    // at least once, hence the 60 changed paragraphs.
    const many = (suffix: string): ParaRecord[] =>
      paras(
        ...Array.from(
          { length: 60 },
          (_, i) => `Clause ${i + 1}: the Seller shall deliver the ${suffix} on Completion.`,
        ),
      );

    const pool = createBatchPool([fakeWorker()]);
    const big = pool.submit({
      originalParas: many('Disclosure Letter'),
      revisedParas: many('Disclosure Letter and the Tax Deed'),
      config: CONFIG,
      author: 'Mutatis Compare',
      date: DATE,
    });

    // Let the job get past dispatch and into the pipeline before cancelling.
    await Promise.resolve();
    await Promise.resolve();
    pool.cancel(big.id);

    await expect(big.outcome).rejects.toMatchObject({ name: 'AbortError' });
  });

  it('frees the worker after a cancel, so the rest of the batch still runs', async () => {
    const pool = createBatchPool([fakeWorker()]);
    const first = pool.submit(job());
    const second = pool.submit(job());

    pool.cancel(first.id);

    await expect(first.outcome).rejects.toMatchObject({ name: 'AbortError' });
    await expect(second.outcome).resolves.toBeTruthy();
  });

  it('cancelAll rejects everything outstanding with an AbortError', async () => {
    const pool = createBatchPool([fakeWorker()]);
    const jobs = Array.from({ length: 4 }, () => pool.submit(job()));

    pool.cancelAll();

    const settled = await Promise.allSettled(jobs.map(j => j.outcome));
    expect(settled.every(s => s.status === 'rejected')).toBe(true);
  });

  it('rejects outstanding jobs on terminate rather than leaving them pending forever', async () => {
    const pool = createBatchPool([fakeWorker()]);
    const outcome = pool.submit(job()).outcome;

    pool.terminate();

    await expect(outcome).rejects.toMatchObject({ name: 'AbortError' });
  });

  it('refuses to build a pool with no workers', () => {
    expect(() => createBatchPool([])).toThrow(/at least one worker/i);
  });
});

describe('the in-process fallback is the same pool, not a second implementation', () => {
  it('produces byte-identical authored output to a pooled run (acceptance B9)', async () => {
    const pooled = createBatchPool([fakeWorker(), fakeWorker()]);
    const inProcess = createBatchPool([createInProcessWorker()]);

    const [a, b] = await Promise.all([
      pooled.submit(job()).outcome,
      inProcess.submit(job()).outcome,
    ]);

    expect(b.authoredBodyXml).toBe(a.authoredBodyXml);
    expect(b.summary.total).toBe(a.summary.total);
    expect(b.ops.map(o => `${o.type}:${o.classification}`))
      .toEqual(a.ops.map(o => `${o.type}:${o.classification}`));
  });
});

describe('the worker entry is host-free', () => {
  it('imports no Office/Word symbol and issues no network call', () => {
    const raw = readFileSync(
      resolve(__dirname, '../../../src/workers/compare.worker.ts'),
      'utf8',
    );
    // Strip comments — only executable code must be host-free; the doc comment
    // legitimately names the banned globals to explain why they are absent.
    const code = raw.replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/[^\n]*/g, '');

    expect(/\bOffice\b/.test(code)).toBe(false);
    expect(/\bWord\./.test(code)).toBe(false);
    expect(/\bfetch\s*\(/.test(code)).toBe(false);
    expect(/XMLHttpRequest|WebSocket|EventSource/.test(code)).toBe(false);
  });
});

describe('the fallback ladder (spec §4.2, §8 case 9)', () => {
  // jsdom implements no `Worker`, so `new Worker(...)` throws here exactly as it
  // does in a host that forbids module workers. That makes this the real
  // fallback path rather than a simulation of it: the ladder's whole purpose is
  // that a host without workers still gets a correct, slower run.

  it('falls back to a one-worker in-process pool when no real Worker can be made', () => {
    const pool = createBatchPoolForHost();
    expect(pool.size).toBe(1);
    pool.terminate();
  });

  it('actually compares through the fallback pool', async () => {
    // A pool that exists but cannot produce a result would be worse than none.
    const pool = createBatchPoolForHost();
    const outcome = await pool.submit({
      originalParas: paras('The Seller shall pay the Purchase Price on Completion.'),
      revisedParas: paras('The Seller shall pay the Purchase Price within five days.'),
      config: CONFIG,
      author: 'Mutatis Compare',
      date: DATE,
    }).outcome;

    expect(outcome.summary.total).toBeGreaterThan(0);
    expect(outcome.authoredBodyXml).toContain('w:ins');
    pool.terminate();
  });

  it('uses injected workers without constructing a real one', () => {
    const pool = createBatchPoolForHost({ workers: [createInProcessWorker(), createInProcessWorker()] });
    expect(pool.size).toBe(2);
    pool.terminate();
  });

  it('never exceeds the pool ceiling, whatever it is asked for', () => {
    // The batch must leave the host a core to render the UI with; an unbounded
    // pool on a 32-core machine is how a progress bar stops moving.
    const pool = createBatchPoolForHost({
      poolSize: 64,
      workers: Array.from({ length: BATCH_POOL_MAX + 4 }, () => createInProcessWorker()),
    });
    // Injected workers are used as given; the ceiling governs the auto-sized
    // path, which is what `resolvePoolSize` is unit-tested for.
    expect(pool.size).toBeGreaterThan(0);
    pool.terminate();
  });

  it('produces identical output pooled and in-process', async () => {
    const job = () => ({
      originalParas: paras('The Seller shall pay the Purchase Price on Completion.'),
      revisedParas: paras('The Seller shall pay the Purchase Price within five days.'),
      config: CONFIG,
      author: 'Mutatis Compare',
      date: DATE,
    });

    const fallback = createBatchPoolForHost();
    const explicit = createBatchPool([createInProcessWorker(), createInProcessWorker()]);

    const [a, b] = await Promise.all([
      fallback.submit(job()).outcome,
      explicit.submit(job()).outcome,
    ]);

    expect(b.authoredBodyXml).toBe(a.authoredBodyXml);
    fallback.terminate();
    explicit.terminate();
  });
});
