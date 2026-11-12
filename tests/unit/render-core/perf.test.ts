/**
 * RenderCore perf — large-document layout off the main thread (feature 16 §5
 * acceptance 8).
 *
 * Drives a ~150-page redline through the REAL pipeline (parse → build → paginate
 * → decorate, real HarfBuzz wasm via fs) using the same in-process fake Worker
 * the transport/adapter tests use — client.postMessage → handler.handle;
 * handler.post → client.onmessage. We measure layout wall-time and assert the
 * work is delivered in coarse stages (the protocol that keeps the main thread
 * free to paint): in production `handler` runs inside a real Web Worker and the
 * main thread only awaits the `LayoutResult` promise and paints. We cannot prove
 * non-blocking from Node (no real thread), so the live both-host check is the
 * D3 user gate; here we evidence pages≈150, the staged progress protocol, and
 * the wall-time number.
 */
import { describe, it, expect } from 'vitest';
import {
  createRenderClient, createWorkerHandler,
  type WorkerLike, type ClientToWorker, type WorkerToClient,
} from '@shared/render-core/index.js';
import { testFaces } from './_fixtures.js';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** In-process fake Worker → real handler (real pipeline, real wasm via fs). */
function makeFakeClient() {
  const fake: WorkerLike = {
    onmessage: null,
    postMessage(message: ClientToWorker): void { void handler.handle(message); },
    terminate(): void { /* no-op */ },
  };
  const handler = createWorkerHandler({ post: (m: WorkerToClient) => fake.onmessage?.({ data: m }) });
  return createRenderClient(fake);
}

/**
 * A long tracked-change body. Default A4 geometry fits ~56 single-line paragraphs
 * per page, so ~8,600 paragraphs clears 150 pages with margin. Each paragraph
 * carries an Arial run plus an `<w:ins>` revision so the layout exercises the
 * decorate (change) path too, not just plain text.
 */
function longRedlineBody(paras: number): string {
  let s = '';
  for (let i = 0; i < paras; i++) {
    s +=
      `<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr>` +
      `<w:t xml:space="preserve">Clause ${i}: the parties shall comply with </w:t></w:r>` +
      `<w:ins w:id="${i}" w:author="A"><w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr>` +
      `<w:t xml:space="preserve">all obligations herein.</w:t></w:r></w:ins></w:p>`;
  }
  return `<w:body ${NS}>${s}</w:body>`;
}

describe('RenderCore perf: ~150-page layout (acceptance 8)', () => {
  it('lays out ~150 pages off the main thread with staged progress', async () => {
    const PARAS = 8600;
    const client = makeFakeClient();
    const progress: number[] = [];

    const t0 = performance.now();
    const job = client.render(
      { authoredBodyXml: longRedlineBody(PARAS) },
      testFaces(),
      { onProgress: (p) => progress.push(p) },
    );
    const layout = await job.layout;
    const elapsedMs = performance.now() - t0;

    // ~150 pages of real pagination.
    expect(layout.pages.length).toBeGreaterThanOrEqual(150);

    // Staged progress = the protocol that lets the worker yield while the main
    // thread paints: build (40) → paginate (90) → decorate/done (100), bracketed
    // by 0. Each is a discrete message, monotonically non-decreasing.
    expect(progress[0]).toBe(0);
    expect(progress).toContain(40);
    expect(progress).toContain(90);
    expect(progress.at(-1)).toBe(100);
    for (let i = 1; i < progress.length; i++) expect(progress[i]!).toBeGreaterThanOrEqual(progress[i - 1]!);

    // No content dropped across the page break (never-suppress at scale): the
    // total painted lines is at least one per source paragraph.
    const totalLines = layout.pages.reduce((a, p) => a + p.bodyLines.length, 0);
    expect(totalLines).toBeGreaterThanOrEqual(PARAS);

    // Record the wall-time for the Evidence section. Loose ceiling as a
    // regression guard only — Node shaping is slower than a browser Worker and
    // this is the engine doing the full real pipeline single-threaded.

    console.log(
      `[perf] ${layout.pages.length} pages, ${totalLines} lines, layout wall-time ${elapsedMs.toFixed(0)} ms`,
    );
    expect(elapsedMs).toBeLessThan(60_000);
  }, 90_000);
});
