/**
 * Feature 16 (RenderCore, Phase 2 — C2): the in-pane paginated RedlinePreview.
 *
 * Drives the real engine pipeline through an in-process fake Worker (the same
 * fake the adapter/worker tests use) so the component mounts a real
 * `LayoutResult` without a real Web Worker. Asserts:
 *   - null src → inert (no page frames mounted);
 *   - a real render source → ≥1 paginated `.rc-page` frame, with EVERY authored
 *     revision token painted (never-suppress, ADR-0008);
 *   - unmount tears the preview down (cleanup runs).
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { render } from 'preact';
import {
  createRenderClient, createWorkerHandler,
  type WorkerLike, type ClientToWorker, type WorkerToClient,
} from '@shared/render-core/index.js';
import { PizZipBundle } from '@shared/ooxml/index.js';
import { RedlinePreview } from '@features/compare/components/RedlinePreview.js';
import type { RenderSource } from '@features/compare/ui-types.js';

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

/** A small redline with both equal text and authored insertions. */
const SRC: RenderSource = {
  authoredBodyXml:
    `<w:body ${NS}>` +
    '<w:p><w:r><w:t xml:space="preserve">The parties </w:t></w:r>' +
    '<w:ins w:id="1" w:author="A"><w:r><w:t xml:space="preserve">shall comply fully.</w:t></w:r></w:ins></w:p>' +
    '</w:body>',
  bundle: new PizZipBundle(),
};

describe('RedlinePreview (feature 16 C2 — paginated in-pane preview)', () => {
  let container: HTMLDivElement;
  beforeEach(() => {
    container = document.createElement('div');
    document.body.appendChild(container);
  });

  it('renders nothing to paginate when there is no render source', () => {
    render(<RedlinePreview src={null} />, container);
    // The section chrome is present, but no page frames are mounted.
    expect(container.querySelector('.dp-compare__preview')).not.toBeNull();
    expect(container.querySelectorAll('.rc-page').length).toBe(0);
  });

  it('mounts paginated frames and paints every authored revision token (ADR-0008)', async () => {
    render(<RedlinePreview src={SRC} client={makeFakeClient()} />, container);

    await vi.waitFor(() => {
      expect(container.querySelectorAll('.rc-page').length).toBeGreaterThanOrEqual(1);
    });

    const painted = new Set<string>();
    container.querySelectorAll('.rc-run').forEach(span => {
      for (const tok of (span.textContent ?? '').split(/\s+/)) if (tok.length) painted.add(tok);
    });
    // The inserted run must be present — nothing is suppressed in the preview.
    for (const tok of ['shall', 'comply', 'fully.']) {
      expect(painted.has(tok), `inserted token "${tok}" missing from the preview`).toBe(true);
    }
  });

  it('tears the preview down on unmount', async () => {
    render(<RedlinePreview src={SRC} client={makeFakeClient()} />, container);
    await vi.waitFor(() => {
      expect(container.querySelectorAll('.rc-page').length).toBeGreaterThanOrEqual(1);
    });
    render(null, container);
    expect(container.querySelector('.dp-compare__preview')).toBeNull();
    expect(container.querySelectorAll('.rc-page').length).toBe(0);
  });
});
