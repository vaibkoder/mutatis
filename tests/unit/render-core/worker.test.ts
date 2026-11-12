import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import {
  createRenderClient, createWorkerHandler,
  type WorkerLike, type ClientToWorker, type WorkerToClient,
} from '@shared/render-core/index.js';
import { testFaces } from './_fixtures.js';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const BODY = `<w:body ${NS}><w:p><w:r><w:t>Hello worker world from RenderCore.</w:t></w:r></w:p></w:body>`;

/** An in-process fake Worker: client.postMessage → handler.handle; handler.post → client.onmessage. */
function makeFakeWorker(): WorkerLike {
  const fake: WorkerLike = {
    onmessage: null,
    postMessage(message: ClientToWorker): void { void handler.handle(message); },
    terminate(): void { /* no-op */ },
  };
  const handler = createWorkerHandler({ post: (m: WorkerToClient) => fake.onmessage?.({ data: m }) });
  return fake;
}

describe('worker transport (client ↔ handler round-trip)', () => {
  it('round-trips a layout and reports progress ending at 100', async () => {
    const client = createRenderClient(makeFakeWorker());
    const progress: number[] = [];
    const job = client.render({ authoredBodyXml: BODY }, testFaces(), { onProgress: (p) => progress.push(p) });
    const layout = await job.layout;

    expect(layout.pages.length).toBeGreaterThanOrEqual(1);
    expect(layout.pages[0]!.bodyLines.length).toBeGreaterThan(0);
    expect(progress[0]).toBe(0);
    expect(progress.at(-1)).toBe(100);
    expect(progress).toContain(40);
    expect(progress).toContain(90);
  });

  it('rejects with an AbortError when a job is cancelled before it completes', async () => {
    const client = createRenderClient(makeFakeWorker());
    const job = client.render({ authoredBodyXml: BODY }, testFaces());
    // Cancel synchronously while the handler is awaiting wasm shaper init.
    client.cancel(job.id);
    await expect(job.layout).rejects.toMatchObject({ name: 'AbortError' });
  });

  it('reports an error message for malformed input rather than throwing on the main thread', async () => {
    const client = createRenderClient(makeFakeWorker());
    // No faces ⇒ FontResolver throws inside the worker; surfaced as an ErrorMessage.
    const job = client.render({ authoredBodyXml: BODY }, []);
    await expect(job.layout).rejects.toThrow(/no faces/i);
  });

  it('the worker entry imports only host-free code (no Office/Word/fetch)', () => {
    const raw = readFileSync(resolve(__dirname, '../../../src/workers/render-core.worker.ts'), 'utf8');
    // Strip comments — only the executable code must be host-free (the doc comment
    // legitimately names the banned globals to explain why they are absent).
    const code = raw.replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/[^\n]*/g, '');
    expect(/\bOffice\b/.test(code)).toBe(false);
    expect(/\bWord\./.test(code)).toBe(false);
    expect(/\bfetch\s*\(/.test(code)).toBe(false);
    expect(/@features\//.test(code)).toBe(false);
    // Every import must come from the host-free engine.
    const imports = [...code.matchAll(/from\s+'([^']+)'/g)].map((m) => m[1]!);
    expect(imports.length).toBeGreaterThan(0);
    for (const spec of imports) expect(spec.startsWith('@shared/render-core/')).toBe(true);
  });
});
