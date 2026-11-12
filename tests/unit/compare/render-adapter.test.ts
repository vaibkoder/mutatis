/**
 * Compare → RenderCore adapter (feature 16 / ADR-0024, Phase 2 B2/B3).
 *
 * Exercises the host-free adapter end-to-end through the REAL engine pipeline,
 * driven by an in-process fake Worker (the same fake `tests/unit/render-core/
 * worker.test.ts` uses — client.postMessage → handler.handle; handler.post →
 * client.onmessage). The fake runs the real parse→layout→paginate→decorate with
 * real HarfBuzz shaping (wasm via fs) in Node, so it produces a real
 * `LayoutResult` the adapter then paints to a PDF / preview.
 */
import { describe, it, expect } from 'vitest';
import { PDFDocument } from '@pdfme/pdf-lib';
import {
  createRenderClient, createWorkerHandler,
  type WorkerLike, type ClientToWorker, type WorkerToClient, type RcRevisionType,
} from '@shared/render-core/index.js';
import { PizZipBundle } from '@shared/ooxml/index.js';
import { parseNumberingScheme } from '@shared/numbering/index.js';
import { extractParts } from '@features/compare/io/parts.js';
import { synthesiseDocx } from '@features/compare/io/index.js';
import { runPipeline } from '@features/compare/pipeline.js';
import { getActiveProfile, getCompareConfig } from '@shared/config/index.js';
import {
  buildRenderInput, renderRedlinePdf, mountRedlinePreview, extractEmbeddedFaces, toRcClassMap,
} from '@features/compare/render/index.js';
import type { RenderSource } from '@features/compare/ui-types.js';
import { smokeDocxBytes, smokeFixturesAvailable } from '../../fixtures/smoke-docx.js';

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

const docxBytes = smokeDocxBytes;

/** Extract revision-text tokens by type from authored `<w:body>` markup. */
function authoredRevisionTokens(xml: string): Record<RcRevisionType, Set<string>> {
  const out: Record<RcRevisionType, Set<string>> = {
    ins: new Set(), del: new Set(), moveFrom: new Set(), moveTo: new Set(),
  };
  const decode = (s: string): string =>
    s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  const tokenise = (inner: string, into: Set<string>): void => {
    for (const tok of decode(inner.replace(/<[^>]*>/g, '')).split(/\s+/)) if (tok.length) into.add(tok);
  };
  const tags: Array<[RcRevisionType, string]> = [
    ['ins', 'w:ins'], ['del', 'w:del'], ['moveFrom', 'w:moveFrom'], ['moveTo', 'w:moveTo'],
  ];
  for (const [type, tag] of tags) {
    const re = new RegExp(`<${tag}(?=[\\s>])[^>]*>([\\s\\S]*?)</${tag}>`, 'g');
    let m: RegExpExecArray | null;
    while ((m = re.exec(xml)) !== null) tokenise(m[1]!, out[type]);
  }
  return out;
}

/** Build a real `RenderSource` from the smoke-test compare. */
async function realRenderSource(): Promise<RenderSource> {
  const originalBundle = new PizZipBundle(docxBytes('original.docx'));
  const revisedBundle = new PizZipBundle(docxBytes('revised.docx'));
  const config = getCompareConfig(getActiveProfile());
  const result = await runPipeline({
    originalParts: extractParts(originalBundle),
    revisedParts: extractParts(revisedBundle),
    config,
    author: config.author,
    date: '2026-06-15T00:00:00Z',
  });
  const parts = extractParts(originalBundle);
  return {
    authoredBodyXml: result.authoredBodyXml,
    ...(parts.stylesXml ? { stylesXml: parts.stylesXml } : {}),
    ...(parts.footnotesXml ? { footnotesXml: parts.footnotesXml } : {}),
    numbering: parseNumberingScheme(originalBundle),
    classByRevId: result.classByRevId,
    bundle: originalBundle,
  };
}

/** Build a Text-mode `RenderSource` (synthesised Calibri default-styles package). */
async function textRenderSource(): Promise<RenderSource> {
  const originalBundle = synthesiseDocx('The Seller shall deliver the Goods.\nClause two is unchanged.');
  const revisedBundle = synthesiseDocx('The Seller may deliver the Goods.\nClause two is unchanged.');
  const config = getCompareConfig(getActiveProfile());
  const result = await runPipeline({
    originalParts: extractParts(originalBundle),
    revisedParts: extractParts(revisedBundle),
    config,
    author: config.author,
    date: '2026-06-15T00:00:00Z',
  });
  const parts = extractParts(originalBundle);
  return {
    authoredBodyXml: result.authoredBodyXml,
    ...(parts.stylesXml ? { stylesXml: parts.stylesXml } : {}),
    classByRevId: result.classByRevId,
    bundle: originalBundle,
  };
}

describe('Compare→RenderCore adapter: PDF', () => {
  it('renders a valid multi-page vector PDF from a compare source', async () => {
    // Synthesise a long redline so the output is genuinely multi-page.
    const paras: string[] = [];
    for (let i = 0; i < 80; i++) {
      paras.push(
        `<w:p><w:r><w:t xml:space="preserve">Clause ${i}: the parties shall comply with </w:t></w:r>` +
        `<w:ins w:id="${i}" w:author="A"><w:r><w:t xml:space="preserve">all obligations set out herein. </w:t></w:r></w:ins></w:p>`,
      );
    }
    const src: RenderSource = {
      authoredBodyXml: `<w:body ${NS}>${paras.join('')}</w:body>`,
      bundle: new PizZipBundle(),
    };
    const bytes = await renderRedlinePdf(src, { client: makeFakeClient() });
    const pdf = await PDFDocument.load(bytes);
    expect(pdf.getPageCount()).toBeGreaterThan(1);
  });
});

describe.skipIf(!smokeFixturesAvailable())('Compare→RenderCore adapter: preview (never-suppress)', () => {
  it('mounts paginated frames and paints every authored revision token (ADR-0008)', async () => {
    const src = await realRenderSource();
    const authored = authoredRevisionTokens(src.authoredBodyXml);
    const authoredTotal = (Object.values(authored) as Set<string>[]).reduce((a, s) => a + s.size, 0);
    expect(authoredTotal).toBeGreaterThan(0); // fixtures actually contain changes

    const host = document.createElement('div');
    const handle = await mountRedlinePreview(host, src, { client: makeFakeClient() });
    try {
      const frames = handle.root.querySelectorAll('.rc-page');
      expect(frames.length).toBeGreaterThanOrEqual(1);

      // Every painted run's text, tokenised — the adapter must drop nothing.
      const painted = new Set<string>();
      handle.root.querySelectorAll('.rc-run').forEach((span) => {
        for (const tok of (span.textContent ?? '').split(/\s+/)) if (tok.length) painted.add(tok);
      });
      for (const type of ['ins', 'del', 'moveFrom', 'moveTo'] as RcRevisionType[]) {
        for (const tok of authored[type]) {
          expect(painted.has(tok), `${type} token "${tok}" missing from the mounted preview`).toBe(true);
        }
      }
    } finally {
      handle.destroy();
    }
  });
});

describe('Compare→RenderCore adapter: Text mode (synthesised default styles, acceptance 11)', () => {
  it('renders a Text compare to a valid PDF through the same engine path, Calibri styles applied', async () => {
    const src = await textRenderSource();
    // The synthesised default-styles package (acceptance 11 source) is carried.
    expect(src.stylesXml).toContain('w:ascii="Calibri"');
    // buildRenderInput threads the synthesised styles into the engine input —
    // Text renders identically to Word, just with default styles (no doc styles).
    expect(buildRenderInput(src).stylesXml).toBe(src.stylesXml);

    const bytes = await renderRedlinePdf(src, { client: makeFakeClient() });
    const pdf = await PDFDocument.load(bytes);
    expect(pdf.getPageCount()).toBeGreaterThanOrEqual(1);
  });
});

describe('Compare→RenderCore adapter: input + faces mapping', () => {
  it('buildRenderInput carries the parts and maps classByRevId → engine annotations', () => {
    const src: RenderSource = {
      authoredBodyXml: `<w:body ${NS}><w:p><w:ins w:id="5"><w:r><w:t>x</w:t></w:r></w:ins></w:p></w:body>`,
      stylesXml: '<w:styles/>',
      classByRevId: new Map([['5', ['definedTerm', 'numericalValue']]]),
      bundle: new PizZipBundle(),
    };
    const input = buildRenderInput(src);
    expect(input.authoredBodyXml).toBe(src.authoredBodyXml);
    expect(input.stylesXml).toBe('<w:styles/>');
    expect(input.classByRevId?.get('5')).toEqual(['definedTerm', 'numericalValue']);
  });

  it('extractEmbeddedFaces returns [] for a doc with no embedded fonts; toRcClassMap is 1:1', () => {
    expect(extractEmbeddedFaces(new PizZipBundle())).toEqual([]);
    const mapped = toRcClassMap(new Map([['9', ['structural', 'formatting']]]));
    expect(mapped.get('9')).toEqual(['structural', 'formatting']);
  });
});
