/**
 * Task 13 — close the Phase-1 engine:
 *  (a) preview ↔ PDF pagination parity (same LayoutResult → identical page/line
 *      counts) — acceptance criterion 7.
 *  (b) never-suppress against REAL Compare output: run the actual compare pipeline
 *      on the smoke-test fixtures, feed its `authoredBodyXml` + parts through
 *      `renderToLayout`, and assert every native revision span in that markup is
 *      rendered AND decorated by its class — acceptance criterion 1 (ADR-0008).
 *
 * (b) imports `@features/compare` from a TEST only (the engine never does — the
 * `shared ↛ features` wall stands; the compare feature maps into RenderCore's
 * types in Phase 2).
 */
import { describe, it, expect, beforeAll } from 'vitest';
import { PDFDocument } from '@pdfme/pdf-lib';
import {
  renderToLayout, renderPdf, renderPreview, createFontResolver, createHarfBuzzShaper,
  type Shaper, type LayoutResult, type RcRevisionType,
} from '@shared/render-core/index.js';
import { PizZipBundle } from '@shared/ooxml/index.js';
import { parseNumberingScheme } from '@shared/numbering/index.js';
import { extractParts } from '@features/compare/io/parts.js';
import { runPipeline } from '@features/compare/pipeline.js';
import { getActiveProfile, getCompareConfig } from '@shared/config/index.js';
import { testFaces } from './_fixtures.js';
import { smokeDocxBytes, smokeFixturesAvailable } from '../../fixtures/smoke-docx.js';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const fonts = createFontResolver(testFaces());
let shaper: Shaper;
beforeAll(async () => { shaper = await createHarfBuzzShaper(); });

const docxBytes = smokeDocxBytes;

// ─── (a) preview ↔ PDF pagination parity ───────────────────────────────────────

describe('preview ↔ PDF pagination parity (acceptance criterion 7)', () => {
  it('paints identical page + per-page line counts from the one shared layout result', async () => {
    const paras: string[] = [];
    for (let i = 0; i < 80; i++) {
      paras.push(`<w:p><w:r><w:t xml:space="preserve">Clause ${i}: the parties to this Agreement shall comply with the obligations set out herein and governed by Singapore law. </w:t></w:r></w:p>`);
    }
    const layout = renderToLayout({ authoredBodyXml: `<w:body ${NS}>${paras.join('')}</w:body>` }, fonts, shaper);
    expect(layout.pages.length).toBeGreaterThan(2);

    // PDF page count == layout pages.
    const pdf = await PDFDocument.load(await renderPdf(layout, fonts));
    expect(pdf.getPageCount()).toBe(layout.pages.length);

    // Preview page frames == layout pages; per-frame painted line count == layout lines.
    const { root: el } = renderPreview(layout);
    const frames = [...el.querySelectorAll('.rc-page')];
    expect(frames.length).toBe(layout.pages.length);
    frames.forEach((frame, i) => {
      const painted = frame.querySelectorAll('.rc-line').length;
      expect(painted).toBe(layout.pages[i]!.bodyLines.length);
    });

    // The shared result is the single source of truth for both painters.
    const pdfPages = pdf.getPageCount();
    const previewPages = frames.length;
    expect(previewPages).toBe(pdfPages);
  });
});

// ─── (b) never-suppress against the REAL compare pipeline ───────────────────────

/** Extract revision-text tokens by type from authored `<w:body>` markup, engine-independently. */
function authoredRevisionTokens(xml: string): Record<RcRevisionType, Set<string>> {
  const out: Record<RcRevisionType, Set<string>> = {
    ins: new Set(), del: new Set(), moveFrom: new Set(), moveTo: new Set(),
  };
  const decode = (s: string): string =>
    s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  const tokenise = (inner: string, into: Set<string>): void => {
    const text = decode(inner.replace(/<[^>]*>/g, ''));
    for (const tok of text.split(/\s+/)) if (tok.length) into.add(tok);
  };
  // \b after the tag name + a [\s>] lookahead excludes w:delText / w:moveFromRangeStart etc.
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

/** Decorated revision-run tokens by type, harvested from a laid-out result. */
function renderedRevisionTokens(layout: LayoutResult): {
  tokens: Record<RcRevisionType, Set<string>>; undecorated: number; total: number;
} {
  const tokens: Record<RcRevisionType, Set<string>> = {
    ins: new Set(), del: new Set(), moveFrom: new Set(), moveTo: new Set(),
  };
  let undecorated = 0, total = 0;
  const scan = (line: { runs: LayoutResult['pages'][number]['bodyLines'][number]['runs'] }): void => {
    for (const r of line.runs) {
      if (!r.revision) continue;
      total++;
      if (!r.decorations || r.decorations.length === 0) undecorated++;
      for (const tok of r.text.split(/\s+/)) if (tok.length) tokens[r.revision.type].add(tok);
    }
  };
  for (const p of layout.pages) {
    for (const l of p.bodyLines) scan(l);
    if (p.footnotes) for (const l of p.footnotes.lines) scan(l);
  }
  return { tokens, undecorated, total };
}

describe.skipIf(!smokeFixturesAvailable())('never-suppress against real Compare output (acceptance criterion 1)', () => {
  it('renders + decorates every native revision span the compare engine authored', async () => {
    const originalBundle = new PizZipBundle(docxBytes('original.docx'));
    const revisedBundle = new PizZipBundle(docxBytes('revised.docx'));
    const config = getCompareConfig(getActiveProfile());

    const result = await runPipeline({
      originalParts: extractParts(originalBundle),
      revisedParts: extractParts(revisedBundle),
      config,
      author: config.author,
      date: '2026-06-14T00:00:00Z',
    });

    // The redline is the ORIGINAL doc carrying revisions → its styles/numbering/footnotes.
    const origParts = extractParts(originalBundle);
    const layout = renderToLayout({
      authoredBodyXml: result.authoredBodyXml,
      ...(origParts.stylesXml ? { stylesXml: origParts.stylesXml } : {}),
      ...(origParts.footnotesXml ? { footnotesXml: origParts.footnotesXml } : {}),
      numbering: parseNumberingScheme(originalBundle),
    }, fonts, shaper);

    const authored = authoredRevisionTokens(result.authoredBodyXml);
    const authoredTotal = (Object.values(authored) as Set<string>[]).reduce((a, s) => a + s.size, 0);
    // The fixtures must actually contain changes, else the property test is vacuous.
    expect(authoredTotal).toBeGreaterThan(0);

    const { tokens: rendered, undecorated, total } = renderedRevisionTokens(layout);
    // Every rendered revision run is styled — nothing slips through unstyled (ADR-0008).
    expect(total).toBeGreaterThan(0);
    expect(undecorated).toBe(0);

    // Every authored revision token of each type appears in a decorated run of that type.
    for (const type of ['ins', 'del', 'moveFrom', 'moveTo'] as RcRevisionType[]) {
      for (const tok of authored[type]) {
        expect(
          rendered[type].has(tok),
          `${type} token "${tok}" authored by Compare but missing from the rendered redline`,
        ).toBe(true);
      }
    }
  });
});
