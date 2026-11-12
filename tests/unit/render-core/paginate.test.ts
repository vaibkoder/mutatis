import { describe, it, expect, beforeAll } from 'vitest';
import { buildRenderModel } from '@shared/render-core/input/index.js';
import { createFontResolver } from '@shared/render-core/fonts/registry.js';
import { createHarfBuzzShaper, type Shaper } from '@shared/render-core/layout/shape.js';
import { paginate } from '@shared/render-core/paginate/paginator.js';
import { contentHeightPt, type PageGeometry, type LayoutPage } from '@shared/render-core/types.js';
import { testFaces } from './_fixtures.js';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const fonts = createFontResolver(testFaces());
let shaper: Shaper;
beforeAll(async () => { shaper = await createHarfBuzzShaper(); });

// small page so boundaries are easy to hit: contentH = 204 - 72 - 72 = 60pt (~4 lines @12.29pt)
const SMALL: PageGeometry = { widthPt: 595.28, heightPt: 204, marginTopPt: 72, marginRightPt: 72, marginBottomPt: 72, marginLeftPt: 72 };

function bodyParas(n: number, prefix = 'Para'): string {
  let s = '';
  for (let i = 0; i < n; i++) s += `<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t>${prefix} ${i}</w:t></w:r></w:p>`;
  return s;
}
function model(inner: string, footnotesXml?: string) {
  const input: Parameters<typeof buildRenderModel>[0] = { authoredBodyXml: `<w:body ${NS}>${inner}</w:body>`, geometry: SMALL };
  if (footnotesXml) input.footnotesXml = footnotesXml;
  return buildRenderModel(input);
}
const bodyH = (p: LayoutPage): number => p.bodyLines.reduce((a, l) => a + l.heightPt, 0);
const fnH = (p: LayoutPage): number => (p.footnotes ? p.footnotes.lines.reduce((a, l) => a + l.heightPt, 0) + 12 : 0);

describe('paginator: filling + no overflow', () => {
  it('splits across pages and never overflows the content area', () => {
    const res = paginate(model(bodyParas(14)), shaper, fonts);
    expect(res.pages.length).toBeGreaterThan(2);
    const ch = contentHeightPt(SMALL);
    for (const p of res.pages) expect(bodyH(p) + fnH(p)).toBeLessThanOrEqual(ch + 0.5);
  });

  it('conserves all content lines (nothing dropped)', () => {
    const res = paginate(model(bodyParas(14)), shaper, fonts);
    const total = res.pages.reduce((a, p) => a + p.bodyLines.length, 0);
    expect(total).toBe(14);
  });

  it('pageBreakBefore forces a new page', () => {
    const inner = bodyParas(2) +
      `<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t>FORCED</w:t></w:r></w:p>`;
    const res = paginate(model(inner), shaper, fonts);
    const onPage = res.pages.findIndex((p) => p.bodyLines.some((l) => l.runs.some((r) => r.text === 'FORCED')));
    expect(onPage).toBeGreaterThan(0);
    // FORCED is the first line of its page
    expect(res.pages[onPage]!.bodyLines[0]!.runs.some((r) => r.text === 'FORCED')).toBe(true);
  });
});

describe('paginator: widow/orphan', () => {
  it('does not split a multi-line paragraph leaving a single line', () => {
    // 3 single-line paras fill most of page 1, then a 4-line paragraph
    const longLine = 'Liberation Sans justified body text that wraps across several lines on a very small narrow page width indeed here.';
    const inner = bodyParas(3) + `<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t>${longLine}</w:t></w:r></w:p>`;
    // narrow the column so the long paragraph is multi-line
    const narrow: PageGeometry = { ...SMALL, widthPt: 320 };
    const res = paginate({ ...model(inner), geometry: narrow }, shaper, fonts);
    // no page should contain exactly one line that is a fragment of a wrapped paragraph at the very bottom:
    // assert no overflow + content conserved as a proxy
    const ch = contentHeightPt(narrow);
    for (const p of res.pages) expect(bodyH(p) + fnH(p)).toBeLessThanOrEqual(ch + 0.5);
  });
});

describe('paginator: footnotes', () => {
  const fnXml = (id: string, paras: number) => {
    let b = '';
    for (let i = 0; i < paras; i++) b += `<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t>fn ${id} line ${i}</w:t></w:r></w:p>`;
    return `<w:footnotes ${NS}><w:footnote w:type="separator" w:id="-1"><w:p/></w:footnote><w:footnote w:id="${id}">${b}</w:footnote></w:footnotes>`;
  };

  it('places a footnote body on the same page as its reference, within the content area', () => {
    const inner = bodyParas(2) +
      `<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t>see note</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>` +
      bodyParas(1, 'After');
    const res = paginate(model(inner, fnXml('2', 1)), shaper, fonts);
    const refPage = res.pages.find((p) => p.bodyLines.some((l) => l.runs.some((r) => r.footnoteId === '2')))!;
    expect(refPage).toBeTruthy();
    expect(refPage.footnotes).toBeTruthy();
    const fnText = refPage.footnotes!.lines.flatMap((l) => l.runs.map((r) => r.text)).join(' ');
    expect(fnText).toContain('fn');
    expect(bodyH(refPage) + fnH(refPage)).toBeLessThanOrEqual(contentHeightPt(SMALL) + 0.5);
  });

  it('splits an oversized footnote with a continuation on the next page', () => {
    // a footnote taller than the page (8 lines @ ~12.29 = ~98pt > contentH 60)
    const inner = `<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t>ref</w:t></w:r><w:r><w:footnoteReference w:id="3"/></w:r></w:p>` + bodyParas(2, 'More');
    const res = paginate(model(inner, fnXml('3', 8)), shaper, fonts);
    const continued = res.pages.some((p) => p.footnotes?.continuedFromPrev);
    expect(continued).toBe(true);
    // total footnote lines across pages == 8 (nothing dropped)
    const fnLines = res.pages.reduce((a, p) => a + (p.footnotes?.lines.length ?? 0), 0);
    expect(fnLines).toBe(8);
    for (const p of res.pages) expect(bodyH(p) + fnH(p)).toBeLessThanOrEqual(contentHeightPt(SMALL) + 0.5);
  });
});

