import { describe, it, expect, beforeAll } from 'vitest';
import {
  renderToLayout, createFontResolver, createHarfBuzzShaper, DEFAULT_THEME,
  type Shaper, type LayoutResult, type PositionedRun,
} from '@shared/render-core/index.js';
import { testFaces } from './_fixtures.js';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const fonts = createFontResolver(testFaces());
let shaper: Shaper;
beforeAll(async () => { shaper = await createHarfBuzzShaper(); });

function allRevisionRuns(layout: LayoutResult): PositionedRun[] {
  const out: PositionedRun[] = [];
  for (const p of layout.pages) {
    for (const l of p.bodyLines) for (const r of l.runs) if (r.revision) out.push(r);
    if (p.footnotes) for (const l of p.footnotes.lines) for (const r of l.runs) if (r.revision) out.push(r);
  }
  return out;
}

describe('never-suppress (acceptance criterion 1)', () => {
  it('renders AND decorates every ins/del/moveFrom/moveTo span in the change-set', () => {
    // Build a paragraph that interleaves equal text with 40 revision spans.
    const wrap = (i: number): string => {
      const word = `w${i}`;
      switch (i % 4) {
        case 0: return `<w:ins w:id="${i}"><w:r><w:t>${word}</w:t></w:r></w:ins>`;
        case 1: return `<w:del w:id="${i}"><w:r><w:delText>${word}</w:delText></w:r></w:del>`;
        case 2: return `<w:moveTo w:id="${i}"><w:r><w:t>${word}</w:t></w:r></w:moveTo>`;
        default: return `<w:moveFrom w:id="${i}"><w:r><w:delText>${word}</w:delText></w:r></w:moveFrom>`;
      }
    };
    let inner = '<w:p>';
    const expected = new Set<string>();
    for (let i = 0; i < 40; i++) {
      inner += `<w:r><w:t>eq${i} </w:t></w:r>` + wrap(i) + '<w:r><w:t> </w:t></w:r>';
      expected.add(`w${i}`);
    }
    inner += '</w:p>';

    const layout = renderToLayout({ authoredBodyXml: `<w:body ${NS}>${inner}</w:body>` }, fonts, shaper);

    const revRuns = allRevisionRuns(layout);
    // every revision run carries a decoration (never silently dropped/unstyled)
    expect(revRuns.length).toBeGreaterThanOrEqual(40);
    for (const r of revRuns) {
      expect(r.decorations && r.decorations.length > 0).toBe(true);
    }
    // every change-set token appears in the rendered output
    const seen = new Set(revRuns.map((r) => r.text));
    for (const w of expected) expect(seen.has(w)).toBe(true);
  });

  it('decorates by type: ins→underline, del→strike, with the configured theme colour', () => {
    const inner = '<w:p>' +
      '<w:ins w:id="1"><w:r><w:t>added</w:t></w:r></w:ins>' +
      '<w:del w:id="2"><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>';
    const layout = renderToLayout({ authoredBodyXml: `<w:body ${NS}>${inner}</w:body>` }, fonts, shaper, { theme: DEFAULT_THEME });
    const runs = allRevisionRuns(layout);
    const added = runs.find((r) => r.text === 'added')!;
    const gone = runs.find((r) => r.text === 'gone')!;
    expect(added.decorations!.some((d) => d.kind === 'insUnderline' && d.color === DEFAULT_THEME.insertColor)).toBe(true);
    expect(gone.decorations!.some((d) => d.kind === 'delStrike' && d.color === DEFAULT_THEME.deleteColor)).toBe(true);
    // changed lines flagged for a margin change bar
    expect(layout.pages[0]!.bodyLines.some((l) => l.changeBar)).toBe(true);
    // insertion highlight realised as a page rect
    expect(layout.pages[0]!.rects.some((r) => r.fill)).toBe(true);
  });

  it('merges adjacent same-revision runs into ONE continuous rule + rect (conventional redline look, bridges word gaps)', () => {
    // Two inserted words on one line: each word is its own positioned run (the
    // space is folded into an x-gap), so naive per-run painting would draw two
    // separate underlines + two separate highlight rects with a gap between.
    const inner = '<w:p><w:ins w:id="1"><w:r><w:t xml:space="preserve">alpha beta</w:t></w:r></w:ins></w:p>';
    const layout = renderToLayout({ authoredBodyXml: `<w:body ${NS}>${inner}</w:body>` }, fonts, shaper, { theme: DEFAULT_THEME });
    const line = layout.pages[0]!.bodyLines[0]!;

    const wordRuns = line.runs.filter((r) => r.revision);
    expect(wordRuns.length).toBe(2); // alpha, beta — separate runs

    // ONE underline segment spanning from the first word's start to the last word's end.
    expect(line.decoSegs).toBeDefined();
    expect(line.decoSegs!.length).toBe(1);
    const seg = line.decoSegs![0]!;
    expect(seg.kind).toBe('insUnderline');
    expect(seg.x1Pt).toBeCloseTo(wordRuns[0]!.xPt, 3);
    expect(seg.x2Pt).toBeCloseTo(wordRuns[1]!.xPt + wordRuns[1]!.widthPt, 3);
    // It bridges the inter-word gap (wider than either word alone).
    expect(seg.x2Pt - seg.x1Pt).toBeGreaterThan(wordRuns[0]!.widthPt);

    // ONE continuous highlight rect over the same span (not one box per word).
    const fills = layout.pages[0]!.rects.filter((r) => r.fill);
    expect(fills.length).toBe(1);
    expect(fills[0]!.xPt).toBeCloseTo(seg.x1Pt, 3);
    expect(fills[0]!.widthPt).toBeCloseTo(seg.x2Pt - seg.x1Pt, 3);
  });

  it('does NOT merge across a different revision type (del then ins → two rules)', () => {
    const inner = '<w:p>' +
      '<w:del w:id="1"><w:r><w:delText xml:space="preserve">old </w:delText></w:r></w:del>' +
      '<w:ins w:id="2"><w:r><w:t>new</w:t></w:r></w:ins></w:p>';
    const layout = renderToLayout({ authoredBodyXml: `<w:body ${NS}>${inner}</w:body>` }, fonts, shaper, { theme: DEFAULT_THEME });
    const segs = layout.pages[0]!.bodyLines[0]!.decoSegs!;
    expect(segs.map((s) => s.kind).sort()).toEqual(['delStrike', 'insUnderline']);
  });

  it('emits the move anchor ONCE per move, not once per word', () => {
    // A moveFrom block of several words shares one changeId. Each word is its own
    // positioned run, so a naive per-run anchor litters the block with '*'s
    // (regression: "Conduct*of*Business*"). The glyph must appear exactly once.
    const inner = '<w:p><w:moveFrom w:id="7" w:name="m1">' +
      '<w:r><w:delText xml:space="preserve">Conduct of Business Pending Completion</w:delText></w:r>' +
      '</w:moveFrom></w:p>';
    const layout = renderToLayout({ authoredBodyXml: `<w:body ${NS}>${inner}</w:body>` }, fonts, shaper, { theme: DEFAULT_THEME });
    const moveRuns = allRevisionRuns(layout).filter((r) => r.revision!.type === 'moveFrom');
    expect(moveRuns.length).toBeGreaterThan(1); // several word runs
    const anchors = moveRuns.filter((r) => r.decorations?.some((d) => d.kind === 'moveAnchor' || d.kind === 'changeAnchor'));
    expect(anchors.length).toBe(1);
    // every move run still carries its strike (never-suppress intact)
    for (const r of moveRuns) expect(r.decorations!.some((d) => d.kind === 'delStrike')).toBe(true);
  });

  it('respects theme config: highlightInserts off ⇒ no highlight decoration/rect', () => {
    const inner = '<w:p><w:ins w:id="1"><w:r><w:t>added</w:t></w:r></w:ins></w:p>';
    const theme = { ...DEFAULT_THEME, highlightInserts: false };
    const layout = renderToLayout({ authoredBodyXml: `<w:body ${NS}>${inner}</w:body>` }, fonts, shaper, { theme });
    const added = allRevisionRuns(layout).find((r) => r.text === 'added')!;
    expect(added.decorations!.some((d) => d.kind === 'highlight')).toBe(false);
    expect(layout.pages[0]!.rects.length).toBe(0);
  });
});
