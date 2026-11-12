import { describe, it, expect, beforeAll } from 'vitest';
import {
  renderToLayout, buildRenderModel, createFontResolver, createHarfBuzzShaper,
  type Shaper, type LayoutResult, type RcBlock,
} from '@shared/render-core/index.js';
import { layoutTable, splitRow } from '@shared/render-core/paginate/tables.js';
import { testFaces } from './_fixtures.js';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const fonts = createFontResolver(testFaces());
let shaper: Shaper;
beforeAll(async () => { shaper = await createHarfBuzzShaper(); });

const cell = (text: string, opts: { span?: number } = {}): string =>
  `<w:tc>${opts.span ? `<w:tcPr><w:gridSpan w:val="${opts.span}"/></w:tcPr>` : ''}<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;
const grid = (...cols: number[]): string =>
  `<w:tblGrid>${cols.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;
const rowXml = (cells: string, cantSplit = false): string =>
  `<w:tr>${cantSplit ? '<w:trPr><w:cantSplit/></w:trPr>' : ''}${cells}</w:tr>`;
const tableXml = (gridXml: string, rows: string): string => `<w:tbl>${gridXml}${rows}</w:tbl>`;
const body = (inner: string): string => `<w:body ${NS}>${inner}</w:body>`;

function firstTable(xml: string): RcBlock & { kind: 'table' } {
  const model = buildRenderModel({ authoredBodyXml: xml });
  const t = model.blocks.find((b): b is RcBlock & { kind: 'table' } => b.kind === 'table');
  if (!t) throw new Error('no table parsed');
  return t;
}

const allBodyRuns = (layout: LayoutResult, pageIdx?: number) =>
  (pageIdx === undefined ? layout.pages : [layout.pages[pageIdx]!])
    .flatMap((p) => p.bodyLines).flatMap((l) => l.runs);

describe('table layout (tblGrid column widths + cell wrap)', () => {
  it('places cells at their tblGrid column offsets and wraps long cell text', () => {
    // 2 cols × 144pt (2880 twips). Cell 1 is long (wraps); cell 2 is a marker.
    const c1 = 'alpha beta gamma delta epsilon zeta eta theta iota kappa lambda';
    const xml = body(tableXml(grid(2880, 2880), rowXml(cell(c1) + cell('MARKER'))));

    const tbl = firstTable(xml);
    const tl = layoutTable(tbl.table, 451, shaper, fonts);
    expect(tl.rows.length).toBe(1);

    // Cell 2 ('MARKER') sits in the second column (x ≈ 144 + pad), right of cell 1.
    const markerLine = tl.rows[0]!.cellLines.find((cl) => cl.line.runs.some((r) => r.text === 'MARKER'))!;
    const markerX = markerLine.line.runs.find((r) => r.text === 'MARKER')!.xPt;
    expect(markerX).toBeGreaterThan(140);
    expect(markerX).toBeLessThan(160);

    // Cell 1 wrapped onto multiple lines (more than the single MARKER line).
    const cell1Lines = tl.rows[0]!.cellLines.filter((cl) => cl.line.runs.some((r) => /alpha|kappa|lambda/.test(r.text)));
    expect(cell1Lines.length).toBeGreaterThan(1);
    expect(tl.rows[0]!.heightPt).toBeGreaterThan(0);
  });

  it('emits a layout (with border rects) through the full paginate path', () => {
    const xml = body(tableXml(grid(2880, 2880), rowXml(cell('left text') + cell('right text'))));
    const layout = renderToLayout({ authoredBodyXml: xml }, fonts, shaper);
    expect(layout.pages.length).toBe(1);
    // border rects present (stroked, not filled)
    expect(layout.pages[0]!.rects.some((r) => r.stroke)).toBe(true);
    // both cells rendered, second to the right of the first
    const right = allBodyRuns(layout).find((r) => r.text === 'right')!;
    const left = allBodyRuns(layout).find((r) => r.text === 'left')!;
    expect(right.xPt).toBeGreaterThan(left.xPt + 100);
  });
});

describe('table pagination (row split + cantSplit + conservation)', () => {
  it('splits a long table across pages without dropping any row', () => {
    const rows = Array.from({ length: 70 }, (_, i) => rowXml(cell(`row ${i} cell content here`))).join('');
    const xml = body(tableXml(grid(8000), rows));
    const layout = renderToLayout({ authoredBodyXml: xml }, fonts, shaper);
    expect(layout.pages.length).toBeGreaterThan(1);
    // every row's single line is present somewhere (content conserved, never dropped)
    const texts = allBodyRuns(layout).map((r) => r.text).join(' ');
    for (let i = 0; i < 70; i++) expect(texts).toContain(`${i}`);
  });

  it('keeps a cantSplit row whole — all its lines land on one page', () => {
    // Filler to push the next row near the page bottom, then a tall cantSplit row.
    const filler = Array.from({ length: 44 }, () => '<w:p><w:r><w:t>filler line of body text here</w:t></w:r></w:p>').join('');
    const tall = 'SPLITME one two three four five six seven eight nine ten eleven twelve thirteen fourteen fifteen sixteen seventeen eighteen';
    const tbl = tableXml(grid(8000), rowXml(cell(tall), true));
    const layout = renderToLayout({ authoredBodyXml: body(filler + tbl) }, fonts, shaper);

    // Which pages carry a 'SPLITME' line? Must be exactly one (row not split).
    const pagesWithRow = layout.pages
      .map((p, idx) => ({ idx, has: p.bodyLines.some((l) => l.runs.some((r) => r.text === 'SPLITME')) }))
      .filter((x) => x.has)
      .map((x) => x.idx);
    expect(pagesWithRow.length).toBe(1);
  });
});

describe('graceful degradation', () => {
  it('handles a table with no tblGrid (even columns) and a gridSpan cell without crashing', () => {
    const xml = body(tableXml(
      '', // no tblGrid → even columns
      rowXml(cell('cellA') + cell('cellB')) + rowXml(cell('spanning header', { span: 2 })),
    ));
    const layout = renderToLayout({ authoredBodyXml: xml }, fonts, shaper);
    const texts = allBodyRuns(layout).map((r) => r.text);
    expect(texts).toContain('cellA');
    expect(texts).toContain('cellB');
    expect(texts).toContain('spanning');
  });

  it('splitRow keeps at least one line and rebases the continuation to the top', () => {
    const xml = body(tableXml(grid(8000), rowXml(cell('one two three four five six seven eight nine ten eleven twelve'))));
    const row = layoutTable(firstTable(xml).table, 200, shaper, fonts).rows[0]!;
    expect(row.cellLines.length).toBeGreaterThan(1);
    const [head, tail] = splitRow(row, row.cellLines[0]!.relTopPt + row.cellLines[0]!.line.heightPt + 1);
    expect(head.cellLines.length).toBeGreaterThanOrEqual(1);
    expect(tail).not.toBeNull();
    // continuation rebased near the top of its own row
    expect(Math.min(...tail!.cellLines.map((c) => c.relTopPt))).toBeLessThanOrEqual(head.cellLines[0]!.relTopPt + 0.001);
  });
});
