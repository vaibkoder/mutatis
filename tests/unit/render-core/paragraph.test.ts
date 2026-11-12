import { describe, it, expect, beforeAll } from 'vitest';
import { buildRenderModel } from '@shared/render-core/input/index.js';
import { parseNumberingScheme } from '@shared/numbering/index.js';
import type { DocxBundle } from '@shared/ooxml/types.js';
import { createFontResolver } from '@shared/render-core/fonts/registry.js';
import { createHarfBuzzShaper, type Shaper } from '@shared/render-core/layout/shape.js';
import { layoutParagraph } from '@shared/render-core/layout/paragraph.js';
import type { RcParagraph } from '@shared/render-core/types.js';
import { testFaces } from './_fixtures.js';

const COL = 451.28; // A4 content width @1in margins
const fonts = createFontResolver(testFaces());
let shaper: Shaper;
beforeAll(async () => { shaper = await createHarfBuzzShaper(); });

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
function firstPara(bodyInner: string): RcParagraph {
  const m = buildRenderModel({ authoredBodyXml: `<w:body ${NS}>${bodyInner}</w:body>` });
  const b = m.blocks.find((x) => x.kind === 'paragraph')!;
  return (b as { paragraph: RcParagraph }).paragraph;
}
const lineEnd = (runs: { xPt: number; widthPt: number }[]): number =>
  runs.reduce((mx, r) => Math.max(mx, r.xPt + r.widthPt), 0);

describe('layoutParagraph: justification', () => {
  const LONG =
    'The Borrower shall on each Interest Payment Date and in accordance with Clause 8.2 ' +
    'pay to the Agent for the account of each Lender the accrued interest on the Loan which ' +
    'has accrued during the Interest Period ending on that date provided that no such payment ' +
    'shall be required to the extent it would exceed the Total Commitments then in force.';

  it('fills the column on every non-last line and never overflows', () => {
    const p = firstPara(`<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t>${LONG}</w:t></w:r></w:p>`);
    const { lines } = layoutParagraph(p, COL, shaper, fonts);
    expect(lines.length).toBeGreaterThan(2);
    lines.forEach((ln, idx) => {
      const end = lineEnd(ln.runs);
      expect(end).toBeLessThanOrEqual(COL + 0.5); // no overflow
      if (idx < lines.length - 1) expect(end).toBeCloseTo(COL, 0); // justified flush
    });
  });

  it('left-aligned (default) does NOT stretch to the column', () => {
    const p = firstPara(`<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t>${LONG}</w:t></w:r></w:p>`);
    const { lines } = layoutParagraph(p, COL, shaper, fonts);
    const firstEnd = lineEnd(lines[0]!.runs);
    expect(firstEnd).toBeLessThan(COL);          // ragged right
    expect(firstEnd).toBeGreaterThan(COL * 0.8); // but a full-ish line
  });
});

describe('layoutParagraph: formatting + alignment + line height', () => {
  it('preserves bold/italic on positioned runs', () => {
    const p = firstPara('<w:p><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>strong</w:t></w:r></w:p>');
    const { lines } = layoutParagraph(p, COL, shaper, fonts);
    const run = lines[0]!.runs.find((r) => r.text === 'strong')!;
    expect(run.rpr.bold).toBe(true);
    expect(run.rpr.italic).toBe(true);
  });

  it('right alignment pushes the line to the right margin', () => {
    const p = firstPara('<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t>short line</w:t></w:r></w:p>');
    const { lines } = layoutParagraph(p, COL, shaper, fonts);
    expect(lineEnd(lines[0]!.runs)).toBeCloseTo(COL, 0);
    expect(lines[0]!.runs[0]!.xPt).toBeGreaterThan(COL * 0.5);
  });

  it('line height follows the usWin formula (~12.29pt @11pt Liberation)', () => {
    const p = firstPara('<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="22"/></w:rPr><w:t>x</w:t></w:r></w:p>');
    const { lines } = layoutParagraph(p, COL, shaper, fonts);
    expect(lines[0]!.heightPt).toBeCloseTo(12.289, 1);
  });
});

describe('layoutParagraph: hanging indent + numbering label', () => {
  it('hangs the number label left of the body text', () => {
    // indent left 36pt (720tw), hanging 18pt (360tw); body text at 36pt, label at 18pt
    const m = buildRenderModel({
      authoredBodyXml: `<w:body ${NS}><w:p><w:pPr><w:ind w:left="720" w:hanging="360"/>` +
        `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>` +
        `<w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t>Clause body text here.</w:t></w:r></w:p></w:body>`,
      numbering: parseNumberingScheme({
        part: (pp: string) => (pp === 'word/numbering.xml'
          ? `<w:numbering ${NS}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`
          : undefined),
      } as unknown as DocxBundle),
    });
    const p = (m.blocks[0] as { paragraph: RcParagraph }).paragraph;
    expect(p.numberLabel).toBe('1.');
    const { lines } = layoutParagraph(p, COL, shaper, fonts);
    const label = lines[0]!.runs.find((r) => r.text === '1.')!;
    const firstWord = lines[0]!.runs.find((r) => r.text === 'Clause')!;
    expect(label.xPt).toBeCloseTo(18, 0);       // hung at indentLeft - hanging
    expect(firstWord.xPt).toBeCloseTo(36, 0);   // body at indentLeft
  });

  it('does NOT overlap label and text when there is no hanging gutter', () => {
    // Numbering-level indents are a deferred bound: a numbered paragraph
    // without its own w:ind has hangingPt = 0. The label must still not paint
    // on top of the first text cluster (regression: "7.1" over "The Seller").
    const m = buildRenderModel({
      authoredBodyXml: `<w:body ${NS}><w:p><w:pPr>` +
        `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>` +
        `<w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t>The Seller shall not be liable.</w:t></w:r></w:p></w:body>`,
      numbering: parseNumberingScheme({
        part: (pp: string) => (pp === 'word/numbering.xml'
          ? `<w:numbering ${NS}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`
          : undefined),
      } as unknown as DocxBundle),
    });
    const p = (m.blocks[0] as { paragraph: RcParagraph }).paragraph;
    expect(p.numberLabel).toBe('1.');
    const { lines } = layoutParagraph(p, COL, shaper, fonts);
    const label = lines[0]!.runs.find((r) => r.text === '1.')!;
    const firstWord = lines[0]!.runs.find((r) => r.text === 'The')!;
    // text begins strictly after the label's right edge (no collision)
    expect(firstWord.xPt).toBeGreaterThanOrEqual(label.xPt + label.widthPt);
  });
});
