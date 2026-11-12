import { describe, it, expect } from 'vitest';
import { buildRenderModel } from '@shared/render-core/input/index.js';
import { parseNumberingScheme } from '@shared/numbering/index.js';
import type { DocxBundle } from '@shared/ooxml/types.js';
import type { RcParagraph } from '@shared/render-core/types.js';

const STYLES = `<?xml version="1.0"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Bold"><w:name w:val="Bold"/><w:basedOn w:val="Normal"/><w:rPr><w:b/></w:rPr></w:style>
</w:styles>`;

const NUMBERING = `<?xml version="1.0"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`;

function body(inner: string, sectPr = ''): string {
  return `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${inner}${sectPr}</w:body>`;
}
function scheme() {
  const bundle = { part: (p: string) => (p === 'word/numbering.xml' ? NUMBERING : undefined) } as unknown as DocxBundle;
  return parseNumberingScheme(bundle);
}
const paras = (m: ReturnType<typeof buildRenderModel>): RcParagraph[] =>
  m.blocks.filter((b) => b.kind === 'paragraph').map((b) => (b as { paragraph: RcParagraph }).paragraph);

describe('input: style cascade', () => {
  it('resolves docDefaults + basedOn chain (Heading1 → bold, 14pt, heading, keepNext)', () => {
    const m = buildRenderModel({ stylesXml: STYLES, authoredBodyXml: body('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>') });
    const p = paras(m)[0]!;
    expect(p.isHeading).toBe(true);
    expect(p.ppr.keepNext).toBe(true);
    expect(p.markRpr.bold).toBe(true);
    expect(p.markRpr.sizePt).toBe(14);
    expect(p.runs[0]!.rpr.fontFamily).toBe('Times New Roman'); // from docDefaults
  });

  it('toggle XOR: bold style + direct <w:b/> ⇒ NOT bold; explicit val=0 ⇒ hard off', () => {
    const m = buildRenderModel({ stylesXml: STYLES, authoredBodyXml: body(
      '<w:p><w:pPr><w:pStyle w:val="Bold"/></w:pPr>' +
      '<w:r><w:t>inherited</w:t></w:r>' +
      '<w:r><w:rPr><w:b/></w:rPr><w:t>toggled</w:t></w:r>' +
      '<w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t>off</w:t></w:r></w:p>') });
    const r = paras(m)[0]!.runs;
    expect(r[0]!.rpr.bold).toBe(true);   // inherits style bold
    expect(r[1]!.rpr.bold).toBe(false);  // XOR with inherited true
    expect(r[2]!.rpr.bold).toBe(false);  // absolute off
  });
});

describe('input: revisions + footnotes + images', () => {
  it('tags ins/del runs and reads delText', () => {
    const m = buildRenderModel({ authoredBodyXml: body(
      '<w:p><w:ins w:id="5" w:author="A"><w:r><w:t>added</w:t></w:r></w:ins>' +
      '<w:del w:id="6"><w:r><w:delText>removed</w:delText></w:r></w:del></w:p>') });
    const r = paras(m)[0]!.runs;
    expect(r[0]!.revision?.type).toBe('ins');
    expect(r[0]!.revision?.author).toBe('A');
    expect(r[0]!.text).toBe('added');
    expect(r[1]!.revision?.type).toBe('del');
    expect(r[1]!.text).toBe('removed');
  });

  it('links footnote references to footnote bodies; image → placeholder', () => {
    const m = buildRenderModel({
      authoredBodyXml: body('<w:p><w:r><w:t>see</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r><w:r><w:drawing/></w:r></w:p>'),
      footnotesXml: `<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:footnote w:type="separator" w:id="-1"><w:p/></w:footnote>
        <w:footnote w:id="2"><w:p><w:r><w:t>the note</w:t></w:r></w:p></w:footnote></w:footnotes>`,
    });
    const r = paras(m)[0]!.runs;
    const ref = r.find((x) => x.footnoteId === '2')!;
    expect(ref.rpr.superscript).toBe(true);
    expect(r.some((x) => x.imagePlaceholder)).toBe(true);
    expect(m.footnotes.has('2')).toBe(true);
    expect(m.footnotes.get('2')!.blocks.length).toBe(1);
    expect(m.footnotes.size).toBe(1); // separator skipped
  });
});

describe('input: class annotations (classByRevId → RcRun.changeClasses)', () => {
  it('attaches classes to runs by native @w:id; non-matching/unmapped runs stay unset', () => {
    const m = buildRenderModel({
      authoredBodyXml: body(
        '<w:p><w:ins w:id="5"><w:r><w:t>added</w:t></w:r></w:ins>' +
        '<w:del w:id="6"><w:r><w:delText>removed</w:delText></w:r></w:del>' +
        '<w:r><w:t>plain</w:t></w:r></w:p>'),
      classByRevId: new Map([
        ['5', ['definedTerm', 'substantiveText']],
        ['6', ['numericalValue']],
      ]),
    });
    const r = paras(m)[0]!.runs;
    expect(r[0]!.changeClasses).toEqual(['definedTerm', 'substantiveText']);
    expect(r[1]!.changeClasses).toEqual(['numericalValue']);
    expect(r[2]!.changeClasses).toBeUndefined(); // unrevised run — never annotated
  });

  it('never drops a revision when no class info is present (ADR-0008)', () => {
    const xml = body(
      '<w:p><w:ins w:id="5"><w:r><w:t>added</w:t></w:r></w:ins>' +
      '<w:del w:id="6"><w:r><w:delText>removed</w:delText></w:r></w:del></w:p>');
    const withClasses = buildRenderModel({ authoredBodyXml: xml, classByRevId: new Map([['5', ['structural']]]) });
    const without = buildRenderModel({ authoredBodyXml: xml });
    // Every native revision renders in BOTH cases; classes are pure annotation.
    expect(paras(withClasses)[0]!.runs.map((x) => x.revision?.type)).toEqual(['ins', 'del']);
    expect(paras(without)[0]!.runs.map((x) => x.revision?.type)).toEqual(['ins', 'del']);
    expect(without.blocks.length).toBe(withClasses.blocks.length);
  });
});

describe('input: numbering labels + geometry + tables', () => {
  it('attaches rendered numbering labels from the shared resolver', () => {
    const m = buildRenderModel({
      authoredBodyXml: body(
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>first</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>second</w:t></w:r></w:p>'),
      numbering: scheme(),
    });
    const ps = paras(m);
    expect(ps[0]!.numberLabel).toBe('1.');
    expect(ps[1]!.numberLabel).toBe('2.');
  });

  it('reads A4 geometry from sectPr and defaults to A4 when absent', () => {
    const withSect = buildRenderModel({ authoredBodyXml: body('<w:p><w:r><w:t>x</w:t></w:r></w:p>',
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>') });
    expect(withSect.geometry.widthPt).toBeCloseTo(595.3, 0);
    expect(withSect.geometry.marginTopPt).toBe(72);
    const noSect = buildRenderModel({ authoredBodyXml: body('<w:p><w:r><w:t>x</w:t></w:r></w:p>') });
    expect(noSect.geometry.heightPt).toBeCloseTo(841.89, 1);
  });

  it('parses a basic table (grid cols + cells)', () => {
    const m = buildRenderModel({ authoredBodyXml: body(
      '<w:tbl><w:tblGrid><w:gridCol w:w="2880"/><w:gridCol w:w="2880"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr></w:tbl>') });
    const tbl = m.blocks.find((b) => b.kind === 'table');
    expect(tbl).toBeTruthy();
    if (tbl?.kind === 'table') {
      expect(tbl.table.gridColsPt).toEqual([144, 144]); // 2880 twips = 144pt
      expect(tbl.table.rows[0]!.cells.length).toBe(2);
    }
  });
});
