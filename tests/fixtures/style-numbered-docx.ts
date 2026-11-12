/**
 * A `.docx` whose clause numbers live ENTIRELY in `styles.xml` + `numbering.xml`
 * — ADR-0042's motivating shape, shared by every seam that has to prove it
 * reaches the numbering context.
 *
 * `document.xml` here carries **no `<w:numPr>` at all**. That is not a
 * simplification: it is how a firm precedent is actually built, and it is why
 * two paragraphs at different list levels compared EQUAL on every field Compare
 * read before ADR-0042. A fixture that put `w:numPr` in the body would let a
 * seam pass its test while still dropping the context, because the paragraph
 * would carry its own level either way.
 *
 * Lives here rather than in one test file because four seams need it
 * (`pipeline`, `batch/runner`, `batch/driver`, `batch/pairing/sketch`) and three
 * of them are in separate files.
 *
 * The prose is drafting written for this suite. No client material.
 */
import { PizZipBundle } from '@shared/ooxml/bundle.js';
import { synthesiseDocx } from '@features/compare/io/text/synthesise-docx.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** `[styleId, text]` — the paragraph's style carries its numbering, the text carries none. */
export type StyledPara = readonly [styleId: string, text: string];

/** Two clause levels, `1.` and `(a)`, inherited through `w:pStyle` and nothing else. */
export const STYLE_NUMBERED_STYLES = `<w:styles ${W}>` +
  '<w:style w:type="paragraph" w:styleId="Clause1"><w:pPr>' +
  '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Clause2"><w:pPr>' +
  '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style>' +
  '</w:styles>';

/**
 * The list definitions, parameterised on the two tokens a seam guard needs to
 * vary — and ONLY those, so a test that changes one changes one.
 *
 * A fixture whose two sides carry the SAME `numbering.xml` cannot tell the two
 * numbering contexts apart: wire the original's context to the revised side and
 * every label comes out the same, so a redefinition reports no op and the suite
 * stays green. That is a suppression (ADR-0008) no count assertion can see.
 * Pass a different scheme to one side to make the two contexts distinguishable.
 *
 *  - `headFmt` — level 0's `w:numFmt`. A counterparty's template redefining
 *    `decimal` as `upperRoman` is the motivating edit.
 *  - `subText` — level 1's `w:lvlText`. `(%2)` shows the own counter only;
 *    `%1.%2` is the `1.1` house style, under which an ancestor's format reaches
 *    every descendant's label.
 */
export function styleNumberedNumbering(
  opts: { headFmt?: string; subText?: string } = {},
): string {
  const { headFmt = 'decimal', subText = '(%2)' } = opts;
  return `<w:numbering ${W}><w:abstractNum w:abstractNumId="0">` +
    `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${headFmt}"/>` +
    '<w:lvlText w:val="%1."/></w:lvl>' +
    '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/>' +
    `<w:lvlText w:val="${subText}"/></w:lvl>` +
    '</w:abstractNum><w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num></w:numbering>';
}

/** The default scheme — `1.` over `(a)`. Byte-identical to what it always was. */
export const STYLE_NUMBERED_NUMBERING = styleNumberedNumbering();

/** The package, built on `synthesiseDocx` so content types and rels are real. */
export function styleNumberedBundle(
  paras: readonly StyledPara[],
  numberingXml: string = STYLE_NUMBERED_NUMBERING,
): PizZipBundle {
  const bundle = synthesiseDocx('placeholder');
  bundle.setPart(
    'word/document.xml',
    `<w:document ${W}><w:body>` +
      paras
        .map(
          ([styleId, text]) =>
            `<w:p><w:pPr><w:pStyle w:val="${styleId}"/></w:pPr>` +
            `<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`,
        )
        .join('') +
      '</w:body></w:document>',
  );
  bundle.setPart('word/styles.xml', STYLE_NUMBERED_STYLES);
  bundle.setPart('word/numbering.xml', numberingXml);
  return bundle;
}

/** The same package as raw bytes. */
export function styleNumberedDocx(
  paras: readonly StyledPara[],
  numberingXml?: string,
): ArrayBuffer {
  const bytes = new Uint8Array(styleNumberedBundle(paras, numberingXml).toUint8Array());
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

/** The same package as a `File`, for the driver seams that hold file handles. */
export function styleNumberedFile(
  name: string,
  paras: readonly StyledPara[],
  numberingXml?: string,
): File {
  return new File([new Uint8Array(styleNumberedBundle(paras, numberingXml).toUint8Array())], name, {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

/** Clauses used across the seam guards, so each one's edit is the only variable. */
export const HEAD: StyledPara = ['Clause1', 'The Seller covenants with the Buyer as follows.'];
export const SUB_A: StyledPara = ['Clause2', 'It shall maintain the insurances listed in the Schedule.'];
export const SUB_B: StyledPara = ['Clause2', 'It shall not amend the Disclosure Letter without consent.'];
export const INSERTED: StyledPara = ['Clause2', 'It shall keep the Records for six years.'];
