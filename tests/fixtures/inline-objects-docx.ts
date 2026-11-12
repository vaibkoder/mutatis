/**
 * Synthetic paragraphs carrying content that is not plain text — an equation,
 * a footnote reference, a hyperlink, a complex field, a bookmark, a text box —
 * for the ADR-0045 regression tests.
 *
 * Every string here is invented. No real document's text, ids or internals
 *: the bookmark and relationship ids are generic
 * placeholders of the shape Word writes, not ids taken from any file.
 */
import type { CompareParts } from '@features/compare/io/index.js';

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml"';

/** A plain text run. */
export const run = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

/** A paragraph around inline content. */
export const para = (inner: string): string => `<w:p>${inner}</w:p>`;

/** An inline Office Math (OMML) equation holding one math run. */
export const math = (text: string): string => `<m:oMath><m:r><m:t>${text}</m:t></m:r></m:oMath>`;

/** A footnote-reference run. */
export const footnoteRef = (id = '1'): string =>
  `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;

/** An endnote-reference run. */
export const endnoteRef = (id = '1'): string =>
  `<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="${id}"/></w:r>`;

/** An inline picture: a DrawingML `pic:pic` whose image is the part behind `rId`. */
export const inlinePicture = (rId = 'rId20'): string =>
  '<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><a:graphic>' +
  '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
  `<pic:blipFill><a:blip r:embed="${rId}"/></pic:blipFill>` +
  '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';

/** An embedded (OLE) object run, drawn by its VML shape. */
export const embeddedObject = (): string => '<w:r><w:object><v:shape id="_x0000_i1025"/></w:object></w:r>';

/** An external hyperlink whose one run carries the Hyperlink character style. */
export const hyperlink = (text: string, rId = 'rId9'): string =>
  `<w:hyperlink r:id="${rId}"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>` +
  `<w:t xml:space="preserve">${text}</w:t></w:r></w:hyperlink>`;

/** A complex REF field (begin / instruction / separate / result / end). */
export const refField = (result: string, bookmark = '_Ref10001'): string =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> REF ${bookmark} \\r \\h </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `<w:r><w:t>${result}</w:t></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

/** A bookmark pair around inline content. */
export const bookmarked = (inner: string, name = '_Ref10001'): string =>
  `<w:bookmarkStart w:id="0" w:name="${name}"/>${inner}<w:bookmarkEnd w:id="0"/>`;

/** An anchored DrawingML text box holding one paragraph. */
export const textBox = (text: string): string =>
  '<w:r><w:drawing><wp:anchor><a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent>' +
  `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>` +
  '</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';

/**
 * The same text box as Word usually writes it: a DrawingML choice and a VML
 * fallback, each holding its own copy of the box's paragraph.
 */
export const alternateTextBox = (text: string): string =>
  '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor><a:graphic><a:graphicData>' +
  `<wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp>` +
  '</a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape><v:textbox>' +
  `<w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent>` +
  '</v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>';

/** A closing paragraph, so the pair under test is never the last of its part (ADR-0035). */
const TAIL = para(run('This Agreement ends here.'));

/** `CompareParts` for a body of paragraphs — no styles, numbering or auxiliary parts. */
export function inlineParts(...paragraphs: string[]): CompareParts {
  return {
    documentXml:
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}>` +
      `<w:body>${paragraphs.join('')}${TAIL}</w:body></w:document>`,
    headers: new Map(),
    footers: new Map(),
    roles: [],
    sectionCount: 1,
    footnotesXml: null,
    endnotesXml: null,
    commentsXml: null,
    stylesXml: null,
    numberingXml: null,
  };
}
