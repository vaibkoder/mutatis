import PizZip from 'pizzip';
import type { Block, Para, Run, Table } from './model.js';

/**
 * Blocks → a minimal, valid .docx, built in memory.
 *
 * The compare engine reads OOXML, so a pasted text becomes a real document and
 * rides the existing pipeline untouched. Only what the blocks carry is written:
 * bold/italic/underline/strike/sup/sub runs, Heading1-6 styles, alignment, left
 * indent and bordered tables. Element order inside `w:pPr`/`w:rPr` follows the
 * schema; Word is strict about it.
 */

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CONTENT_WIDTH_TWIPS = 9026;
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const CONTENT_TYPES =
  `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
  '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
  '</Types>';

const PACKAGE_RELS =
  `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  '</Relationships>';

const DOCUMENT_RELS =
  `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
  '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>' +
  '</Relationships>';

const SETTINGS = `${XML}<w:settings xmlns:w="${W_NS}"/>`;

const HEADING_SIZES = [32, 28, 26, 24, 22, 22]; // half-points

function headingStyles(): string {
  return HEADING_SIZES.map(
    (sz, i) =>
      `<w:style w:type="paragraph" w:styleId="Heading${i + 1}"><w:name w:val="heading ${i + 1}"/><w:basedOn w:val="Normal"/>` +
      `<w:pPr><w:keepNext/><w:outlineLvl w:val="${i}"/></w:pPr><w:rPr><w:b/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr></w:style>`,
  ).join('');
}

const STYLES =
  `${XML}<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>` +
  '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/>' +
  '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  `${headingStyles()}</w:styles>`;

export function escapeXml(text: string): string {
  return text
    // eslint-disable-next-line no-control-regex -- stripping them is the point
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function runXml(run: Run): string {
  if (run.br) return '<w:r><w:br/></w:r>';
  if (run.tab) return '<w:r><w:tab/></w:r>';
  if (run.text === '') return '';
  const props =
    (run.bold ? '<w:b/><w:bCs/>' : '') +
    (run.italic ? '<w:i/><w:iCs/>' : '') +
    (run.strike ? '<w:strike/>' : '') +
    (run.underline ? '<w:u w:val="single"/>' : '') +
    (run.sup ? '<w:vertAlign w:val="superscript"/>' : run.sub ? '<w:vertAlign w:val="subscript"/>' : '');
  // Tabs inside text become real tab runs rather than literal U+0009 in w:t.
  const parts = run.text.split('\t');
  return parts
    .map((part, i) => {
      const t = part === '' ? '' : `<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ''}<w:t xml:space="preserve">${escapeXml(part)}</w:t></w:r>`;
      return (i > 0 ? '<w:r><w:tab/></w:r>' : '') + t;
    })
    .join('');
}

export function paraXml(p: Para): string {
  const ppr =
    (p.heading !== undefined ? `<w:pStyle w:val="Heading${Math.min(Math.max(p.heading, 1), 6)}"/>` : '') +
    (p.indent ? `<w:ind w:left="${Math.round(p.indent)}"/>` : '') +
    (p.align ? `<w:jc w:val="${p.align}"/>` : '');
  const runs = p.runs.map(runXml).join('');
  if (ppr === '' && runs === '') return '<w:p/>';
  return `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${runs}</w:p>`;
}

function tableXml(t: Table): string {
  const cols = Math.max(...t.rows.map((r) => r.length));
  const base = Math.floor(CONTENT_WIDTH_TWIPS / cols);
  const widths = Array<number>(cols).fill(base);
  widths[cols - 1] = CONTENT_WIDTH_TWIPS - base * (cols - 1);
  const edge = (n: string) => `<w:${n} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`;
  const borders = `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(edge).join('')}</w:tblBorders>`;
  const rows = t.rows
    .map((row) => {
      const cells = Array.from({ length: cols }, (_, i) => {
        const cell = row[i] ?? [{ kind: 'p' as const, runs: [] }];
        return `<w:tc><w:tcPr><w:tcW w:w="${widths[i]}" w:type="dxa"/></w:tcPr>${cell.map(paraXml).join('')}</w:tc>`;
      });
      return `<w:tr>${cells.join('')}</w:tr>`;
    })
    .join('');
  return (
    `<w:tbl><w:tblPr><w:tblW w:w="${CONTENT_WIDTH_TWIPS}" w:type="dxa"/>${borders}</w:tblPr>` +
    `<w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${rows}</w:tbl>`
  );
}

export function blocksToDocx(blocks: readonly Block[]): Uint8Array {
  const body = blocks.map((b) => (b.kind === 'p' ? paraXml(b) : tableXml(b))).join('');
  const document =
    `${XML}<w:document xmlns:w="${W_NS}"><w:body>${body}` +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>' +
    '</w:body></w:document>';

  const zip = new PizZip();
  zip.file('[Content_Types].xml', CONTENT_TYPES);
  zip.file('_rels/.rels', PACKAGE_RELS);
  zip.file('word/_rels/document.xml.rels', DOCUMENT_RELS);
  zip.file('word/document.xml', document);
  zip.file('word/styles.xml', STYLES);
  zip.file('word/settings.xml', SETTINGS);
  return zip.generate({ type: 'uint8array', compression: 'DEFLATE' });
}
