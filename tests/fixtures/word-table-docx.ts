/**
 * Synthetic Word tables for the Word-vs-Word table tests.
 *
 * Builds a docx bundle whose body is assembled from small helpers: paragraphs,
 * tables with real `tblPr`/`tblGrid`/`trPr`/`tcPr`, merged cells, multi-paragraph
 * cells and nested tables. Every string is invented — no real document content.
 */
import type { PizZipBundle } from '@shared/ooxml/bundle.js';
import { synthesiseDocx } from '@features/compare/io/text/synthesise-docx.js';

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

/** A whole docx bundle whose `word/document.xml` body is `bodyXml`. */
export function tableBundle(bodyXml: string): PizZipBundle {
  const bundle = synthesiseDocx('placeholder');
  bundle.setPart('word/document.xml', `<w:document ${NS}><w:body>${bodyXml}</w:body></w:document>`);
  return bundle;
}

const esc = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** One run of plain text. */
export const run = (text: string): string => `<w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;

/** A paragraph of plain text (optionally with a style). */
export function para(text: string, styleId?: string): string {
  const pPr = styleId === undefined ? '' : `<w:pPr><w:pStyle w:val="${styleId}"/></w:pPr>`;
  return `<w:p>${pPr}${text === '' ? '' : run(text)}</w:p>`;
}

/** A cell: plain text, several paragraphs, or raw block XML, with optional merge properties. */
export interface CellSpec {
  /** Paragraph texts, or raw block XML when `xml` is set. */
  paras?: readonly string[];
  /** Raw block content (paragraphs and/or a nested `<w:tbl>`). Wins over `paras`. */
  xml?: string;
  gridSpan?: number;
  vMerge?: 'restart' | 'continue';
  /** Cell width in twips (default 2000 × gridSpan). */
  width?: number;
  /** Extra `tcPr` children, e.g. shading — carried verbatim. */
  extraTcPr?: string;
}

export type Cell = string | CellSpec;

function cellXml(cell: Cell): string {
  const spec: CellSpec = typeof cell === 'string' ? { paras: [cell] } : cell;
  const span = spec.gridSpan ?? 1;
  const width = spec.width ?? 2000 * span;
  let tcPr = `<w:tcW w:w="${width}" w:type="dxa"/>`;
  if (span > 1) tcPr += `<w:gridSpan w:val="${span}"/>`;
  if (spec.vMerge === 'restart') tcPr += '<w:vMerge w:val="restart"/>';
  if (spec.vMerge === 'continue') tcPr += '<w:vMerge/>';
  tcPr += spec.extraTcPr ?? '';
  const content = spec.xml ?? (spec.paras ?? ['']).map((t) => para(t)).join('');
  return `<w:tc><w:tcPr>${tcPr}</w:tcPr>${content}</w:tc>`;
}

/** Options for a whole table. */
export interface TableOptions {
  /** Grid column count (default: the widest row's span sum). */
  gridCols?: number;
  /** Extra `trPr` children for every row (e.g. `<w:cantSplit/>`). */
  trPr?: string;
  /** A table style id, carried in `tblPr`. */
  styleId?: string;
}

/** A `<w:tbl>` from rows of cells. */
export function table(rows: readonly (readonly Cell[])[], opts: TableOptions = {}): string {
  const spanOf = (c: Cell): number => (typeof c === 'string' ? 1 : (c.gridSpan ?? 1));
  const widest = Math.max(1, ...rows.map((r) => r.reduce((n, c) => n + spanOf(c), 0)));
  const cols = opts.gridCols ?? widest;
  const style = opts.styleId === undefined ? '' : `<w:tblStyle w:val="${opts.styleId}"/>`;
  const tblPr =
    `<w:tblPr>${style}<w:tblW w:w="0" w:type="auto"/>` +
    '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>';
  const grid = `<w:tblGrid>${'<w:gridCol w:w="2000"/>'.repeat(cols)}</w:tblGrid>`;
  const trPr = opts.trPr === undefined ? '' : `<w:trPr>${opts.trPr}</w:trPr>`;
  const body = rows.map((r) => `<w:tr>${trPr}${r.map(cellXml).join('')}</w:tr>`).join('');
  return `<w:tbl>${tblPr}${grid}${body}</w:tbl>`;
}

/** Count of a tag's opening elements in some XML (self-closing or not). */
export function countTag(xml: string, tag: string): number {
  return (xml.match(new RegExp(`<${tag}[\\s/>]`, 'g')) ?? []).length;
}

/** A fee schedule used across the WT cases — invented values. */
export const FEE_ROWS: readonly (readonly string[])[] = [
  ['Item', 'Amount', 'Due'],
  ['Arrangement fee', 'S$12,000', '5 Business Days'],
  ['Agency fee', 'S$4,000', 'annually'],
  ['Commitment fee', '0.35%', 'quarterly'],
];
