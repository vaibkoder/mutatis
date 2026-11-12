/**
 * Clean `<w:tbl>` OOXML builder + scaffolding for the Text-vs-text Markdown path.
 * Spec: spec/features/10-compare-additional-sources.md §4.1 · ADR-0027.
 *
 * Builds a valid, bordered Word table from a parsed grid (rows of cell strings).
 * The clean (no-revision) form is used for the synthesised package body and for a
 * table-unit record's `rawXml` (so an UNCHANGED table is emitted verbatim by the
 * author's `emitEqual`). The lower-level scaffolding helpers (`wrapTable`,
 * `tableRowXml`, `tableCellXml`, `cleanCellParagraph`, `columnWidths`) are shared
 * with the structured table redline author (`output/author-table.ts`) so the two
 * produce byte-compatible structure.
 *
 * Column widths split the A4 content width (page 11906 − 2×1440 margins, matching
 * `synthesise-docx`'s `<w:sectPr>`) equally across columns. Pure: no Office JS,
 * no Date.now(), no Math.random().
 */
import { escapeXmlText } from './synthesise-docx.js';

/** A4 content width in twips: 11906 page − 2×1440 margins (see synthesise-docx sectPr). */
const CONTENT_WIDTH_TWIPS = 9026;

/** Equal column widths (twips) for `ncols` columns; the last absorbs any remainder. */
export function columnWidths(ncols: number): number[] {
  if (ncols <= 0) return [];
  const base = Math.floor(CONTENT_WIDTH_TWIPS / ncols);
  const widths = Array<number>(ncols).fill(base);
  widths[ncols - 1] = CONTENT_WIDTH_TWIPS - base * (ncols - 1);
  return widths;
}

/** `<w:tblGrid>` from column widths. */
export function tblGridXml(widths: number[]): string {
  return `<w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;
}

/** Single-line borders on all six edges so the table renders as a grid in Word + RenderCore. */
function tblBordersXml(): string {
  const edge = (name: string): string =>
    `<w:${name} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`;
  return (
    '<w:tblBorders>' +
    edge('top') +
    edge('left') +
    edge('bottom') +
    edge('right') +
    edge('insideH') +
    edge('insideV') +
    '</w:tblBorders>'
  );
}

/** A single clean run for cell text, or '' when empty (so nothing is wrapped). */
export function cellRunXml(text: string): string {
  if (text === '') return '';
  return `<w:r><w:t xml:space="preserve">${escapeXmlText(text)}</w:t></w:r>`;
}

/** A clean cell paragraph: `<w:p/>` when empty, else one preserved-space run. */
export function cleanCellParagraph(text: string): string {
  if (text === '') return '<w:p/>';
  return `<w:p>${cellRunXml(text)}</w:p>`;
}

/**
 * A `<w:tc>` cell. `innerXml` is one or more block elements (`<w:p>`/`<w:tbl>`).
 * `tcPrExtra` injects extra `<w:tcPr>` children (e.g. a cell-deletion mark is not
 * used — row/run marks carry deletions — but the seam is kept for symmetry).
 */
export function tableCellXml(widthTwips: number, innerXml: string, tcPrExtra = ''): string {
  return (
    '<w:tc>' +
    `<w:tcPr><w:tcW w:w="${widthTwips}" w:type="dxa"/>${tcPrExtra}</w:tcPr>` +
    innerXml +
    '</w:tc>'
  );
}

/**
 * A `<w:tr>` row. `trPrXml` is the full `<w:trPr>…</w:trPr>` (used to carry a
 * row-insertion/deletion mark in the redline author); omit for a clean row.
 */
export function tableRowXml(cellsXml: string, trPrXml = ''): string {
  return `<w:tr>${trPrXml}${cellsXml}</w:tr>`;
}

/** Wrap rows in `<w:tbl>` with table properties + grid. */
export function wrapTable(widths: number[], rowsXml: string): string {
  return (
    '<w:tbl>' +
    `<w:tblPr><w:tblW w:w="0" w:type="auto"/>${tblBordersXml()}</w:tblPr>` +
    tblGridXml(widths) +
    rowsXml +
    '</w:tbl>'
  );
}

/**
 * Build a clean (no-revision) `<w:tbl>` from a padded grid of cell strings.
 * `rows[0]` is the header; every row is assumed already padded to a uniform width
 * (the parser guarantees this).
 */
export function tableXml(rows: string[][]): string {
  const ncols = rows.length === 0 ? 0 : rows[0]!.length;
  const widths = columnWidths(ncols);
  const rowsXml = rows
    .map((cells) =>
      tableRowXml(
        cells.map((c, i) => tableCellXml(widths[i] ?? widths[ncols - 1] ?? 0, cleanCellParagraph(c))).join(''),
      ),
    )
    .join('');
  return wrapTable(widths, rowsXml);
}
