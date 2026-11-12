/**
 * Text-vs-text paragraph + Markdown-table record extractor.
 * Spec: spec/features/10-compare-additional-sources.md §4.1 · ADR-0027.
 *
 * Builds the `ParaRecord[]` the Compare pipeline diffs, directly from a pasted
 * blob — paragraphs AND Markdown tables (ADR-0027). A paragraph line becomes a
 * normal record (identical in every comparison field to the plain-text path: same
 * `normText`, `hash`, numbering-prefix and signature-block handling). A Markdown
 * table becomes a single **table-unit** record (`table` set) so the surrounding
 * paragraphs ride the unchanged structural/word/move/classify path while the table
 * is diffed structurally by the table author. Word-vs-Word never reaches this
 * extractor, so its `extractParaRecords` and output stay byte-identical.
 *
 * Pure: no Office JS, no Date.now(), no Math.random().
 */
import type { ParaRecord } from '../../model.js';
import type { RowModel } from '../../diff/tables-core.js';
import { fnv1a } from '../../preprocess/hash.js';
import { normaliseText, normaliseTextAggressive, stripNumberingPrefix } from '../../preprocess/normalise.js';
import { isSignatureBlockText } from '../../preprocess/pararecord.js';
import { paragraphXml } from './synthesise-docx.js';
import { tableXml } from './synthesise-table.js';
import { parseBlocks } from './parse-markdown-table.js';

/** Build the RowModel grid (normalised cell text + content hashes) from a parsed table. */
export function toRowModels(rows: string[][]): RowModel[] {
  return rows.map((cells) => {
    const cellModels = cells.map((c) => {
      const text = normaliseText(c);
      return { text, hash: fnv1a(text) };
    });
    return { cells: cellModels, hash: fnv1a(cellModels.map((c) => c.hash).join('|')) };
  });
}

/** A readable, deterministic flattening of a table grid — drives the structural hash + inline preview. */
function tableNormText(grid: RowModel[]): string {
  return grid.map((r) => r.cells.map((c) => c.text).join(' | ')).join('  ¶  ');
}

/** Extract `ParaRecord[]` from a pasted plain-text-or-Markdown blob. */
export function extractTextRecords(text: string): ParaRecord[] {
  const blocks = parseBlocks(text);
  const records: ParaRecord[] = [];
  let index = 0;

  for (const block of blocks) {
    if (block.kind === 'table') {
      const grid = toRowModels(block.rows);
      const normText = tableNormText(grid);
      records.push({
        index: index++,
        partId: 'document',
        rawXml: tableXml(block.rows),
        normText,
        normTextAggressive: normaliseTextAggressive(normText),
        numberingPrefix: '',
        hash: fnv1a(normText),
        styleId: null,
        isTableCell: false,
        isSignatureBlock: isSignatureBlockText(normText),
        isHeaderFooter: false,
        isFootnote: false,
        table: grid,
      });
      continue;
    }

    // Paragraph line — mirrors preprocess/pararecord.ts field-for-field.
    const line = block.text;
    const normWhole = normaliseText(line);
    const { prefix, body } = stripNumberingPrefix(normWhole);
    records.push({
      index: index++,
      partId: 'document',
      rawXml: paragraphXml(line),
      normText: body,
      normTextAggressive: normaliseTextAggressive(body),
      numberingPrefix: prefix,
      hash: fnv1a(body),
      styleId: null,
      isTableCell: false,
      isSignatureBlock: isSignatureBlockText(normWhole),
      isHeaderFooter: false,
      isFootnote: false,
    });
  }

  return records;
}
