/**
 * ADR-0027 — Markdown-table parser unit tests.
 * Deterministic detection: header + `---|---` delimiter → table; everything else
 * (incl. a lone delimiter with no header) stays a paragraph (MT7).
 */
import { describe, it, expect } from 'vitest';

import { parseBlocks, splitRow, type Block } from '@features/compare/io/text/parse-markdown-table.js';

function tables(blocks: Block[]): Array<string[][]> {
  return blocks.filter((b): b is Extract<Block, { kind: 'table' }> => b.kind === 'table').map((b) => b.rows);
}

describe('splitRow', () => {
  it('splits on pipes and trims, dropping optional outer pipes', () => {
    expect(splitRow('| a | b |')).toEqual(['a', 'b']);
    expect(splitRow('a | b')).toEqual(['a', 'b']);
  });

  it('preserves genuinely empty interior cells', () => {
    expect(splitRow('a || b')).toEqual(['a', '', 'b']);
  });

  it('treats a backslash-escaped pipe as a literal pipe inside the cell', () => {
    expect(splitRow('a \\| b | c')).toEqual(['a | b', 'c']);
  });
});

describe('parseBlocks — detection', () => {
  it('parses a header + delimiter + body rows into one table block', () => {
    const blocks = parseBlocks('Item | Amount\n--- | ---\nFee | 100\nTax | 20');
    expect(blocks).toHaveLength(1);
    expect(tables(blocks)[0]).toEqual([
      ['Item', 'Amount'],
      ['Fee', '100'],
      ['Tax', '20'],
    ]);
  });

  it('tolerates leading/trailing outer pipes and alignment colons', () => {
    const blocks = parseBlocks('| Item | Amount |\n| :--- | ---: |\n| Fee | 100 |');
    expect(tables(blocks)[0]).toEqual([
      ['Item', 'Amount'],
      ['Fee', '100'],
    ]);
  });

  it('accepts a header-only table (delimiter but no body rows)', () => {
    const blocks = parseBlocks('A | B\n---|---');
    expect(tables(blocks)).toHaveLength(1);
    expect(tables(blocks)[0]).toEqual([['A', 'B']]);
  });

  it('MT7: a delimiter line with NO header above it stays plain text', () => {
    const blocks = parseBlocks('intro\n| --- |\noutro');
    expect(tables(blocks)).toHaveLength(0);
    expect(blocks.every((b) => b.kind === 'paragraph')).toBe(true);
    expect(blocks).toHaveLength(3);
  });

  it('MT7: text with no delimiter row stays one paragraph per line', () => {
    const blocks = parseBlocks('Clause 1.\nClause 2.\nClause 3.');
    expect(tables(blocks)).toHaveLength(0);
    expect(blocks).toHaveLength(3);
  });

  it('pads ragged rows to the widest row so no cell is dropped', () => {
    const blocks = parseBlocks('a|b|c\n---|---|---\nx|y');
    expect(tables(blocks)[0]).toEqual([
      ['a', 'b', 'c'],
      ['x', 'y', ''],
    ]);
  });

  it('a blank line ends a table; following pipe text with no delimiter is a paragraph', () => {
    const blocks = parseBlocks('h1|h2\n---|---\na|b\n\nc|d');
    expect(tables(blocks)).toHaveLength(1);
    expect(tables(blocks)[0]).toEqual([
      ['h1', 'h2'],
      ['a', 'b'],
    ]);
    // blank paragraph + the trailing "c|d" (no delimiter → not a table)
    const paras = blocks.filter((b) => b.kind === 'paragraph');
    expect(paras).toHaveLength(2);
  });

  it('keeps paragraphs around a table in document order', () => {
    const blocks = parseBlocks('Intro line.\nName | Role\n---|---\nAlice | Buyer\nOutro line.');
    expect(blocks.map((b) => b.kind)).toEqual(['paragraph', 'table', 'paragraph']);
  });

  it('the `---|---` delimiter line never appears as content (never-suppress structure)', () => {
    const blocks = parseBlocks('A | B\n---|---\nx | y');
    const rows = tables(blocks)[0]!;
    expect(rows.flat()).not.toContain('---');
    expect(rows.flat().join('|')).not.toMatch(/-{3,}/);
  });
});
