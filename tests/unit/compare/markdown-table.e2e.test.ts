/**
 * ADR-0027 — end-to-end Markdown-table compare.
 * Runs the real text path (extractTextRecords → runComparePipeline) and confirms
 * the authored redline + RenderCore both see a structured table. MT1, MT2, MT3,
 * MT6, MT7 at the integration level.
 */
import { describe, it, expect } from 'vitest';

import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import type { CompareConfig } from '@shared/config/index.js';
import { runComparePipeline } from '@features/compare/pipeline.js';
import { extractTextRecords } from '@features/compare/io/text/extract-text-records.js';
import { buildRenderModel } from '@shared/render-core/input/index.js';

const CONFIG: CompareConfig = getCompareConfig(SG_PROFILE);
const AUTHOR = 'Mutatis';
const DATE = '2026-06-19T00:00:00Z';

function runText(original: string, revised: string) {
  return runComparePipeline({
    originalParas: extractTextRecords(original),
    revisedParas: extractTextRecords(revised),
    config: CONFIG,
    author: AUTHOR,
    date: DATE,
  });
}

const TABLE_A = 'Item | Amount\n--- | ---\nConsideration | S$1,250,000\nDeposit | S$125,000';
const TABLE_B = 'Item | Amount\n--- | ---\nConsideration | S$1,500,000\nDeposit | S$125,000';

describe('Markdown-table compare — end to end', () => {
  it('MT1 — two identical Markdown tables → a table, no changes', async () => {
    const { authoredBodyXml, ops } = await runText(TABLE_A, TABLE_A);
    expect(authoredBodyXml).toContain('<w:tbl>');
    // `\b` excludes the <w:insideH>/<w:insideV> table border elements.
    expect(authoredBodyXml).not.toMatch(/<w:ins\b|<w:del\b|<w:cellIns\b|<w:cellDel\b/);
    expect(ops.filter((o) => o.type !== 'equal')).toHaveLength(0);
  });

  it('MT2 — one cell edited → a cell-level redline inside a real table', async () => {
    const { authoredBodyXml, summary } = await runText(TABLE_A, TABLE_B);
    expect(authoredBodyXml).toContain('<w:tbl>');
    expect(authoredBodyXml).toMatch(/<w:delText[^>]*>S\$1,250,000<\/w:delText>/);
    expect(authoredBodyXml).toMatch(/<w:t[^>]*>S\$1,500,000<\/w:t>/);
    // The unchanged deposit value is emitted as a clean run (no revision wrapper).
    expect(authoredBodyXml).toMatch(/<w:r><w:t[^>]*>S\$125,000<\/w:t><\/w:r>/);
    // The change is categorised (never suppressed): it appears in the summary.
    expect(summary.total).toBeGreaterThanOrEqual(1);
  });

  it('MT2 — the table change is substantive, never a filterable class', async () => {
    const { ops } = await runText(TABLE_A, TABLE_B);
    const changed = ops.find((o) => o.type !== 'equal');
    expect(changed).toBeDefined();
    expect(changed!.classification).not.toBe('formatting');
    expect(changed!.classification).not.toBe('whitespacePunct');
  });

  it('MT3 — an added column reads structurally (cellIns), not a wholesale rewrite', async () => {
    const out = await runText(
      'Item | Amount\n---|---\nFee | 100',
      'Item | Amount | Note\n---|---|---\nFee | 100 | urgent',
    );
    expect(out.authoredBodyXml).toMatch(/<w:cellIns\b/);
    expect(out.authoredBodyXml).toMatch(/<w:t[^>]*>Note<\/w:t>/);
    // The untouched first columns are not struck.
    expect(out.authoredBodyXml).not.toMatch(/<w:delText/);
  });

  it('MT6 — paragraphs around a table each diff via their own path, order preserved', async () => {
    const orig = 'Recitals follow.\nItem | Amount\n---|---\nFee | 100\nSigned by the parties.';
    const rev = 'Recitals follow now.\nItem | Amount\n---|---\nFee | 150\nSigned by the parties.';
    const { authoredBodyXml } = await runText(orig, rev);
    // The table is present and carries the cell change…
    expect(authoredBodyXml).toContain('<w:tbl>');
    expect(authoredBodyXml).toMatch(/<w:delText[^>]*>100<\/w:delText>/);
    expect(authoredBodyXml).toMatch(/<w:t[^>]*>150<\/w:t>/);
    // …the surrounding paragraph edit is also redlined (word-level), and the
    // unchanged signature line survives. Document order: recitals before the
    // table before the signature line.
    expect(authoredBodyXml).toContain('now');
    expect(authoredBodyXml).toContain('Signed by the parties.');
    const tblAt = authoredBodyXml.indexOf('<w:tbl>');
    expect(authoredBodyXml.indexOf('Recitals')).toBeLessThan(tblAt);
    expect(authoredBodyXml.indexOf('Signed by the parties.')).toBeGreaterThan(tblAt);
  });

  it('MT7 — a plain-text paste (no delimiter) produces no table, paragraphs only', async () => {
    const { authoredBodyXml } = await runText('Clause 1.\nClause 2.', 'Clause 1.\nClause 2 amended.');
    expect(authoredBodyXml).not.toContain('<w:tbl>');
    expect(authoredBodyXml).toContain('<w:ins');
  });

  it('RenderCore parses the authored table into a table block (preview/PDF path)', async () => {
    const { authoredBodyXml } = await runText(TABLE_A, TABLE_B);
    const model = buildRenderModel({ authoredBodyXml });
    expect(model.blocks.some((b) => b.kind === 'table')).toBe(true);
  });
});
