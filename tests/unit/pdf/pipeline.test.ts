/**
 * B2 — PDF pipeline (runPipelinePdf) end-to-end on the committed fixtures.
 * This is where the headline acceptance tests become reachable:
 *   P1 inserted paragraph, P2 visual-only (zero structural ops),
 *   P4 obligation language shall→may, P6 numeric amount change.
 *
 * No redline is authored (PdfPipelineResult has no authoredBodyXml).
 */
import { describe, it, expect } from 'vitest';

import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import type { CompareConfig } from '@shared/config/index.js';

import { runPipelinePdf } from '../../../src/features/compare/pipeline.js';
import { extractPdfParaRecords } from '../../../src/features/compare/io/pdf/parts.js';
import { nodePdfOpener, readFixtureBytes } from '../../fixtures/pdf/node-opener.js';
import { pdfFixture } from '../../fixtures/pdf/manifest.js';

const CONFIG: CompareConfig = getCompareConfig(SG_PROFILE);
const AUTHOR = 'Reviewer';
const DATE = '2026-06-08T00:00:00.000Z';

async function comparePdfs(origName: string, revName: string) {
  const orig = await extractPdfParaRecords(readFixtureBytes(pdfFixture(origName)), nodePdfOpener, CONFIG);
  const rev = await extractPdfParaRecords(readFixtureBytes(pdfFixture(revName)), nodePdfOpener, CONFIG);
  return runPipelinePdf({
    originalParas: orig.paras,
    revisedParas: rev.paras,
    config: CONFIG,
    author: AUTHOR,
    date: DATE,
  });
}

describe('B2 — PDF pipeline acceptance', () => {
  it('P1 — inserted paragraph → exactly one insert, no cascade misalignment', async () => {
    const result = await comparePdfs('single-column.base.pdf', 'single-column.insert.pdf');
    const inserts = result.ops.filter((op) => op.type === 'insert');
    expect(inserts).toHaveLength(1);
    expect(insertedText(result, inserts[0]!)).toContain('Schedule 1');
    // No spurious churn: only the one insert is non-equal.
    expect(result.ops.filter((op) => op.type !== 'equal' && op.type !== 'insert')).toHaveLength(0);
  });

  it('P2 — visual-only change (identical text) → zero structural DiffOps', async () => {
    const result = await comparePdfs('single-column.base.pdf', 'single-column.fontchange.pdf');
    expect(result.summary.total).toBe(0);
    expect(result.ops.every((op) => op.type === 'equal')).toBe(true);
  });

  it('P3 — renumbering cascade → one CascadeAnnotation, members collapsed, none suppressed', async () => {
    const result = await comparePdfs('renumber.base.pdf', 'renumber.revised.pdf');

    // Exactly one cascade: the uniform Clause +1 shift across the four
    // cross-references downstream of the inserted clause.
    expect(result.summary.renumberingCascades).toHaveLength(1);
    const cascade = result.summary.renumberingCascades[0]!;
    expect(cascade.tokenKind).toBe('Clause');
    expect(cascade.shift).toBe(1);
    expect(cascade.affectedCount).toBe(4);

    // ADR-0008: collapse is presentation-only. Every member op is STILL in the
    // op list (tagged collapsedInto), never dropped.
    const crossRefOps = result.ops.filter((op) => op.classification === 'crossRefRenumber');
    expect(crossRefOps).toHaveLength(4);
    expect(crossRefOps.every((op) => op.collapsedInto === cascade.id)).toBe(true);

    // The inserted clause (heading + body) surfaces as structural inserts —
    // the substantive change is not hidden behind the cascade.
    expect(result.ops.filter((op) => op.type === 'insert')).toHaveLength(2);
  });

  it('P4 — shall→may → obligationLanguage, never filtered', async () => {
    const result = await comparePdfs('single-column.base.pdf', 'single-column.shall-may.pdf');
    const changed = result.ops.filter((op) => op.type !== 'equal');
    expect(changed).toHaveLength(1);
    expect(changed[0]!.classes).toContain('obligationLanguage');
  });

  it('P6 — amount change → numericalValue', async () => {
    const result = await comparePdfs('single-column.base.pdf', 'single-column.amount.pdf');
    const changed = result.ops.filter((op) => op.type !== 'equal');
    expect(changed).toHaveLength(1);
    expect(changed[0]!.classes).toContain('numericalValue');
  });

  it('produces no authored redline (PDF has no tracked-changes target)', async () => {
    const result = await comparePdfs('single-column.base.pdf', 'single-column.amount.pdf');
    expect('authoredBodyXml' in result).toBe(false);
  });
});

function insertedText(
  result: { revisedParas: { index: number; normText: string }[] },
  op: { revRange: { start: number; end: number } | null },
): string {
  if (op.revRange === null) return '';
  return result.revisedParas
    .slice(op.revRange.start, op.revRange.end + 1)
    .map((p) => p.normText)
    .join(' ');
}
