/**
 * Compare at scale — the batch index (ADR-0037; spec §5.2; acceptance B5).
 *
 * `export/report-model.ts` established the shape: a flat, pure, date-injected
 * projection that takes NO filter flags, so a count is always the true count.
 * The batch index is the same wall one level up, plus one of its own — a pair
 * the batch was asked to run always yields a row, whatever became of it. A
 * batch index that omitted its failures would look complete and be wrong,
 * which is the batch-scale version of hiding a change.
 */
import { describe, it, expect } from 'vitest';

import { buildBatchIndex, tallyBatch } from '@features/compare/batch/index-model.js';
import { emptyClassificationTally } from '@features/compare/batch/model.js';
import type { BatchRow, BatchRowStatus } from '@features/compare/batch/model.js';

const AT = '2026-09-10T09:00:00.000Z';
const FLATTEN = { template: 'acceptAll', variant: 'acceptAll' } as const;

function row(
  id: string,
  status: BatchRowStatus,
  changeCount = 0,
  extra: Partial<BatchRow> = {},
): BatchRow {
  return {
    id,
    label: `${id}.docx`,
    status,
    changeCount,
    byClassification: emptyClassificationTally(),
    quarantinedOriginal: 0,
    quarantinedRevised: 0,
    artefacts: [],
    artefactOnDemand: false,
    durationMs: 1,
    ...extra,
  };
}

describe('buildBatchIndex', () => {
  it('keeps one row per pair the batch was asked to run', () => {
    const rows = [
      row('a', 'done', 12),
      row('b', 'failed', 0, { reason: 'not a valid .docx package' }),
      row('c', 'cancelled'),
      row('d', 'unrunnable', 0, { reason: 'cross-format comparison is out of scope' }),
    ];

    const index = buildBatchIndex({ mode: 'one-to-many', rows, generatedAt: AT, flatten: FLATTEN });

    expect(index.rows).toHaveLength(4);
    expect(index.totals.pairs).toBe(4);
    // The three non-`done` outcomes are present WITH their reasons, not absent.
    expect(index.rows.map(r => r.status)).toEqual(['done', 'failed', 'cancelled', 'unrunnable']);
    expect(index.rows[1]?.reason).toMatch(/valid \.docx/);
    expect(index.rows[3]?.reason).toMatch(/out of scope/);
  });

  it('counts the changes the on-screen filters hide, exactly like the unfiltered ones', () => {
    // The behavioural form of the ADR-0008 wall, mirroring
    // `export.report-model.test.ts`: feed in changes of the two classes the
    // display filters CAN hide — `formatting` (hideFormattingOnly) and
    // `crossRefRenumber` (collapseRenumbering) — and assert they reach the
    // index at full weight. A filter leaking into this module would show up
    // here as a number that is too small, which is the failure that matters.
    const hidden = row('formatting-only', 'done', 40, {
      byClassification: { ...emptyClassificationTally(), formatting: 30, crossRefRenumber: 10 },
    });
    const plain = row('substantive', 'done', 2, {
      byClassification: { ...emptyClassificationTally(), substantiveText: 2 },
    });

    const index = buildBatchIndex({
      mode: 'one-to-many', rows: [hidden, plain], generatedAt: AT, flatten: FLATTEN,
    });

    expect(index.totals.changes).toBe(42);
    expect(index.rows).toHaveLength(2);
    // And the row's own tally still sums to its own count — no class dropped.
    const tallied = Object.values(index.rows[0]?.byClassification ?? {}).reduce((n, v) => n + v, 0);
    expect(tallied).toBe(index.rows[0]?.changeCount);
  });

  it('states the flatten choice the batch ran under', () => {
    // Spec §8 case 7: per batch, not per pair, and stated ON the index —
    // "accept all tracked changes first" changes what every redline says.
    const index = buildBatchIndex({
      mode: 'one-to-many',
      rows: [row('a', 'done', 1)],
      generatedAt: AT,
      flatten: { template: 'acceptAll', variant: 'rejectAll' },
    });
    expect(index.flatten).toEqual({ template: 'acceptAll', variant: 'rejectAll' });
  });

  it('injects the date rather than reading the clock', () => {
    const a = buildBatchIndex({ mode: 'one-to-many', rows: [], generatedAt: AT, flatten: FLATTEN });
    const b = buildBatchIndex({ mode: 'one-to-many', rows: [], generatedAt: AT, flatten: FLATTEN });
    expect(a.generatedAt).toBe(AT);
    expect(a).toEqual(b);
  });

  it('carries the template label only for one-to-many', () => {
    const withTemplate = buildBatchIndex({
      mode: 'one-to-many', rows: [], generatedAt: AT, flatten: FLATTEN, templateLabel: 'house-nda.docx',
    });
    const without = buildBatchIndex({ mode: 'many-to-many', rows: [], generatedAt: AT, flatten: FLATTEN });

    expect(withTemplate.templateLabel).toBe('house-nda.docx');
    expect('templateLabel' in without).toBe(false);
  });

  it('defaults disclosures to empty rather than undefined', () => {
    expect(buildBatchIndex({ mode: 'one-to-many', rows: [], generatedAt: AT, flatten: FLATTEN }).disclosures)
      .toEqual([]);
  });

  it('carries disclosures through verbatim', () => {
    const disclosures = ['Artefacts for 12 variants are produced on download.'];
    const index = buildBatchIndex({ mode: 'one-to-many', rows: [], generatedAt: AT, flatten: FLATTEN, disclosures });
    expect(index.disclosures).toEqual(disclosures);
  });
});

describe('tallyBatch', () => {
  it('derives every total from the rows, so it cannot disagree with them', () => {
    const totals = tallyBatch([
      row('a', 'done', 5),
      row('b', 'done', 7),
      row('c', 'failed'),
      row('d', 'cancelled'),
      row('e', 'cancelled'),
      row('f', 'unrunnable'),
    ]);

    expect(totals).toEqual({
      pairs: 6, done: 2, failed: 1, cancelled: 2, unrunnable: 1, changes: 12,
    });
  });

  it('counts in-flight rows in `pairs` and nowhere else', () => {
    const totals = tallyBatch([row('a', 'queued'), row('b', 'running'), row('c', 'done', 3)]);

    expect(totals.pairs).toBe(3);
    expect(totals.done).toBe(1);
    expect(totals.changes).toBe(3);
    expect(totals.failed + totals.cancelled + totals.unrunnable).toBe(0);
  });

  it('counts only completed pairs toward changes', () => {
    // A failed pair carries no trustworthy count; it must not inflate the total.
    expect(tallyBatch([row('a', 'failed', 99)]).changes).toBe(0);
  });

  it('handles an empty batch', () => {
    expect(tallyBatch([])).toEqual({
      pairs: 0, done: 0, failed: 0, cancelled: 0, unrunnable: 0, changes: 0,
    });
  });
});
