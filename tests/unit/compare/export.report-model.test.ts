/**
 * report-model unit tests.
 *
 * The load-bearing assertion is the ADR-0008 wall: `buildCompareReport` reads
 * the FULL `ops` and emits one row per non-equal op — the formatting-only ops
 * the on-screen `applyFilters` WOULD hide are still present in the report.
 * Filters subtract from the display; they never subtract from the export.
 */
import { describe, it, expect } from 'vitest';
import { buildCompareReport } from '../../../src/features/compare/export/report-model.js';
import { applyFilters } from '../../../src/features/compare/components/ChangeList.js';
import { pdfNotComparedLines } from '../../../src/features/compare/export/safety-copy.js';
import type { ChangeClass, ChangeSummary, DiffOp, ParaRecord } from '../../../src/features/compare/model.js';

function makeOp(partial: Partial<DiffOp> & Pick<DiffOp, 'id' | 'type' | 'classification'>): DiffOp {
  return {
    origRange: { start: 0, end: 0 },
    revRange: { start: 0, end: 0 },
    partId: 'document',
    classes: [partial.classification],
    author: 'Mutatis',
    date: '2026-06-09T00:00:00.000Z',
    ...partial,
  };
}

function makePara(index: number, normText: string): ParaRecord {
  return {
    index,
    partId: 'document',
    rawXml: '<w:p/>',
    normText,
    normTextAggressive: normText.toLowerCase(),
    numberingPrefix: '',
    hash: `h${index}`,
    styleId: null,
    isTableCell: false,
    isSignatureBlock: false,
    isHeaderFooter: false,
    isFootnote: false,
  };
}

function summaryWith(total: number): ChangeSummary {
  const byCategory = {
    structural: 0, obligationLanguage: 0, numericalValue: 0, definedTerm: 0,
    crossRefRenumber: 0, substantiveText: 0, whitespacePunct: 0, formatting: 0,
  } as Record<ChangeClass, number>;
  return {
    total,
    byCategory,
    renumberingCascades: [],
    movedParagraphs: 0,
    partChanges: { headers: 0, footers: 0, footnotes: 0 },
    styleDefinitionNote: null,
    spellingVariantCount: 0,
    quarantinedCommentsOriginal: 0,
    quarantinedCommentsRevised: 0,
    filtersActive: { hideFormattingOnly: true, collapseRenumbering: true },
  };
}

// A change set that deliberately mixes the one filterable class (formatting)
// with substantive ones, plus an 'equal' op that must be excluded, plus a pure
// delete/insert that resolves text from ParaRecords. `o4` carries
// `whitespacePunct` as a SECONDARY class, which is the only way it occurs: the
// class can never be an op's primary classification (ADR-0036), so a fixture
// claiming `classification: 'whitespacePunct'` would be asserting a shape the
// classifier cannot produce.
function mixedOps(): DiffOp[] {
  return [
    makeOp({ id: 'o1', type: 'equal', classification: 'substantiveText' }), // excluded
    makeOp({
      id: 'o2', type: 'replace', classification: 'numericalValue',
      wordDiffs: [{ op: 'eq', text: 'pay ' }, { op: 'del', text: '£100' }, { op: 'ins', text: '£200' }],
    }),
    makeOp({ id: 'o3', type: 'delete', classification: 'formatting', revRange: null }),
    makeOp({
      id: 'o4', type: 'insert', origRange: null,
      classification: 'substantiveText', classes: ['substantiveText', 'whitespacePunct'],
    }),
    makeOp({ id: 'o5', type: 'replace', classification: 'obligationLanguage',
      wordDiffs: [{ op: 'del', text: 'may' }, { op: 'ins', text: 'shall' }] }),
  ];
}

const originalParas = [makePara(0, 'Deleted formatting paragraph.')];
const revisedParas = [makePara(0, 'Inserted whitespace paragraph.')];

function build(ops: DiffOp[]) {
  return buildCompareReport({
    mode: 'pdf',
    ops,
    summary: summaryWith(ops.filter(o => o.type !== 'equal').length),
    originalParas,
    revisedParas,
    disclosures: pdfNotComparedLines(),
    generatedAt: '9 June 2026',
  });
}

describe('buildCompareReport — ADR-0008 wall', () => {
  it('emits one row per non-equal op (excludes equal)', () => {
    const ops = mixedOps();
    const report = build(ops);
    const nonEqual = ops.filter(o => o.type !== 'equal');
    expect(report.rows).toHaveLength(nonEqual.length);
    expect(report.changeCount).toBe(nonEqual.length);
  });

  it('changeCount equals summary.total', () => {
    const ops = mixedOps();
    const report = build(ops);
    expect(report.changeCount).toBe(report.rows.length);
    const total = ops.filter(o => o.type !== 'equal').length;
    expect(report.changeCount).toBe(total);
  });

  it('keeps the formatting-only op the display filter hides', () => {
    const ops = mixedOps();
    const displayed = applyFilters(ops, {
      hideFormattingOnly: true, collapseRenumbering: true,
    });
    const report = build(ops);
    // Sanity: the display filter really does subtract, and subtracts exactly o3.
    // Compare like with like — `applyFilters` sees the 'equal' op, `rows` never
    // does, so the raw lengths are not comparable (they used to be compared
    // anyway, and held only because two ops were being filtered, not one).
    expect(displayed.map(o => o.id)).not.toContain('o3');
    const displayedChanges = displayed.filter(o => o.type !== 'equal');
    expect(displayedChanges.length).toBeLessThan(report.rows.length);
    // …and the report keeps it.
    expect(report.rows.some(r => r.classification === 'formatting')).toBe(true);
    // o4's punctuation change is substantive-primary, so no filter ever touched
    // it on screen either — it is in both.
    expect(displayed.map(o => o.id)).toContain('o4');
    expect(report.rows.map(r => r.id)).toContain('o4');
  });

  it('produces identical rows regardless of any filter intent (filters never reach the builder)', () => {
    const ops = mixedOps();
    // The builder takes no filter argument; the same ops in always give the same rows.
    const a = build(ops);
    const b = build(ops);
    expect(a.rows.map(r => r.id)).toEqual(b.rows.map(r => r.id));
    expect(a.rows.map(r => r.id)).toEqual(['o2', 'o3', 'o4', 'o5']);
  });

  it('resolves replace text from wordDiffs and pure ins/del text from ParaRecords', () => {
    const report = build(mixedOps());
    const replace = report.rows.find(r => r.id === 'o2')!;
    expect(replace.original).toBe('pay £100');
    expect(replace.revised).toBe('pay £200');
    const del = report.rows.find(r => r.id === 'o3')!;
    expect(del.original).toBe('Deleted formatting paragraph.');
    expect(del.revised).toBe('');
    const ins = report.rows.find(r => r.id === 'o4')!;
    expect(ins.revised).toBe('Inserted whitespace paragraph.');
    expect(ins.original).toBe('');
  });
});

describe('buildCompareReport — framing + disclosures', () => {
  it('PDF mode → candidate-differences kind + title', () => {
    const report = build(mixedOps());
    expect(report.kind).toBe('candidate-differences');
    expect(report.title).toBe('Candidate differences report');
    expect(report.modeLabel).toContain('PDF');
  });

  it('carries the blind-spot disclosures verbatim and a local-generation note', () => {
    const report = build(mixedOps());
    expect(report.disclosures.length).toBeGreaterThan(0);
    expect(report.disclosures).toEqual(pdfNotComparedLines());
    expect(report.disclosures.some(l => l.toLowerCase().includes('signatures'))).toBe(true);
    expect(report.generatedNote).toContain('9 June 2026');
    expect(report.generatedNote.toLowerCase()).toContain('nothing left your machine');
  });

  it('includes image-notice lines when provided', () => {
    const report = buildCompareReport({
      mode: 'pdf',
      ops: mixedOps(),
      summary: summaryWith(4),
      originalParas,
      revisedParas,
      disclosures: pdfNotComparedLines(),
      imageNotice: ['Revised: 1 page contains image content not text-compared (page 3).'],
      generatedAt: '9 June 2026',
    });
    expect(report.imageNotice).toBeDefined();
    expect(report.imageNotice).toHaveLength(1);
  });
});
