import { describe, it, expect } from 'vitest';
import { buildRedlineContext } from '../src/renderer/src/ai/context.js';
import { buildSystemPrompt, narrationPrompt } from '../src/renderer/src/ai/prompt.js';
import { buildDeviationMatrix } from '../src/renderer/src/ai/deviation-matrix.js';
import { makeDocxBytes, SAMPLE_ORIGINAL, SAMPLE_REVISED } from './helpers/make-docx.js';
import { runComparison } from '../src/renderer/src/compare/run-comparison.js';

/**
 * The context payload is what replaces RAG: the model is handed the engine's own
 * change set instead of retrieving over a corpus. If it silently under-reports,
 * the model answers confidently about a document it only half saw — which is the
 * exact failure this project exists to avoid.
 *
 * These run the REAL pipeline over real .docx bytes, not a fixture of what the
 * pipeline is assumed to produce.
 */

const AT = '2026-01-01T00:00:00.000Z';

async function compareSamples(charBudget?: number) {
  const outcome = await runComparison({
    originalBytes: makeDocxBytes(SAMPLE_ORIGINAL),
    revisedBytes: makeDocxBytes(SAMPLE_REVISED),
    originalName: 'original.docx',
    revisedName: 'revised.docx',
    date: AT,
  });
  return { outcome, charBudget };
}

describe('buildRedlineContext', () => {
  it('carries the engine change set with before/after text, never markup', async () => {
    const { outcome } = await compareSamples();
    const context = outcome.context;

    expect(context.changes.length).toBeGreaterThan(0);
    expect(context.baseline).toBe('original.docx');
    expect(context.revised).toBe('revised.docx');

    const serialised = JSON.stringify(context.changes);
    expect(serialised).not.toContain('<w:');
    expect(serialised).not.toContain('w:ins');
    expect(serialised).not.toContain('~~');
  });

  it('anchors every change to its DiffOp id — the only stable anchor that exists', async () => {
    const { outcome } = await compareSamples();
    const opIds = new Set(outcome.ops.map((op) => op.id));
    for (const change of outcome.context.changes) {
      expect(opIds.has(change.id)).toBe(true);
    }
  });

  it('reports the same total the summary does', async () => {
    const { outcome } = await compareSamples();
    expect(outcome.context.totals.total).toBe(outcome.summary.total);
  });

  it('always discloses blind spots, including the clause-label limitation', async () => {
    const { outcome } = await compareSamples();
    expect(outcome.context.blindSpots.length).toBeGreaterThan(0);
    expect(outcome.context.blindSpots.join(' ')).toMatch(/Schedules, Annexes/);
    expect(outcome.context.blindSpots.join(' ')).toMatch(/No defined-terms glossary/);
  });

  it('is deterministic for a fixed timestamp', async () => {
    const a = await compareSamples();
    const b = await compareSamples();
    expect(JSON.stringify(a.outcome.context)).toBe(JSON.stringify(b.outcome.context));
  });
});

describe('the context budget', () => {
  /**
   * The budget is the one place this payload can under-report. It must SAY it
   * did — a model handed a silently trimmed change set will answer as though it
   * had seen the lot, which is the failure mode the whole product exists to
   * avoid.
   */
  const para = (index: number, text: string) => ({
    index,
    partId: 'document',
    rawXml: '',
    normText: text,
    normTextAggressive: text.toLowerCase(),
    numberingPrefix: `${index + 1}. `,
    hash: `h${index}`,
    styleId: null,
    isTableCell: false,
    isSignatureBlock: false,
    isHeaderFooter: false,
    isFootnote: false,
  });

  const ops = Array.from({ length: 6 }, (_, i) => ({
    id: `op-${i}`,
    type: 'replace' as const,
    origRange: { start: i, end: i },
    revRange: { start: i, end: i },
    partId: 'document',
    classification: 'substantive' as never,
    classes: [] as never[],
    author: 'test',
    date: AT,
  }));

  const paras = Array.from({ length: 6 }, (_, i) => para(i, 'x'.repeat(100)));

  const input = {
    ops,
    summary: {
      total: 6,
      byCategory: { substantive: 6 } as never,
      renumberingCascades: [],
      movedParagraphs: 0,
      spellingVariantCount: 0,
      quarantinedCommentsOriginal: 0,
      quarantinedCommentsRevised: 0,
      partChanges: { headers: 0, footers: 0, footnotes: 0 },
      styleDefinitionNote: null,
      filtersActive: { hideFormattingOnly: false, collapseRenumbering: false },
    },
    originalParas: paras,
    revisedParas: paras,
    baseline: 'a.docx',
    revised: 'b.docx',
    jurisdiction: 'UK' as const,
    quarantinedOriginal: 0,
    quarantinedRevised: 0,
    generated: AT,
  };

  it('keeps everything when the budget is ample', () => {
    const context = buildRedlineContext({ ...input, charBudget: 1_000_000 });
    expect(context.changes).toHaveLength(6);
    expect(context.omitted).toBe(0);
    expect(context.blindSpots.join(' ')).not.toMatch(/omitted to fit/);
  });

  it('omits, COUNTS and DISCLOSES when the budget is exceeded', () => {
    const context = buildRedlineContext({ ...input, charBudget: 500 });
    expect(context.changes.length).toBeLessThan(6);
    expect(context.omitted).toBe(6 - context.changes.length);
    expect(context.blindSpots.join(' ')).toMatch(/omitted to fit the context budget/);
    expect(context.blindSpots.join(' ')).toMatch(/redline beside you is complete/);
  });

  it('still reports the TRUE total even when it carries fewer changes', () => {
    const context = buildRedlineContext({ ...input, charBudget: 500 });
    expect(context.totals.total).toBe(6);
  });

  it('never returns an empty change set just because one change is oversize', () => {
    const context = buildRedlineContext({ ...input, charBudget: 1 });
    expect(context.changes).toHaveLength(1);
    expect(context.omitted).toBe(5);
  });

  it('derives clause labels from the numbering prefix', () => {
    const context = buildRedlineContext({ ...input, charBudget: 1_000_000 });
    expect(context.changes[0]?.clause).toBe('1');
    expect(context.changes[3]?.clause).toBe('4');
  });
});

describe('changes the payload counts but cannot carry', () => {
  /**
   * The engine counts header, footer and footnote changes into `summary.total`
   * but keeps them out of the op list, and drops `property` ops from the change
   * export because their before and after text is identical. Both are real
   * changes the model is NOT handed. Unstated, the totals stop reconciling
   * against the entries — and a model that cannot reconcile them guesses.
   */
  const AT = '2026-01-01T00:00:00.000Z';

  const para = (index: number, text: string) => ({
    index,
    partId: 'document',
    rawXml: '',
    normText: text,
    normTextAggressive: text.toLowerCase(),
    numberingPrefix: `${index + 1}. `,
    hash: `h${index}`,
    styleId: null,
    isTableCell: false,
    isSignatureBlock: false,
    isHeaderFooter: false,
    isFootnote: false,
  });

  const op = (id: string, type: 'replace' | 'property') => ({
    id,
    type,
    origRange: { start: 0, end: 0 },
    revRange: { start: 0, end: 0 },
    partId: 'document',
    classification: 'substantiveText' as const,
    classes: [] as never[],
    author: 'test',
    date: AT,
  });

  const paras = [para(0, 'the same words on both sides')];

  const base = {
    originalParas: paras,
    revisedParas: paras,
    baseline: 'a.docx',
    revised: 'b.docx',
    jurisdiction: 'UK' as const,
    quarantinedOriginal: 0,
    quarantinedRevised: 0,
    generated: AT,
    charBudget: 1_000_000,
  };

  const summary = (over: Partial<Parameters<typeof buildRedlineContext>[0]['summary']> = {}) =>
    ({
      total: 2,
      byCategory: { substantiveText: 2 } as never,
      renumberingCascades: [],
      movedParagraphs: 0,
      spellingVariantCount: 0,
      quarantinedCommentsOriginal: 0,
      quarantinedCommentsRevised: 0,
      partChanges: { headers: 0, footers: 0, footnotes: 0 },
      styleDefinitionNote: null,
      filtersActive: { hideFormattingOnly: false, collapseRenumbering: false },
      ...over,
    }) as never;

  it('discloses property ops, which carry no text delta to export', () => {
    const context = buildRedlineContext({
      ...base,
      ops: [op('op-0', 'replace'), op('op-1', 'property')] as never,
      summary: summary(),
    });

    // The engine's export drops it, so the entry genuinely is not there...
    expect(context.changes.map((c) => c.id)).toEqual(['op-0']);
    // ...which is only safe because the model is told.
    expect(context.blindSpots.join(' ')).toMatch(/1 change\(s\) alter a paragraph's style/);
    expect(context.blindSpots.join(' ')).toMatch(/Do not call those paragraphs unchanged/);
  });

  it('says nothing about property ops when there are none', () => {
    const context = buildRedlineContext({
      ...base,
      ops: [op('op-0', 'replace')] as never,
      summary: summary(),
    });
    expect(context.blindSpots.join(' ')).not.toMatch(/style, numbering or level/);
  });

  it('discloses changes outside the body, and names where they are', () => {
    const context = buildRedlineContext({
      ...base,
      ops: [op('op-0', 'replace')] as never,
      summary: summary({ partChanges: { headers: 2, footers: 0, footnotes: 3 } }),
    });

    const spots = context.blindSpots.join(' ');
    expect(spots).toMatch(/5 change\(s\) sit outside the document body/);
    expect(spots).toMatch(/2 in headers, 3 in footnotes/);
    // Headers are not rendered at all, so that has to be said too.
    expect(spots).toMatch(/not drawn in the redline either/);
  });

  it('does not claim the redline is missing anything when only footnotes changed', () => {
    const context = buildRedlineContext({
      ...base,
      ops: [op('op-0', 'replace')] as never,
      summary: summary({ partChanges: { headers: 0, footers: 0, footnotes: 1 } }),
    });

    const spots = context.blindSpots.join(' ');
    expect(spots).toMatch(/1 change\(s\) sit outside the document body \(1 in footnotes\)/);
    // Footnotes ARE rendered — claiming otherwise would be a false disclosure.
    expect(spots).not.toMatch(/not drawn in the redline either/);
  });

  it('counts what it actually carried, not total minus the budget drop', () => {
    // total 5: one exportable replace, one property op, three outside the body.
    // The old wording said `total - omitted`, which would claim 5 of 5 here.
    const context = buildRedlineContext({
      ...base,
      charBudget: 1,
      ops: [op('op-0', 'replace'), op('op-1', 'replace'), op('op-2', 'property')] as never,
      summary: summary({ total: 5, partChanges: { headers: 0, footers: 0, footnotes: 3 } }),
    });

    const spots = context.blindSpots.join(' ');
    expect(context.changes).toHaveLength(1);
    expect(spots).toMatch(/carries 1 of the 5 changes the engine counted/);
    expect(spots).not.toMatch(/carries 5 of/);
  });

  it('discloses a style-definition change, which belongs to no paragraph', () => {
    const context = buildRedlineContext({
      ...base,
      ops: [op('op-0', 'replace')] as never,
      summary: summary({ styleDefinitionNote: 'Heading 1 changed from 14pt to 16pt' }),
    });
    expect(context.blindSpots.join(' ')).toMatch(/Heading 1 changed from 14pt to 16pt/);
  });

  it('no longer claims formatting-only changes are included as changes', () => {
    const context = buildRedlineContext({
      ...base,
      ops: [op('op-0', 'replace')] as never,
      summary: summary(),
    });
    const spots = context.blindSpots.join(' ');
    expect(spots).not.toMatch(/Formatting-only and whitespace changes are .*included as changes/);
    // Rule 4 still holds, and still says so.
    expect(spots).toMatch(/Nothing has been filtered out of the redline/);
  });
});

describe('buildSystemPrompt', () => {
  it('says plainly that nothing is open when nothing is', () => {
    const prompt = buildSystemPrompt(null, []);
    expect(prompt).toContain('NO COMPARISON IS OPEN');
    expect(prompt).not.toContain('CHANGE SET');
  });

  it('carries the active comparison in full and the others as totals only', async () => {
    const { outcome } = await compareSamples();
    const active = { id: 't1', title: 'Active', context: outcome.context };
    const other = { id: 't2', title: 'Other', context: outcome.context };

    const prompt = buildSystemPrompt(active, [other]);
    expect(prompt).toContain('OPEN COMPARISON: Active');
    expect(prompt).toContain('CHANGE SET');
    expect(prompt).toContain('ALSO OPEN, NOT INCLUDED HERE');
    expect(prompt).toContain('You have their totals only');
    // The other comparison's change text must NOT be in the prompt twice.
    expect(prompt.split('CHANGE SET').length - 1).toBe(1);
  });

  it('tells the model it is beside the redline, not in front of it', () => {
    const prompt = buildSystemPrompt(null, []);
    expect(prompt).toContain('You sit beside it, not in front of it.');
    expect(prompt).toContain('Never claim a change exists');
  });
});

describe('narrationPrompt', () => {
  it('forbids producing numbers and hands over the finished table', async () => {
    const { outcome } = await compareSamples();
    const matrix = buildDeviationMatrix(
      [
        { id: 't1', title: 'One', context: outcome.context },
        { id: 't2', title: 'Two', context: outcome.context },
      ],
      AT,
    );
    const prompt = narrationPrompt(matrix);
    expect(prompt).toContain('THIS TURN RUNS ON RAILS');
    expect(prompt).toContain('Do not produce any number that is not in the table');
    expect(prompt).toContain('THE MATRIX:');
    expect(prompt).toContain('classification\t');
  });
});
