/**
 * C1 — SideBySideDiff: renders ins/del/eq inline for a replace, nothing hidden.
 * C2 — PdfFindings: summary + navigable list render from PDF ops; a low-confidence
 *       region is BADGED and still visible (ADR-0008).
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { render } from 'preact';

import type { ChangeSummary, DiffOp, ParaRecord, WordDiff } from '../../../src/features/compare/model.js';
import { SideBySideDiff, splitSides } from '../../../src/features/compare/components/SideBySideDiff.js';
import { PdfFindings } from '../../../src/features/compare/components/PdfFindings.js';

function wd(op: WordDiff['op'], text: string): WordDiff {
  return { op, text };
}

function replaceOp(id: string, wordDiffs: WordDiff[], extra: Partial<DiffOp> = {}): DiffOp {
  return {
    id,
    type: 'replace',
    origRange: { start: 0, end: 0 },
    revRange: { start: 0, end: 0 },
    partId: 'document',
    wordDiffs,
    classification: 'substantiveText',
    classes: ['substantiveText'],
    author: 'A',
    date: 'D',
    ...extra,
  };
}

function para(text: string): ParaRecord {
  return {
    index: 0,
    partId: 'document',
    rawXml: '',
    normText: text,
    normTextAggressive: text.toLowerCase(),
    numberingPrefix: '',
    hash: 'h',
    styleId: null,
    isTableCell: false,
    isSignatureBlock: false,
    isHeaderFooter: false,
    isFootnote: false,
  };
}

describe('C1 — splitSides', () => {
  it('routes eq to both, del to original, ins to revised', () => {
    const diffs = [wd('eq', 'The Seller shall deliver within '), wd('del', 'thirty'), wd('ins', 'sixty'), wd('eq', ' days.')];
    const { original, revised } = splitSides(diffs);
    expect(original.map(d => d.text)).toEqual(['The Seller shall deliver within ', 'thirty', ' days.']);
    expect(revised.map(d => d.text)).toEqual(['The Seller shall deliver within ', 'sixty', ' days.']);
  });
});

describe('C1 — SideBySideDiff', () => {
  let container: HTMLDivElement;
  beforeEach(() => {
    container = document.createElement('div');
    document.body.appendChild(container);
  });

  it('renders ins, del and eq inline for a replace — nothing hidden', () => {
    const op = replaceOp('r1', [wd('eq', 'within '), wd('del', 'thirty'), wd('ins', 'sixty'), wd('eq', ' days')]);
    render(<SideBySideDiff op={op} />, container);
    // The deleted token appears on the Original side, the inserted on the Revised.
    expect(container.querySelectorAll('del.dp-compare__tok--del').length).toBe(1);
    expect(container.querySelectorAll('ins.dp-compare__tok--ins').length).toBe(1);
    expect(container.querySelector('del')?.textContent).toBe('thirty');
    expect(container.querySelector('ins')?.textContent).toBe('sixty');
    // Both columns present; the equal text is rendered on both sides.
    expect(container.querySelectorAll('.dp-compare__sbs-col').length).toBe(2);
  });

  it('badges a low-confidence change but still renders it', () => {
    const op = replaceOp('r2', [wd('del', 'A'), wd('ins', 'B')], { lowConfidence: true });
    render(<SideBySideDiff op={op} />, container);
    expect(container.querySelector('.dp-compare__sbs--low-confidence')).not.toBeNull();
    expect(container.querySelector('.dp-compare__sbs-confidence')).not.toBeNull();
    expect(container.querySelector('del')?.textContent).toBe('A');
    expect(container.querySelector('ins')?.textContent).toBe('B');
  });
});

describe('C2 — PdfFindings', () => {
  let container: HTMLDivElement;
  beforeEach(() => {
    container = document.createElement('div');
    document.body.appendChild(container);
  });

  const summary: ChangeSummary = {
    total: 2,
    byCategory: {
      structural: 0,
      obligationLanguage: 0,
      numericalValue: 1,
      definedTerm: 0,
      crossRefRenumber: 0,
      substantiveText: 1,
      whitespacePunct: 0,
      formatting: 0,
    },
    renumberingCascades: [],
    movedParagraphs: 0,
    partChanges: { headers: 0, footers: 0, footnotes: 0 },
    styleDefinitionNote: null,
    spellingVariantCount: 0,
    quarantinedCommentsOriginal: 0,
    quarantinedCommentsRevised: 0,
    filtersActive: { hideFormattingOnly: false, collapseRenumbering: false },
  };

  const filters = { hideFormattingOnly: false, collapseRenumbering: false };

  it('renders the summary total and a navigable list from PDF ops', () => {
    const ops = [
      replaceOp('a', [wd('del', 'thirty'), wd('ins', 'sixty')], { classification: 'numericalValue', classes: ['numericalValue'] }),
      replaceOp('b', [wd('del', 'old'), wd('ins', 'new')], { lowConfidence: true }),
    ];
    render(
      <PdfFindings ops={ops} summary={summary} filters={filters} originalParas={[para('x')]} revisedParas={[para('y')]} />,
      container,
    );
    expect(container.querySelector('.dp-compare__summary-total-count')?.textContent).toBe('2');
    expect(container.querySelectorAll('[role="listitem"]').length).toBe(2);
  });

  it('badges the low-confidence op and keeps it visible (ADR-0008)', () => {
    const ops = [replaceOp('b', [wd('del', 'old'), wd('ins', 'new')], { lowConfidence: true })];
    render(
      <PdfFindings ops={ops} summary={{ ...summary, total: 1 }} filters={filters} originalParas={[para('x')]} revisedParas={[para('y')]} />,
      container,
    );
    const row = container.querySelector('.dp-compare__change-row--low-confidence');
    expect(row).not.toBeNull();
    const badge = container.querySelector('.dp-compare__change-badge');
    expect(badge?.textContent).toBe('low confidence');
    // Visible, not filtered away.
    expect(container.querySelectorAll('[role="listitem"]').length).toBe(1);
  });

  it('shows the side-by-side detail when a row is selected', () => {
    const ops = [replaceOp('a', [wd('del', 'thirty'), wd('ins', 'sixty')])];
    render(
      <PdfFindings ops={ops} summary={{ ...summary, total: 1 }} filters={filters} originalParas={[para('x')]} revisedParas={[para('y')]} />,
      container,
    );
    expect(container.querySelector('.dp-compare__sbs')).toBeNull();
    (container.querySelector('[role="listitem"]') as HTMLElement).click();
    render(
      <PdfFindings ops={ops} summary={{ ...summary, total: 1 }} filters={filters} originalParas={[para('x')]} revisedParas={[para('y')]} />,
      container,
    );
    expect(container.querySelector('.dp-compare__sbs')).not.toBeNull();
    expect(container.querySelector('del')?.textContent).toBe('thirty');
  });
});
