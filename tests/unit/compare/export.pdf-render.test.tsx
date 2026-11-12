/**
 * print-view render + PDF-panel wiring.
 *
 * (b) ExportPrintRoot renders the full CompareReport (title, disclosures, image
 *     notice, and EVERY row with del/ins markup).
 * (c) The PDF panel shows the Export button only when done and, on click, prints
 *     the FULL change set (rows == summary.total) — never the filtered subset.
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { render } from 'preact';

import { ComparePanel } from '../../../src/features/compare/ComparePanel.js';
import { makeStubPdfDriverState } from '../../../src/features/compare/ui-types.js';
import { ExportPrintRoot } from '../../../src/features/compare/export/print-view.js';
import { buildCompareReport } from '../../../src/features/compare/export/report-model.js';
import { pdfNotComparedLines } from '../../../src/features/compare/export/safety-copy.js';
import type { ChangeClass, ChangeSummary, DiffOp, ParaRecord } from '../../../src/features/compare/model.js';

function summaryWith(total: number): ChangeSummary {
  const byCategory = {
    structural: 0, obligationLanguage: 0, numericalValue: 0, definedTerm: 0,
    crossRefRenumber: 0, substantiveText: 0, whitespacePunct: 0, formatting: 0,
  } as Record<ChangeClass, number>;
  return {
    total, byCategory, renumberingCascades: [], movedParagraphs: 0,
    partChanges: { headers: 0, footers: 0, footnotes: 0 },
    styleDefinitionNote: null,
    spellingVariantCount: 0, quarantinedCommentsOriginal: 0, quarantinedCommentsRevised: 0,
    filtersActive: { hideFormattingOnly: true, collapseRenumbering: true },
  };
}

function op(id: string, type: DiffOp['type'], classification: ChangeClass, extra: Partial<DiffOp> = {}): DiffOp {
  return {
    id, type, classification,
    origRange: { start: 0, end: 0 }, revRange: { start: 0, end: 0 },
    partId: 'document', classes: [classification], author: 'Mutatis',
    date: '2026-06-09T00:00:00.000Z', ...extra,
  };
}

const sampleOps: DiffOp[] = [
  op('o1', 'equal', 'substantiveText'),
  op('o2', 'replace', 'numericalValue', {
    wordDiffs: [{ op: 'eq', text: 'pay ' }, { op: 'del', text: '£100' }, { op: 'ins', text: '£200' }],
  }),
  op('o3', 'replace', 'formatting', {
    wordDiffs: [{ op: 'del', text: 'colour' }, { op: 'ins', text: 'color' }],
  }),
];

const para: ParaRecord = {
  index: 0, partId: 'document', rawXml: '<w:p/>', normText: '', normTextAggressive: '',
  numberingPrefix: '', hash: 'h0', styleId: null, isTableCell: false,
  isSignatureBlock: false, isHeaderFooter: false, isFootnote: false,
};

describe('A6(b) — ExportPrintRoot render', () => {
  let container: HTMLDivElement;
  beforeEach(() => {
    container = document.createElement('div');
    document.body.appendChild(container);
  });

  function buildSample() {
    return buildCompareReport({
      mode: 'pdf', ops: sampleOps, summary: summaryWith(2),
      originalParas: [para], revisedParas: [para],
      disclosures: pdfNotComparedLines(),
      imageNotice: ['Revised: 1 page contains image content not text-compared (page 3).'],
      generatedAt: '9 June 2026',
    });
  }

  it('renders the title, mode label and local-generation note', () => {
    render(<ExportPrintRoot report={buildSample()} />, container);
    expect(container.textContent).toContain('Candidate differences report');
    expect(container.textContent).toContain('PDF-vs-PDF comparison');
    expect(container.textContent?.toLowerCase()).toContain('nothing left your machine');
  });

  it('renders the disclosures and the image notice', () => {
    render(<ExportPrintRoot report={buildSample()} />, container);
    expect(container.textContent).toContain('What is not compared');
    expect(container.textContent?.toLowerCase()).toContain('signatures');
    expect(container.textContent).toContain('not text-compared');
  });

  it('renders EVERY non-equal row with del/ins markup', () => {
    render(<ExportPrintRoot report={buildSample()} />, container);
    const rows = container.querySelectorAll('.dp-export-print__row');
    expect(rows).toHaveLength(2); // o2 + o3; o1 (equal) excluded
    // Each side renders its full text under del/ins styling (del+eq / ins+eq).
    expect(container.querySelector('del')?.textContent).toBe('pay £100');
    expect(container.querySelector('ins')?.textContent).toBe('pay £200');
    // The formatting row (display-filtered on screen) is present in the export.
    expect(container.textContent).toContain('FMT');
  });
});

describe('A6(c) — PDF panel Export wiring', () => {
  let container: HTMLDivElement;
  beforeEach(() => {
    container = document.createElement('div');
    document.body.appendChild(container);
    window.print = vi.fn();
    document.getElementById('dp-export-print-host')?.remove();
  });

  function openPdfMode(driver = makeStubPdfDriverState()) {
    render(<ComparePanel pdfDriver={driver} />, container);
    const tabs = [...container.querySelectorAll('[role="tab"]')] as HTMLButtonElement[];
    tabs.find(t => t.textContent === 'PDFs')!.click();
    render(<ComparePanel pdfDriver={driver} />, container);
    return driver;
  }

  const exportBtn = () =>
    container.querySelector('.dp-compare__export-btn') as HTMLButtonElement | null;

  it('hides the Export button until a run is done', () => {
    openPdfMode(makeStubPdfDriverState());
    expect(exportBtn()).toBeNull();
  });

  it('shows the Export button (candidate-differences framing) when done', () => {
    const d = makeStubPdfDriverState();
    d.status.value = 'done';
    d.summary.value = summaryWith(2);
    d.ops.value = sampleOps;
    openPdfMode(d);
    expect(exportBtn()).not.toBeNull();
    expect(exportBtn()!.textContent).toContain('candidate-differences');
    expect(exportBtn()!.textContent?.toLowerCase()).not.toContain('redline');
  });

  it('prints the FULL change set (rows == summary.total), not the filtered list', () => {
    const d = makeStubPdfDriverState();
    d.status.value = 'done';
    d.summary.value = summaryWith(2);
    d.ops.value = sampleOps; // incl. a formatting op the on-screen filter hides
    d.originalParas.value = [para];
    d.revisedParas.value = [para];
    openPdfMode(d);

    exportBtn()!.click();

    expect(window.print).toHaveBeenCalledTimes(1);
    const host = document.getElementById('dp-export-print-host')!;
    const rows = host.querySelectorAll('.dp-export-print__row');
    expect(rows).toHaveLength(2); // both non-equal ops — formatting NOT dropped
  });
});

/**
 * ADR-0042's fifth op type reached the printed report as the literal word
 * "property" — the implementation's name for the change, in a document handed to
 * a lawyer. Every other type already read as English, so nothing caught it.
 */
describe('the printed report names the change, not the op type', () => {
  let container: HTMLDivElement;
  beforeEach(() => {
    container = document.createElement('div');
    document.body.appendChild(container);
  });

  it('prints a property row as formatting / numbering', () => {
    const report = buildCompareReport({
      mode: 'word',
      ops: [
        op('p1', 'property', 'formatting', {
          propertyChange: {
            styleId: true,
            numberingIdentity: true,
            numberLabel: true,
            from: { styleId: 'Clause3', numberLabel: '(vi)', ilvl: 3 },
            to: { styleId: 'Clause2', numberLabel: '(d)', ilvl: 2 },
          },
        }),
      ],
      summary: summaryWith(1),
      originalParas: [para],
      revisedParas: [para],
      disclosures: [],
      generatedAt: '2026-09-12',
    });

    render(<ExportPrintRoot report={report} />, container);
    const text = container.textContent ?? '';

    expect(text).toContain('formatting / numbering');
    // The word this replaces. It is not a term of art in any jurisdiction.
    expect(container.querySelector('.dp-export-print__row-type')!.textContent).not.toBe('property');
  });
});
