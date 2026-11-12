/**
 * RenderCore Phase 2, C1 — Text-vs-text rendered-redline snapshot → vector PDF.
 *
 * The TextComparePanel "Export rendered redline (snapshot)" button shows on done;
 * clicking it renders the surfaced render source (the FULL change set — ADR-0008)
 * to a programmatic vector PDF via RenderCore and downloads it. The editable
 * "Open redline (.docx)" flow stays the PRIMARY artefact and is untouched.
 *
 * The adapter is mocked so neither the export click nor the in-pane preview spins
 * up a real Web Worker under vitest.
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { render } from 'preact';

import { ComparePanel } from '../../../src/features/compare/ComparePanel.js';
import { makeStubTextDriverState, type RenderSource } from '../../../src/features/compare/ui-types.js';
import { renderRedlinePdf, downloadPdf } from '../../../src/features/compare/render/index.js';
import { PizZipBundle } from '../../../src/shared/ooxml/index.js';
import type { ChangeClass, ChangeSummary, DiffOp, ParaRecord } from '../../../src/features/compare/model.js';

// Mock the RenderCore adapter — the export click renders a vector PDF off the
// main thread and downloads it; the in-pane preview mounts the same source. Both
// must run without a real Web Worker under vitest.
vi.mock('../../../src/features/compare/render/index.js', () => ({
  renderRedlinePdf: vi.fn(async () => new Uint8Array([0x25, 0x50, 0x44, 0x46])), // %PDF
  downloadPdf: vi.fn(),
  redlinePdfFilename: vi.fn(() => 'mutatis-rendered-redline-2026-06-15.pdf'),
  mountRedlinePreview: vi.fn(async () => ({
    root: document.createElement('div'), refresh: () => {}, destroy: () => {},
  })),
}));

const SRC: RenderSource = { authoredBodyXml: '<w:body/>', bundle: new PizZipBundle() };

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

function para(index: number, normText: string, extra: Partial<ParaRecord> = {}): ParaRecord {
  return {
    index, partId: 'document', rawXml: '<w:p/>', normText,
    normTextAggressive: normText.toLowerCase(), numberingPrefix: '', hash: `h${index}`,
    styleId: null, isTableCell: false, isSignatureBlock: false, isHeaderFooter: false,
    isFootnote: false, ...extra,
  };
}

// equal heading + a numbered replace + a pure delete + a formatting op (display-hidden).
const flowOps: DiffOp[] = [
  op('e0', 'equal', 'substantiveText', { origRange: { start: 0, end: 0 }, revRange: { start: 0, end: 0 } }),
  op('r1', 'replace', 'numericalValue', {
    origRange: { start: 1, end: 1 }, revRange: { start: 1, end: 1 },
    wordDiffs: [{ op: 'eq', text: 'pay ' }, { op: 'del', text: '£100' }, { op: 'ins', text: '£200' }],
  }),
  op('d2', 'delete', 'substantiveText', { origRange: { start: 2, end: 2 }, revRange: null }),
  op('f3', 'replace', 'formatting', {
    origRange: { start: 3, end: 3 }, revRange: { start: 2, end: 2 },
    wordDiffs: [{ op: 'del', text: 'colour' }, { op: 'ins', text: 'color' }],
  }),
];

const originalParas = [
  para(0, 'Heading One', { styleId: 'Heading1' }),
  para(1, 'pay £100', { numberingPrefix: '1.8' }),
  para(2, 'A deleted clause.'),
  para(3, 'colour'),
];
const revisedParas = [
  para(0, 'Heading One', { styleId: 'Heading1' }),
  para(1, 'pay £200', { numberingPrefix: '1.8' }),
  para(2, 'color'),
];

describe('C1 — TextComparePanel snapshot → vector-PDF download', () => {
  let container: HTMLDivElement;
  beforeEach(() => {
    container = document.createElement('div');
    document.body.appendChild(container);
    vi.mocked(renderRedlinePdf).mockClear();
    vi.mocked(downloadPdf).mockClear();
  });

  function openTextMode(driver = makeStubTextDriverState()) {
    render(<ComparePanel textDriver={driver} />, container);
    const tabs = [...container.querySelectorAll('[role="tab"]')] as HTMLButtonElement[];
    tabs.find(t => t.textContent === 'Pasted text')!.click();
    render(<ComparePanel textDriver={driver} />, container);
    return driver;
  }

  function doneDriver() {
    const d = makeStubTextDriverState();
    d.status.value = 'done';
    d.summary.value = summaryWith(3);
    d.ops.value = flowOps;
    d.originalParas.value = originalParas;
    d.revisedParas.value = revisedParas;
    d.redlineReady.value = true;
    return d;
  }

  const exportBtn = () =>
    container.querySelector('.dp-compare__export-btn') as HTMLButtonElement | null;

  it('hides the Export button until a run is done', () => {
    openTextMode(makeStubTextDriverState());
    expect(exportBtn()).toBeNull();
  });

  it('shows the snapshot button on done and keeps the editable .docx flow intact', () => {
    const d = doneDriver();
    d.renderSource.value = SRC;
    openTextMode(d);
    expect(exportBtn()).not.toBeNull();
    expect(exportBtn()!.textContent).toContain('rendered redline');
    // The editable .docx ("Open redline") button is untouched and primary.
    const openBtn = container.querySelector('.dp-compare__open-redline-btn');
    expect(openBtn).not.toBeNull();
    expect(openBtn!.textContent).toContain('.docx');
  });

  it('renders the FULL change set to a vector PDF + downloads it, .docx untouched', async () => {
    const d = doneDriver();
    d.renderSource.value = SRC;
    // Guard: the editable-redline path (openRedline) must NOT be hit by export.
    const openSpy = vi.spyOn(d, 'openRedline');
    openTextMode(d);

    exportBtn()!.click();
    await vi.waitFor(() => expect(downloadPdf).toHaveBeenCalledTimes(1));

    expect(renderRedlinePdf).toHaveBeenCalledTimes(1);
    expect(vi.mocked(renderRedlinePdf).mock.calls[0]![0]).toBe(SRC);
    const [bytes, filename] = vi.mocked(downloadPdf).mock.calls[0]!;
    expect(bytes).toBeInstanceOf(Uint8Array);
    expect(filename).toMatch(/\.pdf$/);
    expect(openSpy).not.toHaveBeenCalled();
  });
});
