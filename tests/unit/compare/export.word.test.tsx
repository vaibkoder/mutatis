/**
 * RenderCore Phase 2, C1 — Word-vs-Word rendered-redline snapshot → vector PDF.
 *
 * DocxCompareBody shows the "Export rendered redline (snapshot)" button on done;
 * clicking it renders the surfaced render source (the FULL change set — ADR-0008)
 * to a programmatic vector PDF via RenderCore and downloads it. The button is
 * gated on a surfaced render source. The host-doc tracked-changes redline (the
 * PRIMARY artefact) is never touched — no materialise / driver action fires.
 *
 * The adapter is mocked so the click never spins up a real Web Worker under
 * vitest (the real worker path is covered by the adapter/preview tests).
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { render } from 'preact';

import { ComparePanel } from '../../../src/features/compare/ComparePanel.js';
import { makeStubDriverState, type RenderSource } from '../../../src/features/compare/ui-types.js';
import { renderRedlinePdf, downloadPdf } from '../../../src/features/compare/render/index.js';
import { PizZipBundle } from '../../../src/shared/ooxml/index.js';
import type { ChangeClass, ChangeSummary, DiffOp, ParaRecord } from '../../../src/features/compare/model.js';

// Mock the RenderCore adapter — the export click must call `renderRedlinePdf`
// (off-main-thread render) + `downloadPdf`, never spin up a real Web Worker.
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

const wordOps: DiffOp[] = [
  op('e0', 'equal', 'substantiveText'),
  op('r1', 'replace', 'obligationLanguage', {
    origRange: { start: 1, end: 1 }, revRange: { start: 1, end: 1 },
    wordDiffs: [{ op: 'del', text: 'may' }, { op: 'ins', text: 'shall' }],
  }),
  op('f2', 'replace', 'formatting', {
    origRange: { start: 2, end: 2 }, revRange: { start: 2, end: 2 },
    wordDiffs: [{ op: 'del', text: 'colour' }, { op: 'ins', text: 'color' }],
  }),
];
const originalParas = [para(0, 'Recitals'), para(1, 'may'), para(2, 'colour')];
const revisedParas = [para(0, 'Recitals'), para(1, 'shall'), para(2, 'color')];

describe('C1 — DocxCompareBody snapshot → vector-PDF download', () => {
  let container: HTMLDivElement;
  beforeEach(() => {
    container = document.createElement('div');
    document.body.appendChild(container);
    vi.mocked(renderRedlinePdf).mockClear();
    vi.mocked(downloadPdf).mockClear();
  });

  const exportBtn = () =>
    container.querySelector('.dp-compare__export-btn') as HTMLButtonElement | null;

  function renderDocx(driver = makeStubDriverState()) {
    // 'docx' is the default mode — no tab switch needed.
    render(<ComparePanel driver={driver} />, container);
    return driver;
  }

  function doneDriver() {
    const d = makeStubDriverState();
    d.status.value = 'done';
    d.summary.value = summaryWith(2);
    d.ops.value = wordOps;
    d.originalParas.value = originalParas;
    d.revisedParas.value = revisedParas;
    return d;
  }

  it('hides the Export button until a run is done', () => {
    renderDocx(makeStubDriverState());
    expect(exportBtn()).toBeNull();
  });

  it('shows the redline PDF download button on done', () => {
    const d = doneDriver();
    d.renderSource.value = SRC;
    renderDocx(d);
    expect(exportBtn()).not.toBeNull();
    expect(exportBtn()!.textContent).toContain('Download redline (PDF)');
  });

  it('disables the export button until a render source is surfaced', () => {
    const d = doneDriver(); // renderSource still null
    renderDocx(d);
    expect(exportBtn()!.disabled).toBe(true);
  });

  it('renders the FULL set to a vector PDF + downloads it, without any host-doc action', async () => {
    const d = doneDriver();
    d.renderSource.value = SRC;
    // Guard: the host-doc redline path (runCompare/goToChange) must NOT be hit by export.
    const runSpy = vi.spyOn(d, 'runCompare');
    const navSpy = vi.spyOn(d, 'goToChange');
    renderDocx(d);

    exportBtn()!.click();
    await vi.waitFor(() => expect(downloadPdf).toHaveBeenCalledTimes(1));

    // Rendered from the surfaced render source (the full change set; ADR-0008).
    expect(renderRedlinePdf).toHaveBeenCalledTimes(1);
    expect(vi.mocked(renderRedlinePdf).mock.calls[0]![0]).toBe(SRC);
    // Downloaded as a .pdf Blob artefact (the bytes from the render).
    const [bytes, filename] = vi.mocked(downloadPdf).mock.calls[0]!;
    expect(bytes).toBeInstanceOf(Uint8Array);
    expect(filename).toMatch(/\.pdf$/);
    // The editable host-doc redline is untouched — no driver/host action fired.
    expect(runSpy).not.toHaveBeenCalled();
    expect(navSpy).not.toHaveBeenCalled();
  });
});
