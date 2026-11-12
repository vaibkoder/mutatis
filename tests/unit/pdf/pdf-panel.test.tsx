/**
 * D2 — PDF compare panel: deterministic copy + state assertions against a stub
 * driver. The safety-critical disclosures (candidate-differences framing, the
 * "what is NOT compared" list, the per-run image-presence notice, the scanned
 * unsupported message) must render — these are the guardrails, not polish.
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { render } from 'preact';

import { ComparePanel } from '../../../src/features/compare/ComparePanel.js';
import { makeStubPdfDriverState } from '../../../src/features/compare/ui-types.js';
import type { ChangeSummary } from '../../../src/features/compare/model.js';

function emptySummary(total = 0): ChangeSummary {
  return {
    total,
    byCategory: {
      structural: 0, obligationLanguage: 0, numericalValue: 0, definedTerm: 0,
      crossRefRenumber: 0, substantiveText: 0, whitespacePunct: 0, formatting: 0,
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
}

describe('D2 — PdfComparePanel', () => {
  let container: HTMLDivElement;
  beforeEach(() => {
    container = document.createElement('div');
    document.body.appendChild(container);
  });

  function openPdfMode(pdfDriver = makeStubPdfDriverState()) {
    render(<ComparePanel pdfDriver={pdfDriver} />, container);
    const tabs = [...container.querySelectorAll('[role="tab"]')] as HTMLButtonElement[];
    const pdfTab = tabs.find(t => t.textContent === 'PDFs')!;
    pdfTab.click();
    render(<ComparePanel pdfDriver={pdfDriver} />, container);
    return pdfDriver;
  }

  it('exposes a PDFs mode tab', () => {
    render(<ComparePanel pdfDriver={makeStubPdfDriverState()} />, container);
    const labels = [...container.querySelectorAll('[role="tab"]')].map(t => t.textContent);
    expect(labels).toContain('PDFs');
  });

  it('frames results as candidate differences (not a redline)', () => {
    openPdfMode();
    expect(container.textContent).toContain('candidate differences');
    expect(container.textContent?.toLowerCase()).toContain('not a redline');
  });

  it('always shows the "what is NOT compared" disclosure', () => {
    openPdfMode();
    expect(container.textContent).toContain('What is not compared');
    expect(container.textContent).toContain('signatures');
    expect(container.textContent?.toLowerCase()).toContain('scanned');
    expect(container.textContent?.toLowerCase()).toContain('visual formatting');
  });

  it('surfaces the scanned/unsupported error message (P7)', () => {
    const d = makeStubPdfDriverState();
    d.status.value = 'error';
    d.errorMessage.value = 'No extractable text — cannot compare (appears scanned/image-only).';
    openPdfMode(d);
    const alert = container.querySelector('[role="alert"]');
    expect(alert?.textContent).toContain('cannot compare');
  });

  it('renders the per-run image-presence notice when pages carry image content', () => {
    const d = makeStubPdfDriverState();
    d.status.value = 'done';
    d.summary.value = emptySummary(0);
    d.ops.value = [];
    d.imageNotice.value = {
      original: { pagesWithContentImages: [], contentImageCount: 0, boilerplateImageCount: 0, imageDominantPages: [] },
      revised: { pagesWithContentImages: [3], contentImageCount: 1, boilerplateImageCount: 0, imageDominantPages: [] },
    };
    openPdfMode(d);
    expect(container.textContent).toContain('not text-compared');
    expect(container.textContent).toContain('page 3');
  });

  it('omits the image notice when no page carries content images', () => {
    const d = makeStubPdfDriverState();
    d.status.value = 'done';
    d.summary.value = emptySummary(0);
    d.ops.value = [];
    d.imageNotice.value = {
      original: { pagesWithContentImages: [], contentImageCount: 0, boilerplateImageCount: 1, imageDominantPages: [] },
      revised: { pagesWithContentImages: [], contentImageCount: 0, boilerplateImageCount: 1, imageDominantPages: [] },
    };
    openPdfMode(d);
    expect(container.querySelector('.dp-compare__pdf-image-notice')).toBeNull();
  });

  // ── Upload-slot regression coverage (the 3 bugs that shipped non-functional) ──

  it('renders one file <input> per slot with NO for/id double-association (WebKit picker regression)', () => {
    openPdfMode();
    const slots = [...container.querySelectorAll('.dp-compare__pdf-slot')];
    expect(slots).toHaveLength(2);
    for (const slot of slots) {
      const inputs = slot.querySelectorAll('input[type="file"]');
      expect(inputs).toHaveLength(1);
      const input = inputs[0];
      // The input must be nested inside its <label> with NO id, and the label
      // must carry NO `for` — a double-association double-fires the picker in
      // WebKit and the native dialog never opens.
      expect(input.getAttribute('id')).toBeNull();
      expect(input.closest('label')).not.toBeNull();
      expect(slot.querySelector('label[for]')).toBeNull();
      expect(input.getAttribute('accept')).toContain('pdf');
    }
  });

  it('shows a visible "Choose PDF…" affordance, switching to "different" once selected', () => {
    const d = makeStubPdfDriverState();
    d.originalReady.value = true; // first slot is selected, second is not
    openPdfMode(d);
    const labels = [...container.querySelectorAll('.dp-compare__pdf-slot .dp-compare__upload')];
    expect(labels).toHaveLength(2);
    expect(labels[0].textContent).toContain('Choose a different PDF');
    expect(labels[1].textContent).toContain('Choose PDF');
  });

  it('invokes driver.uploadPdf with the picked file (Shell-injected driver path)', () => {
    const d = makeStubPdfDriverState();
    const calls: Array<{ side: string; name: string }> = [];
    d.uploadPdf = (side, file) => calls.push({ side, name: file.name });
    openPdfMode(d);

    const input = container.querySelector(
      '.dp-compare__pdf-slot--original input[type="file"]',
    ) as HTMLInputElement;
    const file = new File(['%PDF-1.4'], 'base.pdf', { type: 'application/pdf' });
    Object.defineProperty(input, 'files', { value: [file], configurable: true });
    input.dispatchEvent(new Event('change', { bubbles: true }));

    expect(calls).toEqual([{ side: 'original', name: 'base.pdf' }]);
  });

  const runBtn = () => container.querySelector('.dp-compare__run-btn') as HTMLButtonElement;

  it('keeps "Compare PDFs" disabled until BOTH sides are ready', () => {
    openPdfMode(makeStubPdfDriverState()); // neither ready
    expect(runBtn().disabled).toBe(true);

    const oneSide = makeStubPdfDriverState();
    oneSide.originalReady.value = true; // only original
    openPdfMode(oneSide);
    expect(runBtn().disabled).toBe(true);
  });

  it('enables "Compare PDFs" once both sides are ready', () => {
    const both = makeStubPdfDriverState();
    both.originalReady.value = true;
    both.revisedReady.value = true;
    openPdfMode(both);
    expect(runBtn().disabled).toBe(false);
  });
});
