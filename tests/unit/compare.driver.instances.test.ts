/**
 * Compare driver — instantiability (ADR-0037, acceptance B1).
 *
 * The driver used to hold every signal, and its three source/abort mutables, at
 * module scope. That is a hard singleton: one Compare run at a time, and any
 * second consumer silently scribbling on the first one's state. It is the
 * prerequisite that blocked batch redlining, where N pairs run concurrently and
 * a drill-down opens a pair of its own.
 *
 * The refactor is behaviour-neutral — the process-wide `compareDriverState` is
 * now simply *one instance* of the factory, and the per-signal exports alias it —
 * so the rest of the suite would not notice a regression here. These are the
 * tests that would.
 */
import { describe, it, expect } from 'vitest';

import {
  createCompareDriver,
  createTextCompareDriver,
  createPdfCompareDriver,
  compareDriverState,
  compareStatus,
  compareOps,
  originalReady,
  textCompareDriverState,
  textStatus,
  textOriginal,
  pdfCompareDriverState,
  pdfStatus,
} from '@features/compare/driver.js';

describe('createCompareDriver — instances are independent', () => {
  it('gives each instance its own signals', () => {
    const a = createCompareDriver();
    const b = createCompareDriver();

    expect(a.status).not.toBe(b.status);

    a.status.value = 'running';
    a.progress.value = 42;
    a.hideFormattingOnly.value = false;

    expect(b.status.value).toBe('idle');
    expect(b.progress.value).toBe(0);
    expect(b.hideFormattingOnly.value).toBe(true);
  });

  it('gives each instance its own source state', () => {
    const a = createCompareDriver();
    const b = createCompareDriver();

    a.useCurrentDocument('original');

    expect(a.originalReady.value).toBe(true);
    expect(b.originalReady.value).toBe(false);
  });

  it('scopes the re-entrancy guard per instance, so one run cannot block another', async () => {
    const a = createCompareDriver();
    const b = createCompareDriver();

    // `a` is mid-run. The old module-level guard (`if (compareStatus.value ===
    // 'running') return;`) would have made `b.runCompare()` a silent no-op —
    // exactly the failure that made batch work impossible.
    a.status.value = 'running';

    await b.runCompare();

    // `b` read its OWN status, got past the guard, and reached the real check.
    expect(b.status.value).toBe('error');
    expect(b.errorMessage.value).toMatch(/upload both documents/i);
    // `a` is untouched.
    expect(a.status.value).toBe('running');
    expect(a.errorMessage.value).toBeNull();
  });

  it('does not leak an instance\'s results into another instance', () => {
    const a = createCompareDriver();
    const b = createCompareDriver();

    a.ops.value = [];
    a.originalParas.value = [{ index: 0 } as never];

    expect(b.ops.value).toBeNull();
    expect(b.originalParas.value).toEqual([]);
  });
});

describe('the process-wide drivers are one instance of their factory', () => {
  it('aliases the legacy per-signal exports onto compareDriverState', () => {
    expect(compareStatus).toBe(compareDriverState.status);
    expect(compareOps).toBe(compareDriverState.ops);
    expect(originalReady).toBe(compareDriverState.originalReady);
  });

  it('aliases the text and pdf exports onto their driver states', () => {
    expect(textStatus).toBe(textCompareDriverState.status);
    expect(textOriginal).toBe(textCompareDriverState.originalText);
    expect(pdfStatus).toBe(pdfCompareDriverState.status);
  });

  it('does not share state with a freshly created instance', () => {
    const fresh = createCompareDriver();
    expect(fresh.status).not.toBe(compareDriverState.status);
  });
});

describe('the text and pdf drivers are instantiable too', () => {
  it('isolates text instances', () => {
    const a = createTextCompareDriver();
    const b = createTextCompareDriver();

    a.originalText.value = 'one';
    a.status.value = 'running';

    expect(b.originalText.value).toBe('');
    expect(b.status.value).toBe('idle');
  });

  it('isolates pdf instances', () => {
    const a = createPdfCompareDriver();
    const b = createPdfCompareDriver();

    a.originalReady.value = true;
    a.progress.value = 70;

    expect(b.originalReady.value).toBe(false);
    expect(b.progress.value).toBe(0);
  });

  it('scopes the pdf upload state per instance', () => {
    const a = createPdfCompareDriver();
    const b = createPdfCompareDriver();

    a.uploadPdf('original', new File([new Uint8Array([1, 2])], 'a.pdf'));

    expect(a.originalReady.value).toBe(true);
    expect(b.originalReady.value).toBe(false);
  });
});
