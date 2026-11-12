/**
 * standalone web Compare e2e.
 *
 * Proves the web Compare surface works end-to-end through the HostPort seam:
 *   1. WebCompare renders the three mode tabs and the live drivers (not stubs);
 *      the Word-vs-Word slots are upload-both (no "Use current document" on web).
 *   2. Text-vs-text on web yields the SAME ops/summary the add-in pipeline
 *      produces for the same input (the driver routes through the same pipeline).
 *   3. openRedline() on web routes through getActiveHost().writeBack — i.e. the
 *      WebHostPort Blob `.docx` download — and makes NO Office call (no Word.run).
 *
 * The office-free build assertion (web chunk + the shared chunk it loads contain
 * no Word.run/getFileAsync/insertFileFromBase64/Office.context, with Compare in
 * the graph) is the manual `npm run build` grep recorded in the plan's Evidence.
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { render } from 'preact';
import { act } from 'preact/test-utils';

import { setActiveHost, _resetActiveHost } from '@shared/host/index.js';
import { WebHostPort } from '@host-web/index.js';
import { getActiveProfile, getCompareConfig } from '@shared/config/index.js';
import { synthesiseDocx } from '@features/compare/io/text/synthesise-docx.js';
import { extractParts } from '@features/compare/io/parts.js';
import { runPipeline } from '@features/compare/pipeline.js';
import { textCompareDriverState as d } from '@features/compare/driver.js';
import { WebCompare } from '../../../src/web/WebCompare.js';

const ORIGINAL = 'The Seller shall deliver the Goods.\nClause 2.\nClause 3.';
const REVISED = 'The Seller may deliver the Goods.\nClause 2.\nNEW CLAUSE.\nClause 3.';

let host: WebHostPort;
let container: HTMLDivElement;

beforeEach(() => {
  host = new WebHostPort();
  setActiveHost(host);
  container = document.createElement('div');
  document.body.appendChild(container);
  // Reset the shared text-driver signals.
  d.status.value = 'idle';
  d.errorMessage.value = null;
  d.ops.value = null;
  d.summary.value = null;
  d.originalText.value = '';
  d.revisedText.value = '';
  d.redlineReady.value = false;
});

afterEach(() => {
  render(null, container);
  container.remove();
  _resetActiveHost();
  vi.restoreAllMocks();
  delete (globalThis as { Word?: unknown }).Word;
});

describe('WebCompare — surface (T4/T5)', () => {
  it('renders the three input-mode tabs', () => {
    act(() => { render(<WebCompare />, container); });
    const tabs = [...container.querySelectorAll('.dp-compare__mode-btn')].map(b => b.textContent);
    expect(tabs).toEqual(['Word documents', 'Pasted text', 'PDFs']);
  });

  it('Word-vs-Word slots offer Upload but NOT "Use current document" (upload-only on web)', () => {
    act(() => { render(<WebCompare />, container); });
    // docx mode is the default tab.
    expect(container.querySelector('.dp-compare__upload')).toBeTruthy();
    expect(container.querySelector('.dp-compare__use-current')).toBeNull();
  });
});

describe('WebCompare — text-vs-text parity + Blob download (T6)', () => {
  it('the live driver yields the same ops/summary the add-in pipeline produces', async () => {
    // Independent reference run through the exact add-in pipeline, same profile.
    const config = getCompareConfig(getActiveProfile());
    const expected = await runPipeline({
      originalParts: extractParts(synthesiseDocx(ORIGINAL)),
      revisedParts: extractParts(synthesiseDocx(REVISED)),
      config,
      author: config.author,
      date: '2026-06-10T00:00:00Z',
    });

    d.originalText.value = ORIGINAL;
    d.revisedText.value = REVISED;
    await d.runCompare();

    expect(d.status.value).toBe('done');
    // Summary is date/author-independent → must match the reference exactly.
    expect(d.summary.value).toEqual(expected.summary);
    // Ops sequence (type + classification) matches the reference run.
    const shape = (ops: typeof expected.ops) => ops.map(o => `${o.type}:${o.classification ?? '-'}`);
    expect(shape(d.ops.value!)).toEqual(shape(expected.ops));
  });

  it('openRedline routes through WebHostPort.writeBack (Blob download), no Office call', async () => {
    const writeBackSpy = vi.spyOn(host, 'writeBack');
    const createObjectURL = vi.fn((_blob: Blob) => 'blob:web-compare');
    const revokeObjectURL = vi.fn();
    (globalThis.URL as unknown as { createObjectURL: unknown }).createObjectURL = createObjectURL;
    (globalThis.URL as unknown as { revokeObjectURL: unknown }).revokeObjectURL = revokeObjectURL;
    // Fail loudly if any Office bridge is reached on the web path.
    (globalThis as { Word?: unknown }).Word = new Proxy({}, {
      get() { throw new Error('Office (Word) must not be touched on the web Compare path'); },
    });

    d.originalText.value = ORIGINAL;
    d.revisedText.value = REVISED;
    await d.runCompare();
    expect(d.redlineReady.value).toBe(true);
    expect(writeBackSpy).not.toHaveBeenCalled(); // eager-generate, click-to-open

    await d.openRedline();

    // Routed through the host seam → WebHostPort.writeBack with the redline bundle.
    expect(writeBackSpy).toHaveBeenCalledTimes(1);
    const [bundleArg, optsArg] = writeBackSpy.mock.calls[0]!;
    expect(optsArg).toEqual({ mode: 'newDocument' });
    expect(typeof (bundleArg as { toUint8Array?: unknown }).toUint8Array).toBe('function');
    // The Blob download was triggered (no Office insert).
    expect(createObjectURL).toHaveBeenCalledTimes(1);
    expect(createObjectURL.mock.calls[0]![0]).toBeInstanceOf(Blob);
  });
});
