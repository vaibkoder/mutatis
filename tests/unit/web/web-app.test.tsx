/**
 * apps/web — WebApp shell.
 *
 * Drives the real end-to-end UI flow with a real WebHostPort: render → pick a
 * `.docx` → the checker runs → findings render via the (reused) FindingRow → the
 * document view shows the paragraphs. No Office anywhere.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { render } from 'preact';
import { act } from 'preact/test-utils';
import { synthesiseDocx } from '@features/compare/io/text/synthesise-docx.js';
import { WebHostPort } from '@host-web/index.js';
import { WebApp } from '../../../src/web/WebApp.js';

const DOCX_MIME =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

let container: HTMLDivElement;

beforeEach(() => {
  container = document.createElement('div');
  document.body.appendChild(container);
});

afterEach(() => {
  render(null, container);
  container.remove();
});

/** Build a File whose bytes are a real .docx (one <w:p> per line). */
function docxFile(text: string, name = 'test.docx'): File {
  const bytes = new Uint8Array(synthesiseDocx(text).toUint8Array());
  return new File([bytes], name, { type: DOCX_MIME });
}

async function flush(): Promise<void> {
  await act(async () => {
    await new Promise(r => setTimeout(r, 0));
    await new Promise(r => setTimeout(r, 0));
  });
}

describe('WebApp', () => {
  it('renders the file picker and an idle prompt before any document', () => {
    act(() => { render(<WebApp host={new WebHostPort()} />, container); });
    expect(container.textContent).toContain('Choose .docx');
    expect(container.textContent).toContain('Choose a Word');
    expect(container.querySelector('input[type="file"]')).not.toBeNull();
  });

  it('after picking a .docx, runs the checker and renders findings + document view', async () => {
    const host = new WebHostPort();
    act(() => { render(<WebApp host={host} />, container); });

    const input = container.querySelector<HTMLInputElement>('input[type="file"]')!;
    const file = docxFile('The completion date is [●].');
    Object.defineProperty(input, 'files', { value: [file], configurable: true });

    await act(async () => {
      input.dispatchEvent(new Event('change', { bubbles: true }));
      await new Promise(r => setTimeout(r, 0));
    });
    await flush();

    // Document view rendered the paragraph (keyed by the loader's para id).
    expect(container.querySelector('[data-para-id="body-para-0"]')).not.toBeNull();
    // The placeholder finding rendered via the reused FindingRow.
    expect(container.querySelector('.dp-finding-row')).not.toBeNull();
    expect(container.textContent?.toLowerCase()).toContain('placeholder');
  });

  it('lights up cross-reference + defined-terms alongside factual consistency (phase 3b)', async () => {
    const host = new WebHostPort();
    act(() => { render(<WebApp host={host} />, container); });

    const input = container.querySelector<HTMLInputElement>('input[type="file"]')!;
    // A broken cross-reference, a prose definition that is never used, and a
    // factual placeholder — one document exercising all three checkers.
    const file = docxFile(
      [
        'Subject to Clause 9, the parties shall perform their obligations.',
        '"Widget" means a small mechanical device.',
        'The completion date is [●].',
      ].join('\n'),
    );
    Object.defineProperty(input, 'files', { value: [file], configurable: true });

    await act(async () => {
      input.dispatchEvent(new Event('change', { bubbles: true }));
      await new Promise(r => setTimeout(r, 0));
    });
    await flush();

    const text = container.textContent ?? '';
    // Cross-reference feature group rendered (Clause 9 is not-found).
    expect(text).toContain('Cross-references');
    expect(text).toContain('Broken (not found)');
    // Defined-terms feature group rendered (Widget defined, never used).
    expect(text).toContain('Defined terms');
    // Factual consistency still runs.
    expect(text.toLowerCase()).toContain('placeholder');
    // REF-field degraded banner is shown (no live Word host).
    expect(text).toContain('REF-field detection is not available outside Word');
  });

  it('lights up party-role consistency on web (feature 09)', async () => {
    const host = new WebHostPort();
    act(() => { render(<WebApp host={host} />, container); });

    const input = container.querySelector<HTMLInputElement>('input[type="file"]')!;
    // One entity introduced under two different roles → inconsistent-label finding.
    const file = docxFile(
      [
        'ACME Pte Ltd (the "Buyer") is a party to this agreement.',
        'ACME Pte Ltd (the "Purchaser") shall complete the transaction.',
      ].join('\n'),
    );
    Object.defineProperty(input, 'files', { value: [file], configurable: true });

    await act(async () => {
      input.dispatchEvent(new Event('change', { bubbles: true }));
      await new Promise(r => setTimeout(r, 0));
    });
    await flush();

    const text = container.textContent ?? '';
    // The party-role group rendered with its category label and the two role names.
    expect(text).toContain('Same party, two labels');
    expect(text).toContain('Buyer');
    expect(text).toContain('Purchaser');
  });

  it('top-level Proof | Compare switch toggles between the findings flow and WebCompare (12f T5)', () => {
    act(() => { render(<WebApp host={new WebHostPort()} />, container); });

    // Default view is Proof: the file picker is shown, no Compare panel.
    expect(container.querySelector('.dp-web__file')).not.toBeNull();
    expect(container.querySelector('.dp-compare')).toBeNull();

    const [proofBtn, compareBtn] = [
      ...container.querySelectorAll<HTMLButtonElement>('.dp-web__view-btn'),
    ];
    expect(proofBtn?.textContent).toBe('Proof');
    expect(compareBtn?.textContent).toBe('Compare');

    // Switch to Compare → WebCompare renders (styled by dp-compare), Proof toolbar gone.
    act(() => { compareBtn!.click(); });
    expect(container.querySelector('.dp-compare')).not.toBeNull();
    expect(container.querySelector('.dp-compare__mode-btn')).not.toBeNull();
    expect(container.querySelector('.dp-web__file')).toBeNull();

    // Switch back to Proof → findings flow returns, Compare panel gone.
    act(() => { proofBtn!.click(); });
    expect(container.querySelector('.dp-web__file')).not.toBeNull();
    expect(container.querySelector('.dp-compare')).toBeNull();
  });
});
