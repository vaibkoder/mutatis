/**
 * E — CopyButton (feature 11 — spec §6.1, plan T14).
 *
 * The copy/download pair is the only path document content leaves Mutatis,
 * so the load-bearing assertions are: the clipboard receives the *formatted*
 * envelope (F7), the format toggle switches the serialisation, the download
 * anchor carries the right MIME + extension + filename, and a `disabledReason`
 * hard-blocks both actions. jsdom has no `navigator.clipboard`, so we install a
 * mock. `exportFormat`/`noticeVisible` are module-level signals — reset each test.
 */
import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { render } from 'preact';
import { act } from 'preact/test-utils';

import { CopyButton } from '../../../../src/features/ai-export/ui/CopyButton.js';
import { formatEnvelope } from '../../../../src/features/ai-export/index.js';
import {
  exportFormat, noticeVisible, dismissAwarenessNotice,
} from '../../../../src/features/ai-export/ui/prefs.js';
import type { AiExportEnvelope } from '../../../../src/features/ai-export/types.js';

const ENV: AiExportEnvelope = {
  _mutatis: '1.0',
  generated: '2026-06-09T10:00:00.000Z',
  document: 'Acme SPA.docx',
  jurisdiction: 'SG',
  checks_run: ['defined-terms'],
  blind_spots: ['Only defined terms were exported.'],
};

const writeText = vi.fn((_text: string) => Promise.resolve());
let blobs: Blob[];

let container: HTMLDivElement;

beforeEach(() => {
  // jsdom ships no clipboard / object-URL plumbing — stub both.
  Object.defineProperty(navigator, 'clipboard', { value: { writeText }, configurable: true });
  writeText.mockClear();
  blobs = [];
  URL.createObjectURL = vi.fn((b: Blob) => { blobs.push(b); return 'blob:mock'; }) as typeof URL.createObjectURL;
  URL.revokeObjectURL = vi.fn();

  // Module-level signals persist across tests — reset to defaults.
  exportFormat.value = 'markdown';
  dismissAwarenessNotice();

  container = document.createElement('div');
  document.body.appendChild(container);
});

afterEach(() => {
  vi.restoreAllMocks();
  render(null, container);
  container.remove();
});

function mount(props: Partial<Parameters<typeof CopyButton>[0]> = {}) {
  act(() => {
    render(<CopyButton build={() => ENV} fileLabel="defined-terms" {...props} />, container);
  });
}
const copyBtn = () => container.querySelector('.dp-ai-export__copy-btn') as HTMLButtonElement;
const downloadBtn = () => container.querySelector('.dp-ai-export__download-btn') as HTMLButtonElement;

describe('E — CopyButton', () => {
  it('copies the formatted envelope (Markdown by default) to the clipboard', () => {
    mount();
    act(() => { copyBtn().click(); });
    expect(writeText).toHaveBeenCalledTimes(1);
    expect(writeText).toHaveBeenCalledWith(formatEnvelope(ENV, 'markdown'));
  });

  it('raises the one-time awareness notice on the first copy', async () => {
    mount();
    expect(noticeVisible.value).toBe(false);
    // The notice is raised only after the awaited clipboard write resolves.
    await act(async () => {
      copyBtn().click();
      await new Promise(r => setTimeout(r, 0));
    });
    expect(noticeVisible.value).toBe(true);
    expect(localStorage.getItem('ai-export.noticeShown')).toBe('true');
  });

  it('switches the serialisation when the format toggle flips to JSON', () => {
    mount();
    const jsonOpt = container.querySelector('[id$="-json"]') as HTMLButtonElement;
    act(() => { jsonOpt.click(); });
    act(() => { copyBtn().click(); });
    const copied = writeText.mock.calls[0]![0];
    expect(copied).toBe(formatEnvelope(ENV, 'json'));
    expect(copied.startsWith('{')).toBe(true);
  });

  it('builds a download anchor with the correct MIME, extension and filename', () => {
    const realCreate = document.createElement.bind(document);
    const anchors: HTMLAnchorElement[] = [];
    vi.spyOn(document, 'createElement').mockImplementation(((tag: string) => {
      const el = realCreate(tag);
      if (tag === 'a') { (el as HTMLAnchorElement).click = vi.fn(); anchors.push(el as HTMLAnchorElement); }
      return el;
    }) as typeof document.createElement);

    mount();
    act(() => { downloadBtn().click(); });

    expect(blobs).toHaveLength(1);
    expect(blobs[0]!.type).toBe('text/markdown');
    expect(anchors).toHaveLength(1);
    expect(anchors[0]!.download).toBe('mutatis-export-defined-terms-2026-06-09.md');
    expect(anchors[0]!.click).toHaveBeenCalledTimes(1);
  });

  it('disables both buttons and shows the tooltip when disabledReason is set', () => {
    mount({ disabledReason: 'Run Defined Terms first' });
    expect(copyBtn().disabled).toBe(true);
    expect(downloadBtn().disabled).toBe(true);
    expect(copyBtn().getAttribute('title')).toBe('Run Defined Terms first');
    act(() => { copyBtn().click(); });
    expect(writeText).not.toHaveBeenCalled();
  });

  it('does nothing when build() returns null (no current results)', () => {
    act(() => {
      render(<CopyButton build={() => null} fileLabel="defined-terms" />, container);
    });
    act(() => { copyBtn().click(); });
    expect(writeText).not.toHaveBeenCalled();
  });
});
