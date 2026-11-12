/**
 * host-web — WebHostPort adapter.
 *
 * Exercises the browser-backed HostPort surface: ingest → loadDocument,
 * document-changed notifications, reveal (DOM scroll + highlight), and writeBack
 * (Blob download). jsdom lacks URL.createObjectURL / Element.scrollIntoView, so
 * those are stubbed per-test.
 */
import { describe, it, expect, vi, afterEach } from 'vitest';
import { synthesiseDocx } from '@features/compare/io/text/synthesise-docx.js';
import { WebHostPort } from '@host-web/index.js';

/** A real .docx ArrayBuffer from a plain-text blob (one <w:p> per line). */
function docxBuffer(text: string): ArrayBuffer {
  const u8 = new Uint8Array(synthesiseDocx(text).toUint8Array());
  return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
}

afterEach(() => {
  document.body.innerHTML = '';
  vi.restoreAllMocks();
});

describe('WebHostPort', () => {
  it('ingest → loadDocument parses the held bundle into a snapshot', async () => {
    const host = new WebHostPort();
    host.ingest(docxBuffer('Hello world\nSecond paragraph'));
    const snap = await host.loadDocument();
    expect(snap.paragraphs.map(p => p.text)).toEqual(['Hello world', 'Second paragraph']);
    expect(snap.paragraphs[0]!.id).toBe('body-para-0');
  });

  it('loadDocument before ingest throws', async () => {
    await expect(new WebHostPort().loadDocument()).rejects.toThrow(/ingest/);
  });

  it('onDocumentChanged fires on ingest and stops after dispose', () => {
    const host = new WebHostPort();
    const cb = vi.fn();
    const sub = host.onDocumentChanged(cb);
    host.ingest(docxBuffer('a'));
    expect(cb).toHaveBeenCalledTimes(1);
    sub.dispose();
    host.ingest(docxBuffer('b'));
    expect(cb).toHaveBeenCalledTimes(1);
  });

  it('reveal scrolls to + highlights the paragraph element', async () => {
    document.body.innerHTML =
      '<p data-para-id="body-para-0">x</p><p data-para-id="body-para-1">y</p>';
    const el0 = document.querySelector<HTMLElement>('[data-para-id="body-para-0"]')!;
    el0.scrollIntoView = vi.fn();
    const host = new WebHostPort();
    await host.reveal({ paraId: 'body-para-0', label: 'x' });
    expect(el0.classList.contains('dp-web-doc__para--active')).toBe(true);
    expect(el0.scrollIntoView).toHaveBeenCalledTimes(1);
  });

  it('reveal is a no-op for a missing element / no paraId', async () => {
    const host = new WebHostPort();
    await expect(host.reveal({ label: 'document-level' })).resolves.toBeUndefined();
    await expect(host.reveal({ paraId: 'body-para-999', label: 'gone' })).resolves.toBeUndefined();
  });

  it('writeBack downloads a .docx Blob with the ingested filename', async () => {
    const created = vi.fn((_blob: Blob) => 'blob:fake');
    const revoked = vi.fn();
    (URL as unknown as { createObjectURL: unknown }).createObjectURL = created;
    (URL as unknown as { revokeObjectURL: unknown }).revokeObjectURL = revoked;

    const realCreate = document.createElement.bind(document);
    let anchor: HTMLAnchorElement | undefined;
    vi.spyOn(document, 'createElement').mockImplementation((tag: string) => {
      const el = realCreate(tag as 'a');
      if (tag === 'a') {
        anchor = el as HTMLAnchorElement;
        anchor.click = vi.fn();
      }
      return el;
    });

    const host = new WebHostPort();
    host.ingest(docxBuffer('Hello'), 'contract.docx');
    const bundle = await host.readDocx();
    await host.writeBack(bundle);

    expect(created).toHaveBeenCalledTimes(1);
    expect(created.mock.calls[0]![0]).toBeInstanceOf(Blob);
    expect(anchor?.download).toBe('contract.docx');
    expect(anchor?.click).toHaveBeenCalledTimes(1);
    expect(revoked).toHaveBeenCalledTimes(1);
  });

  it('exposes a storage provider', () => {
    const host = new WebHostPort();
    expect(host.storage.local).toBeDefined();
    expect(host.storage.docSettings).toBeDefined();
  });
});
