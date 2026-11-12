import { describe, it, expect, vi, afterEach } from 'vitest';
import { OfficeHostPort } from '@host-office/index.js';
import { ooxmlService } from '@shared/ooxml/index.js';
import { _resetCapabilities, detectCapabilities } from '@shared/office/index.js';

// ---------------------------------------------------------------------------
// Minimal Word mock shared by loadDocument (single-pass loader) and reveal
// (resolveAnchor → body.search → select). Only the surface those paths touch.
// ---------------------------------------------------------------------------
function mockPara(text: string) {
  return {
    text,
    style: 'Normal',
    load() {},
    getRange() { return {}; },
    listItemOrNullObject: { isNullObject: true, load() {} },
    parentTableCellOrNullObject: { isNullObject: true, load() {} },
  };
}

function installWordMock(texts: string[]) {
  // Shared call-log so tests can assert ordering (select MUST follow the first
  // sync that materialises the deferred search hit — the reason for two syncs).
  const calls: string[] = [];
  const select = vi.fn(() => { calls.push('select'); });
  const hit = { select, scrollIntoView() {}, insertText() {} };
  const search = vi.fn(() => ({ load() {}, items: [hit] }));
  const ctx = {
    document: {
      body: {
        paragraphs: { load() {}, items: texts.map(mockPara) },
        footnotes: { load() {}, items: [] as unknown[] },
        getRange: () => ({}),
        search,
      },
    },
    sync: async () => { calls.push('sync'); },
  };
  (globalThis as unknown as { Word: unknown }).Word = {
    run: async (fn: (c: typeof ctx) => unknown) => fn(ctx),
  };
  return { select, search, calls };
}

afterEach(() => {
  delete (globalThis as unknown as { Word?: unknown }).Word;
  delete (globalThis as unknown as { Office?: unknown }).Office;
  _resetCapabilities();
  vi.restoreAllMocks();
});

/**
 * A document with stories (ADR-0047): each story a body whose search finds
 * `text` only where it is. Records which story each search ran in.
 */
function installStoriesMock(where: string, opts: { wordApi15?: boolean; sections?: number } = {}) {
  const searched: string[] = [];
  const selected: string[] = [];
  const story = (name: string) => ({
    search: vi.fn(() => {
      searched.push(name);
      return { load() {}, items: name === where ? [{ select: () => { selected.push(name); } }] : [] };
    }),
  });
  const kinds = { Primary: 'default', FirstPage: 'first', EvenPages: 'even' } as const;
  const sections = Array.from({ length: opts.sections ?? 1 }, (_, k) => ({
    getHeader: (t: keyof typeof kinds) => story(`header:s${k}:${kinds[t]}`),
    getFooter: (t: keyof typeof kinds) => story(`footer:s${k}:${kinds[t]}`),
  }));
  const ctx = {
    document: {
      body: {
        ...story('body'),
        footnotes: { load() {}, items: [{ body: story('footnote:1') }, { body: story('footnote:2') }] },
        endnotes: { load() {}, items: [{ body: story('endnote:1') }] },
      },
      sections: { load() {}, items: sections },
    },
    sync: async () => {},
  };
  (globalThis as unknown as { Word: unknown }).Word = { run: async (fn: (c: typeof ctx) => unknown) => fn(ctx) };
  (globalThis as unknown as { Office: unknown }).Office = {
    context: { requirements: { isSetSupported: (_: string, v: string) => v !== '1.5' || opts.wordApi15 === true } },
  };
  _resetCapabilities();
  detectCapabilities();
  return { searched, selected };
}

describe('OfficeHostPort', () => {
  it('loadDocument loads the snapshot via the single-pass loader', async () => {
    installWordMock(['Hello world', 'Second paragraph']);
    const host = new OfficeHostPort();
    const snap = await host.loadDocument();
    expect(snap.paragraphs).toHaveLength(2);
    expect(snap.paragraphs[0]!.text).toBe('Hello world');
    expect(snap.paragraphs[0]!.id).toBe('body-para-0');
  });

  it('reveal resolves the located span and selects it after the first sync', async () => {
    const { select, calls } = installWordMock(['Hello world']);
    const host = new OfficeHostPort();
    await host.loadDocument();
    calls.length = 0; // ignore the loader's syncs
    await host.reveal({ paraId: 'body-para-0', charStart: 0, charEnd: 5, label: 'x' });
    // The deferred-search hit's select() is the office navigation primitive,
    // and it must run AFTER a sync materialises the search hit (the double-sync
    // contract). Sequence: sync → select → sync.
    expect(select).toHaveBeenCalledTimes(1);
    expect(calls).toEqual(['sync', 'select', 'sync']);
  });

  it('findAndReveal searches the body and selects the first hit after the first sync', async () => {
    const { select, search, calls } = installWordMock(['Hello world']);
    const host = new OfficeHostPort();
    await host.findAndReveal('Hello');
    expect(search).toHaveBeenCalledWith('Hello', { matchCase: false });
    expect(select).toHaveBeenCalledTimes(1);
    // The hit's select() must follow the first sync that materialises the search.
    expect(calls).toEqual(['sync', 'select', 'sync']);
  });

  it('findAndReveal with a header scope searches that section’s header of that type first', async () => {
    const { searched, selected } = installStoriesMock('header:s1:first', { sections: 2 });
    await new OfficeHostPort().findAndReveal('Draft', { story: 'header', section: 1, type: 'first' });
    expect(searched[0]).toBe('header:s1:first');
    expect(selected).toEqual(['header:s1:first']);
  });

  it('findAndReveal falls back to the other headers, then the main text, when the scoped one has no hit', async () => {
    const { searched, selected } = installStoriesMock('header:s0:default', { sections: 2 });
    await new OfficeHostPort().findAndReveal('Draft', { story: 'header', section: 1, type: 'default' });
    expect(selected).toEqual(['header:s0:default']);
    const none = installStoriesMock('nowhere', { sections: 1 });
    await new OfficeHostPort().findAndReveal('Draft', { story: 'footer', section: 0, type: 'default' });
    expect(none.searched).toEqual(['footer:s0:default', 'footer:s0:first', 'footer:s0:even', 'body']);
    expect(searched.every(s => s.startsWith('header') || s === 'body')).toBe(true);
  });

  it('findAndReveal searches the notes where WordApi 1.5 offers them', async () => {
    const { searched, selected } = installStoriesMock('footnote:2', { wordApi15: true });
    await new OfficeHostPort().findAndReveal('GST', { story: 'footnote' });
    expect(searched).toEqual(['footnote:1', 'footnote:2']);
    expect(selected).toEqual(['footnote:2']);
  });

  it('findAndReveal falls back to the main text for a note without WordApi 1.5', async () => {
    const { searched } = installStoriesMock('footnote:2', { wordApi15: false });
    await new OfficeHostPort().findAndReveal('GST', { story: 'endnote' });
    expect(searched).toEqual(['body']);
  });

  it('findAndReveal is a no-op for empty / whitespace-only search text', async () => {
    const { select, search } = installWordMock(['Hello world']);
    const host = new OfficeHostPort();
    await host.findAndReveal('   ');
    expect(search).not.toHaveBeenCalled();
    expect(select).not.toHaveBeenCalled();
  });

  it('reveal is a no-op (no select) when the location has no paraId', async () => {
    const { select } = installWordMock(['Hello world']);
    const host = new OfficeHostPort();
    await host.loadDocument();
    await host.reveal({ label: 'document-level' });
    expect(select).not.toHaveBeenCalled();
  });

  it('reveal swallows a stale location (paragraph absent from snapshot)', async () => {
    installWordMock(['Hello world']);
    const host = new OfficeHostPort();
    await host.loadDocument();
    // resolveAnchor throws for an unknown paragraph id; reveal must not reject.
    await expect(
      host.reveal({ paraId: 'body-para-999', charStart: 0, charEnd: 3, label: 'gone' }),
    ).resolves.toBeUndefined();
  });

  it('onDocumentChanged returns a disposable subscription', () => {
    installWordMock(['x']);
    const host = new OfficeHostPort();
    const cb = vi.fn();
    const sub = host.onDocumentChanged(cb);
    expect(typeof sub.dispose).toBe('function');
    sub.dispose(); // must not throw
  });

  it('readDocx / writeBack delegate to the OOXML Service', async () => {
    installWordMock(['x']);
    const fakeBundle = {} as never;
    const read = vi.spyOn(ooxmlService, 'readDocx').mockResolvedValue(fakeBundle);
    const write = vi.spyOn(ooxmlService, 'writeBack').mockResolvedValue();
    const host = new OfficeHostPort();
    await host.readDocx();
    await host.writeBack(fakeBundle);
    expect(read).toHaveBeenCalledTimes(1);
    expect(write).toHaveBeenCalledWith(fakeBundle, undefined);
  });

  it('exposes the storage provider', () => {
    installWordMock(['x']);
    const host = new OfficeHostPort();
    expect(host.storage).toBeDefined();
    expect(host.storage.local).toBeDefined();
  });
});
