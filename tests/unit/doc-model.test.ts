import { describe, it, expect, vi, beforeEach } from 'vitest';

// Mock the loader so DocModel.load() doesn't touch Office JS.
vi.mock('@shared/doc-model/loader.js', () => ({
  loadDocSnapshot: vi.fn(async () => ({
    paragraphs: [
      {
        id: 'body-para-0',
        index: 0,
        text: 'Clause 1 text',
        styleName: 'Normal',
        listString: '1',
        listLevel: 0,
        container: 'body' as const,
        getRange: () => { throw new Error('not used'); },
      },
    ],
    footnotes: [],
    loadedAt: 0,
  })),
}));

import { DocModel } from '@shared/doc-model/index.js';

describe('DocModel', () => {
  beforeEach(() => {
    // Word global is undefined in jsdom → subscribeToEvents is a no-op, safe.
  });

  it('loads a snapshot and exposes paragraphs', async () => {
    await DocModel.load();
    expect(DocModel.paragraphs.value).toHaveLength(1);
    expect(DocModel.paragraphs.value[0]?.id).toBe('body-para-0');
  });

  it('resolves a paragraph by id', async () => {
    await DocModel.load();
    expect(DocModel.getById('body-para-0')?.text).toBe('Clause 1 text');
    expect(DocModel.getById('missing')).toBeUndefined();
  });

  it('fires onInvalidate callbacks after a load', async () => {
    const cb = vi.fn();
    const unsubscribe = DocModel.onInvalidate(cb);
    await DocModel.load();
    expect(cb).toHaveBeenCalled();
    unsubscribe();
  });

  it('stops calling a callback after unsubscribe', async () => {
    const cb = vi.fn();
    const unsubscribe = DocModel.onInvalidate(cb);
    unsubscribe();
    await DocModel.load();
    expect(cb).not.toHaveBeenCalled();
  });
});
