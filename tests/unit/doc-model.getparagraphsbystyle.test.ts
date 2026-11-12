import { describe, it, expect, vi } from 'vitest';

// Mock the loader so DocModel.load() doesn't touch Office JS.
vi.mock('@shared/doc-model/loader.js', () => ({
  loadDocSnapshot: vi.fn(async () => ({
    paragraphs: [
      {
        id: 'body-para-0', index: 0, text: 'Schedule 1', styleName: 'Schedule Heading',
        container: 'body' as const, getRange: () => { throw new Error('unused'); },
      },
      {
        id: 'body-para-1', index: 1, text: 'Body text', styleName: 'Normal',
        container: 'body' as const, getRange: () => { throw new Error('unused'); },
      },
      {
        id: 'body-para-2', index: 2, text: 'Schedule 2', styleName: 'schedule heading',
        container: 'body' as const, getRange: () => { throw new Error('unused'); },
      },
      {
        id: 'body-para-3', index: 3, text: 'A heading', styleName: 'Heading 1',
        container: 'body' as const, getRange: () => { throw new Error('unused'); },
      },
    ],
    footnotes: [
      {
        id: 'fn-para-0', index: 0, text: 'footnote', styleName: 'Footnote Text',
        container: 'footnote' as const, getRange: () => { throw new Error('unused'); },
      },
    ],
    loadedAt: 0,
  })),
}));

import { DocModel } from '@shared/doc-model/index.js';

describe('DocModel.getParagraphsByStyle', () => {
  it('returns paragraphs of the given style in document order, case-insensitively', async () => {
    await DocModel.load();
    const result = DocModel.getParagraphsByStyle(['Schedule Heading']);
    expect(result.map(p => p.id)).toEqual(['body-para-0', 'body-para-2']);
  });

  it('returns an empty array for an unknown style', async () => {
    await DocModel.load();
    expect(DocModel.getParagraphsByStyle(['Nonexistent Style'])).toEqual([]);
  });

  it('returns an empty array when no styles are requested', async () => {
    await DocModel.load();
    expect(DocModel.getParagraphsByStyle([])).toEqual([]);
  });

  it('matches across multiple styles and spans body + footnotes', async () => {
    await DocModel.load();
    const result = DocModel.getParagraphsByStyle(['Heading 1', 'Footnote Text']);
    expect(result.map(p => p.id)).toEqual(['body-para-3', 'fn-para-0']);
  });
});
