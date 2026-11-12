/**
 * C3 — Questionnaire fill tests.
 * C-T2: Values written via insertText; EFFECTIVE_DATE renders DD Month YYYY.
 *
 * HOST-ONLY: processInChunks / wordRun are mocked; the test verifies
 * that values are formatted and dispatched correctly.
 */
import { describe, it, expect, vi, beforeEach } from 'vitest';

vi.mock('@shared/office/index.js', () => ({
  wordRun: vi.fn(),
  batchLoad: vi.fn(),
  processInChunks: vi.fn(async (
    items: unknown[],
    _chunkSize: number,
    fn: (chunk: unknown[], ctx: unknown) => Promise<unknown[]>,
  ) => {
    // Simulated chunked processing.
    const allTags = ['CLIENT_NAME', 'EFFECTIVE_DATE', 'COMPANY_NUMBER'];
    const insertTextCalls: Record<string, string> = {};
    const ctx = {
      document: {
        contentControls: {
          load: vi.fn(),
          items: allTags.map(tag => ({
            tag,
            insertText: vi.fn((val: string) => { insertTextCalls[tag] = val; }),
          })),
        },
      },
      sync: vi.fn(async () => undefined),
    };
    await fn(items, ctx);
    return [];
  }),
}));

import { fillQuestionnaire, buildInitialAnswers } from '@features/content/templates/fill.js';
import type { TemplateSchema } from '@features/content/templates/types.js';

describe('C3 — fillQuestionnaire (C-T2)', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('formats date value as DD Month YYYY before insertText (C-T2)', async () => {
    const { processInChunks } = await import('@shared/office/index.js');
    const mockChunks = processInChunks as ReturnType<typeof vi.fn>;

    const insertedValues: Record<string, string> = {};

    mockChunks.mockImplementationOnce(async (
      items: Array<{ tag: string; type: string; rawValue: string | boolean | number }>,
      _chunkSize: number,
      fn: (chunk: unknown[], ctx: unknown) => Promise<unknown[]>,
    ) => {
      const ctx = {
        document: {
          contentControls: {
            load: vi.fn(),
            items: items.map(item => ({
              tag: item.tag,
              insertText: vi.fn((val: string) => { insertedValues[item.tag] = val; }),
            })),
          },
        },
        sync: vi.fn(async () => undefined),
      };
      await fn(items, ctx);
      return [];
    });

    await fillQuestionnaire([
      { tag: 'CLIENT_NAME', type: 'text', rawValue: 'Acme Ltd' },
      { tag: 'EFFECTIVE_DATE', type: 'date', rawValue: '2026-06-06' },
    ]);

    // Date must be rendered day-first.
    expect(insertedValues['EFFECTIVE_DATE']).toBe('06 June 2026');
    // Text is passed through trimmed.
    expect(insertedValues['CLIENT_NAME']).toBe('Acme Ltd');
  });

  it('returns empty result for empty items list', async () => {
    const result = await fillQuestionnaire([]);
    expect(result.written).toHaveLength(0);
    expect(result.notFound).toHaveLength(0);
  });
});

describe('C3 — buildInitialAnswers', () => {
  it('uses document current text if available', () => {
    const schema: TemplateSchema = {
      templateId: 't1',
      jurisdictionProfile: 'EW',
      variables: [
        { tag: 'CLIENT_NAME', label: 'Client', type: 'text', required: true },
        { tag: 'EFFECTIVE_DATE', label: 'Date', type: 'date', required: true },
      ],
      conditionals: [],
      repeating: [],
    };
    const docControls = [
      { tag: 'CLIENT_NAME', text: 'Existing Corp Ltd' },
      { tag: 'EFFECTIVE_DATE', text: '' },
    ];

    const answers = buildInitialAnswers(schema, docControls);
    expect(answers['CLIENT_NAME']).toBe('Existing Corp Ltd');
    // Empty text falls back to default (empty string).
    expect(answers['EFFECTIVE_DATE']).toBe('');
  });

  it('uses variable default when document text is empty', () => {
    const schema: TemplateSchema = {
      templateId: 't2',
      jurisdictionProfile: 'EW',
      variables: [
        { tag: 'ENTITY_TYPE', label: 'Entity type', type: 'entityType', required: true, default: 'Limited' },
      ],
      conditionals: [],
      repeating: [],
    };

    const answers = buildInitialAnswers(schema, [{ tag: 'ENTITY_TYPE', text: '' }]);
    expect(answers['ENTITY_TYPE']).toBe('Limited');
  });
});
