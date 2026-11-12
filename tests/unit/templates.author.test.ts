/**
 * C2 — Mark variable authoring tests.
 * C-T1: two PlainText controls created with correct tag/title; schema persisted.
 *
 * HOST-ONLY: actual Office JS insertContentControl is not called here — the
 * Word.run context is mocked.  The logic under test is tag validation,
 * duplicate detection, and schema persistence.
 */
import { describe, it, expect, vi, beforeEach } from 'vitest';

// Mock @shared/office so wordRun can be controlled.
vi.mock('@shared/office/index.js', () => ({
  wordRun: vi.fn(async (fn: (ctx: unknown) => Promise<unknown>) => {
    // Simulate a minimal Word.RequestContext.
    const cc = {
      tag: '',
      title: '',
      appearance: '',
    };
    const ctx = {
      document: {
        getSelection: () => ({
          insertContentControl: vi.fn(() => cc),
        }),
      },
      sync: vi.fn(async () => undefined),
    };
    return fn(ctx);
  }),
  batchLoad: vi.fn(),
  processInChunks: vi.fn(),
}));

// Mock @shared/storage to use in-memory localStorage.
vi.mock('@shared/storage/index.js', async () => {
  const { createLocalStore } = await import('@shared/storage/local.js');
  const store = createLocalStore();
  return {
    getStorageProvider: () => ({ local: store, docSettings: store }),
  };
});

import { validateTag, markVariable, loadSchema } from '@features/content/templates/author.js';

beforeEach(() => {
  localStorage.clear();
});

describe('C2 — validateTag', () => {
  it('accepts UPPER_SNAKE tags', () => {
    expect(validateTag('CLIENT_NAME')).toBeNull();
    expect(validateTag('EFFECTIVE_DATE')).toBeNull();
    expect(validateTag('DIRECTOR_1_NAME')).toBeNull();
  });

  it('rejects lowercase tags', () => {
    expect(validateTag('client_name')).toBeTruthy();
  });

  it('rejects tags that do not start with a letter', () => {
    expect(validateTag('1CLIENT')).toBeTruthy();
    expect(validateTag('_CLIENT')).toBeTruthy();
  });

  it('rejects empty tags', () => {
    expect(validateTag('')).toBeTruthy();
  });
});

describe('C2 — markVariable (C-T1)', () => {
  it('persists variable to schema after marking', async () => {
    const result = await markVariable({
      tag: 'CLIENT_NAME',
      label: 'Client name',
      variable: {
        tag: 'CLIENT_NAME',
        label: 'Client name',
        type: 'text',
        required: true,
      },
    });

    expect(result.ok).toBe(true);

    const schema = await loadSchema();
    expect(schema.variables).toHaveLength(1);
    expect(schema.variables[0]?.tag).toBe('CLIENT_NAME');
    expect(schema.variables[0]?.label).toBe('Client name');
    expect(schema.variables[0]?.type).toBe('text');
  });

  it('rejects duplicate tags (C-T1 uniqueness)', async () => {
    await markVariable({
      tag: 'EFFECTIVE_DATE',
      label: 'Effective date',
      variable: { tag: 'EFFECTIVE_DATE', label: 'Effective date', type: 'date', required: true },
    });

    const result = await markVariable({
      tag: 'EFFECTIVE_DATE',
      label: 'Another',
      variable: { tag: 'EFFECTIVE_DATE', label: 'Another', type: 'text', required: false },
    });

    expect(result.ok).toBe(false);
    expect(result.error).toMatch(/already exists/i);
  });

  it('can mark two variables with distinct tags', async () => {
    const r1 = await markVariable({
      tag: 'CLIENT_NAME',
      label: 'Client name',
      variable: { tag: 'CLIENT_NAME', label: 'Client name', type: 'text', required: true },
    });
    const r2 = await markVariable({
      tag: 'EFFECTIVE_DATE',
      label: 'Effective date',
      variable: { tag: 'EFFECTIVE_DATE', label: 'Effective date', type: 'date', required: true },
    });

    expect(r1.ok).toBe(true);
    expect(r2.ok).toBe(true);

    const schema = await loadSchema();
    expect(schema.variables).toHaveLength(2);
    expect(schema.variables.map(v => v.tag)).toContain('CLIENT_NAME');
    expect(schema.variables.map(v => v.tag)).toContain('EFFECTIVE_DATE');
  });
});
