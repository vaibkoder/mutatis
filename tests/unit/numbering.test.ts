import { describe, it, expect } from 'vitest';
import {
  composeFormat,
  mapToWordListNumbering,
  previewLevel,
  validateScheme,
  duplicateScheme,
  SCHEMA_VERSION,
  BUILTIN_IDS,
} from '@features/numbering/schema.js';
import type { NumberingScheme } from '@features/numbering/types.js';
import { UK_PROFILE } from '@shared/config/index.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makeScheme(overrides: Partial<NumberingScheme> = {}): NumberingScheme {
  return {
    id: 'test-scheme',
    name: 'Test Scheme',
    jurisdiction: 'SG-UK',
    schemaVersion: SCHEMA_VERSION,
    suppressForSgUk: false,
    levels: [
      { levelIndex: 0, listNumbering: 'arabic',      formatString: '%1.',    prefix: '', suffix: '', indent: { textIndentPt: 0, numberIndentPt: 0 }, labelStyle: 'decimal-compound', restartAfterLevel: -1 },
      { levelIndex: 1, listNumbering: 'arabic',      formatString: '%1.%2.', prefix: '', suffix: '', indent: { textIndentPt: 36, numberIndentPt: 18 }, labelStyle: 'decimal-compound', restartAfterLevel: 0 },
      { levelIndex: 2, listNumbering: 'lowerLetter', formatString: '(%3)',   prefix: '', suffix: '', indent: { textIndentPt: 72, numberIndentPt: 36 }, labelStyle: 'paren-lower-alpha', restartAfterLevel: 1 },
    ],
    ...overrides,
  };
}

// ---------------------------------------------------------------------------
// composeFormat
// ---------------------------------------------------------------------------

describe('composeFormat', () => {
  it('returns formatString verbatim when prefix/suffix are empty', () => {
    expect(composeFormat('', '%1.%2.', '')).toBe('%1.%2.');
  });

  it('prepends prefix and appends suffix', () => {
    expect(composeFormat('(', '%1', ')')).toBe('(%1)');
  });

  it('preserves %n tokens intact', () => {
    const result = composeFormat('', '%1.%2.%3.', '');
    expect(result).toBe('%1.%2.%3.');
  });
});

// ---------------------------------------------------------------------------
// mapToWordListNumbering
// ---------------------------------------------------------------------------

describe('mapToWordListNumbering', () => {
  it('maps known styles to capitalized Word.ListNumbering string literals', () => {
    expect(mapToWordListNumbering('arabic')).toBe('Arabic');
    expect(mapToWordListNumbering('lowerLetter')).toBe('LowerLetter');
    expect(mapToWordListNumbering('upperRoman')).toBe('UpperRoman');
    expect(mapToWordListNumbering('none')).toBe('None');
  });

  it('throws on unknown style', () => {
    expect(() => mapToWordListNumbering('bogus')).toThrow();
  });
});

// ---------------------------------------------------------------------------
// previewLevel
// ---------------------------------------------------------------------------

describe('previewLevel', () => {
  it('renders "1." for level 0 arabic compound', () => {
    const scheme = makeScheme();
    expect(previewLevel(scheme, 0)).toBe('1.');
  });

  it('renders "1.1." for level 1 compound decimal (default counters = 1)', () => {
    const scheme = makeScheme();
    expect(previewLevel(scheme, 1)).toBe('1.1.');
  });

  it('renders "(a)" for level 2 paren-lower-alpha', () => {
    const scheme = makeScheme();
    expect(previewLevel(scheme, 2)).toBe('(a)');
  });

  it('renders "(i)" for lowerRoman parenthesised', () => {
    const scheme = makeScheme({
      levels: [
        { levelIndex: 0, listNumbering: 'arabic', formatString: '%1.', prefix: '', suffix: '', indent: { textIndentPt: 0, numberIndentPt: 0 }, labelStyle: 'decimal-compound', restartAfterLevel: -1 },
        { levelIndex: 1, listNumbering: 'lowerRoman', formatString: '(%2)', prefix: '', suffix: '', indent: { textIndentPt: 36, numberIndentPt: 18 }, labelStyle: 'paren-lower-roman', restartAfterLevel: 0 },
      ],
    });
    expect(previewLevel(scheme, 1)).toBe('(i)');
  });

  it('renders "(A)" for upperLetter parenthesised', () => {
    const scheme = makeScheme({
      levels: [
        { levelIndex: 0, listNumbering: 'upperLetter', formatString: '(%1)', prefix: '', suffix: '', indent: { textIndentPt: 0, numberIndentPt: 0 }, labelStyle: 'paren-upper-alpha', restartAfterLevel: -1 },
      ],
    });
    expect(previewLevel(scheme, 0)).toBe('(A)');
  });

  it('returns empty string for an undefined level index', () => {
    const scheme = makeScheme();
    expect(previewLevel(scheme, 99)).toBe('');
  });

  it('commercial preset level 3 previews as "(a)"', () => {
    const commercial = UK_PROFILE.numberingPresets.find(p => p.id === 'preset-uk-sg-commercial')!;
    expect(previewLevel(commercial, 3)).toBe('(a)');
  });

  it('commercial preset level 4 previews as "(i)"', () => {
    const commercial = UK_PROFILE.numberingPresets.find(p => p.id === 'preset-uk-sg-commercial')!;
    expect(previewLevel(commercial, 4)).toBe('(i)');
  });

  it('commercial preset level 5 previews as "(A)"', () => {
    const commercial = UK_PROFILE.numberingPresets.find(p => p.id === 'preset-uk-sg-commercial')!;
    expect(previewLevel(commercial, 5)).toBe('(A)');
  });

  it('recitals preset level 0 previews as "(A)"', () => {
    const recitals = UK_PROFILE.numberingPresets.find(p => p.id === 'preset-recitals')!;
    expect(previewLevel(recitals, 0)).toBe('(A)');
  });

  it('schedule-part preset level 0 previews as "Part 1" (Part grouping)', () => {
    const schedulePart = UK_PROFILE.numberingPresets.find(p => p.id === 'preset-schedule-part')!;
    // Level 0 is "Part %1" → "Part 1". Intentional: %2 at level 1 = level-1 counter.
    expect(previewLevel(schedulePart, 0)).toBe('Part 1');
  });

  it('schedule-part preset level 1 previews as "1." (paragraph counter, not "Part 1 / 1.")', () => {
    const schedulePart = UK_PROFILE.numberingPresets.find(p => p.id === 'preset-schedule-part')!;
    // Level 1 uses "%2." = level-1 counter = "1." — correct for Part-body paragraphs.
    expect(previewLevel(schedulePart, 1)).toBe('1.');
  });
});

// ---------------------------------------------------------------------------
// validateScheme
// ---------------------------------------------------------------------------

describe('validateScheme', () => {
  it('accepts a well-formed scheme', () => {
    const result = validateScheme(makeScheme());
    expect(result.valid).toBe(true);
    expect(result.errors).toHaveLength(0);
  });

  it('rejects an empty name', () => {
    const result = validateScheme(makeScheme({ name: '' }));
    expect(result.valid).toBe(false);
    expect(result.errors.some(e => e.toLowerCase().includes('name'))).toBe(true);
  });

  it('rejects non-contiguous levels', () => {
    const scheme = makeScheme({
      levels: [
        { levelIndex: 0, listNumbering: 'arabic', formatString: '%1.', prefix: '', suffix: '', indent: { textIndentPt: 0, numberIndentPt: 0 }, labelStyle: 'decimal-compound', restartAfterLevel: -1 },
        { levelIndex: 2, listNumbering: 'arabic', formatString: '%3.', prefix: '', suffix: '', indent: { textIndentPt: 72, numberIndentPt: 36 }, labelStyle: 'decimal-compound', restartAfterLevel: 0 },
      ],
    });
    const result = validateScheme(scheme);
    expect(result.valid).toBe(false);
    expect(result.errors.some(e => e.includes('contiguous'))).toBe(true);
  });

  it('rejects a %N token referencing an undefined level', () => {
    const scheme = makeScheme({
      levels: [
        { levelIndex: 0, listNumbering: 'arabic', formatString: '%1.%2.', prefix: '', suffix: '', indent: { textIndentPt: 0, numberIndentPt: 0 }, labelStyle: 'decimal-compound', restartAfterLevel: -1 },
      ],
    });
    const result = validateScheme(scheme);
    expect(result.valid).toBe(false);
    expect(result.errors.some(e => e.includes('level 1'))).toBe(true);
  });

  it('rejects unmatched closing parenthesis', () => {
    const scheme = makeScheme({
      levels: [
        { levelIndex: 0, listNumbering: 'arabic', formatString: '%1)', prefix: '', suffix: '', indent: { textIndentPt: 0, numberIndentPt: 0 }, labelStyle: 'plain', restartAfterLevel: -1 },
      ],
    });
    const result = validateScheme(scheme);
    expect(result.valid).toBe(false);
    expect(result.errors.some(e => e.toLowerCase().includes('parenthesis'))).toBe(true);
  });

  it('rejects a newer schemaVersion than supported', () => {
    const result = validateScheme(makeScheme({ schemaVersion: SCHEMA_VERSION + 1 }));
    expect(result.valid).toBe(false);
    expect(result.errors.some(e => e.includes('schemaVersion'))).toBe(true);
  });

  it('accepts the commercial preset as-is', () => {
    const commercial = UK_PROFILE.numberingPresets.find(p => p.id === 'preset-uk-sg-commercial')!;
    const result = validateScheme(commercial);
    expect(result.valid).toBe(true);
  });

  it('accepts the recitals preset as-is', () => {
    const recitals = UK_PROFILE.numberingPresets.find(p => p.id === 'preset-recitals')!;
    const result = validateScheme(recitals);
    expect(result.valid).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// duplicateScheme
// ---------------------------------------------------------------------------

describe('duplicateScheme', () => {
  it('creates a copy with a new id, builtIn:false, and " (copy)" suffix', () => {
    const original = makeScheme({ id: 'orig', name: 'Original', builtIn: true });
    const copy = duplicateScheme(original, 'new-id');
    expect(copy.id).toBe('new-id');
    expect(copy.builtIn).toBe(false);
    expect(copy.name).toBe('Original (copy)');
  });

  it('preserves all levels from the original', () => {
    const commercial = UK_PROFILE.numberingPresets.find(p => p.id === 'preset-uk-sg-commercial')!;
    const copy = duplicateScheme(commercial, 'custom-id');
    expect(copy.levels.length).toBe(commercial.levels.length);
    expect(copy.levels[3].listNumbering).toBe('lowerLetter');
  });
});

// ---------------------------------------------------------------------------
// boundaryState + clampLevel (level-change — pure)
// ---------------------------------------------------------------------------

describe('boundaryState', () => {
  it('returns false for both when levels is empty', async () => {
    const { boundaryState } = await import('@features/numbering/level-change.js');
    expect(boundaryState([])).toEqual({ canPromote: false, canDemote: false });
  });

  it('cannot promote when all items are at level 0', async () => {
    const { boundaryState } = await import('@features/numbering/level-change.js');
    expect(boundaryState([0, 0])).toEqual({ canPromote: false, canDemote: true });
  });

  it('cannot demote when all items are at level 8', async () => {
    const { boundaryState } = await import('@features/numbering/level-change.js');
    expect(boundaryState([8, 8])).toEqual({ canPromote: true, canDemote: false });
  });

  it('can both promote and demote when levels are mixed', async () => {
    const { boundaryState } = await import('@features/numbering/level-change.js');
    expect(boundaryState([0, 3, 8])).toEqual({ canPromote: true, canDemote: true });
  });
});

describe('clampLevel', () => {
  it('clamps below 0 to 0', async () => {
    const { clampLevel } = await import('@features/numbering/level-change.js');
    expect(clampLevel(-1)).toBe(0);
  });

  it('clamps above 8 to 8', async () => {
    const { clampLevel } = await import('@features/numbering/level-change.js');
    expect(clampLevel(9)).toBe(8);
  });

  it('passes through valid levels unchanged', async () => {
    const { clampLevel } = await import('@features/numbering/level-change.js');
    expect(clampLevel(3)).toBe(3);
    expect(clampLevel(0)).toBe(0);
    expect(clampLevel(8)).toBe(8);
  });
});

// ---------------------------------------------------------------------------
// formatStringToArray (schema helper, used by apply engine)
// ---------------------------------------------------------------------------

describe('formatStringToArray', () => {
  // Word's format array is 0-based: %N (1-based notation) → N-1.
  it('splits compound decimal format into 0-based tokens', async () => {
    const { formatStringToArray } = await import('@features/numbering/schema.js');
    expect(formatStringToArray('%1.%2.%3.')).toEqual([0, '.', 1, '.', 2, '.']);
  });
  it('splits paren-letter format to 0-based index', async () => {
    const { formatStringToArray } = await import('@features/numbering/schema.js');
    expect(formatStringToArray('(%4)')).toEqual(['(', 3, ')']);
  });
  it('handles Part prefix format with 0-based index', async () => {
    const { formatStringToArray } = await import('@features/numbering/schema.js');
    expect(formatStringToArray('Part %1')).toEqual(['Part ', 0]);
  });
});

// ---------------------------------------------------------------------------
// classifyParagraphs (apply engine — pure)
// ---------------------------------------------------------------------------

describe('classifyParagraphs', () => {
  const makeParaNode = (overrides: Record<string, unknown> = {}) => ({
    id: 'p1', index: 0, text: 'text', styleName: 'Normal',
    container: 'body',
    ...overrides,
  });

  it('classifies a footnote paragraph as footnote', async () => {
    const { classifyParagraphs } = await import('@features/numbering/apply.js');
    const para = makeParaNode({ container: 'footnote' }) as import('@shared/doc-model/types.js').ParaNode;
    const result = classifyParagraphs([para], []);
    expect(result[0]?.kind).toBe('footnote');
  });

  it('classifies a table-cell paragraph as table', async () => {
    const { classifyParagraphs } = await import('@features/numbering/apply.js');
    const para = makeParaNode({ container: 'tableCell' }) as import('@shared/doc-model/types.js').ParaNode;
    const result = classifyParagraphs([para], []);
    expect(result[0]?.kind).toBe('table');
  });

  it('classifies a static-numbered paragraph as static (ADR-0006 — no conversion)', async () => {
    const { classifyParagraphs } = await import('@features/numbering/apply.js');
    const para = makeParaNode({ id: 'p-static' }) as import('@shared/doc-model/types.js').ParaNode;
    const staticFindings = [{
      id: 'f1', category: 'static-numbering/manually-typed', severity: 'info' as const,
      title: 'Static', location: { paraId: 'p-static', label: '1.' },
    }] as import('@shared/findings/types.js').Finding[];
    const result = classifyParagraphs([para], staticFindings);
    expect(result[0]?.kind).toBe('static');
  });

  it('classifies a normal body paragraph as convertible', async () => {
    const { classifyParagraphs } = await import('@features/numbering/apply.js');
    const para = makeParaNode() as import('@shared/doc-model/types.js').ParaNode;
    const result = classifyParagraphs([para], []);
    expect(result[0]?.kind).toBe('convertible');
  });

  it('never converts a static paragraph even if it is in the body', async () => {
    const { classifyParagraphs } = await import('@features/numbering/apply.js');
    const para = makeParaNode({ id: 'p-static', container: 'body' }) as import('@shared/doc-model/types.js').ParaNode;
    const staticFindings = [{
      id: 'f1', category: 'static-numbering/manually-typed', severity: 'info' as const,
      title: 'Static', location: { paraId: 'p-static', label: '1.' },
    }] as import('@shared/findings/types.js').Finding[];
    const result = classifyParagraphs([para], staticFindings);
    expect(result[0]?.kind).toBe('static');
  });
});

// ---------------------------------------------------------------------------
// buildReviewPatterns (review — pure, config-driven)
// ---------------------------------------------------------------------------

describe('buildReviewPatterns', () => {
  it('generates clause patterns from config primary noun', async () => {
    const { buildReviewPatterns } = await import('@features/numbering/review.js');
    const patterns = buildReviewPatterns(UK_PROFILE);
    expect(patterns.some(p => p.includes('clause'))).toBe(true);
  });

  it('generates Schedule pattern', async () => {
    const { buildReviewPatterns } = await import('@features/numbering/review.js');
    const patterns = buildReviewPatterns(UK_PROFILE);
    expect(patterns.some(p => p.includes('Schedule'))).toBe(true);
  });

  it('generates paragraph sub-unit pattern', async () => {
    const { buildReviewPatterns } = await import('@features/numbering/review.js');
    const patterns = buildReviewPatterns(UK_PROFILE);
    expect(patterns.some(p => p.startsWith('paragraph'))).toBe(true);
  });

  it('includes standalone lettered limb pattern', async () => {
    const { buildReviewPatterns } = await import('@features/numbering/review.js');
    const patterns = buildReviewPatterns(UK_PROFILE);
    expect(patterns.some(p => p.includes('\\([a-z]\\)'))).toBe(true);
  });

  it('includes recital pattern (case-insensitive wildcard [Rr]ecital)', async () => {
    const { buildReviewPatterns } = await import('@features/numbering/review.js');
    const patterns = buildReviewPatterns(UK_PROFILE);
    // Pattern uses [Rr]ecital — check for 'ecital' which is always present.
    expect(patterns.some(p => p.includes('ecital'))).toBe(true);
  });

  it('includes Part pattern (Part is in otherDivisions)', async () => {
    const { buildReviewPatterns } = await import('@features/numbering/review.js');
    const patterns = buildReviewPatterns(UK_PROFILE);
    expect(patterns.some(p => p.includes('Part'))).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// exportSchemes / importSchemes (scheme-store — T6)
// ---------------------------------------------------------------------------

/** FileReader-based text helper for jsdom compat (blob.text() not available in jsdom). */
function readBlobText(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = () => reject(reader.error);
    reader.readAsText(blob);
  });
}

describe('exportSchemes / importSchemes round-trip', () => {
  it('export produces valid JSON with schemaVersion and schemes array', async () => {
    const { exportSchemes } = await import('@features/numbering/scheme-store.js');
    const commercial = UK_PROFILE.numberingPresets.find(p => p.id === 'preset-uk-sg-commercial')!;
    const blob = exportSchemes([commercial]);
    const text = await readBlobText(blob);
    const parsed = JSON.parse(text);
    expect(parsed.schemaVersion).toBe(SCHEMA_VERSION);
    expect(Array.isArray(parsed.schemes)).toBe(true);
    expect(parsed.schemes[0].id).toBe('preset-uk-sg-commercial');
  });

  it('exported blob contains no document content — only numbering definitions', async () => {
    const { exportSchemes } = await import('@features/numbering/scheme-store.js');
    const commercial = UK_PROFILE.numberingPresets.find(p => p.id === 'preset-uk-sg-commercial')!;
    const blob = exportSchemes([commercial]);
    const text = await readBlobText(blob);
    // Must not contain document-specific content markers.
    expect(text).not.toContain('paragraphText');
    expect(text).not.toContain('docContent');
    // Must contain scheme-specific fields.
    expect(text).toContain('levels');
    expect(text).toContain('formatString');
  });

  it('import round-trips: exported scheme imports back with correct levels', async () => {
    const { exportSchemes, importSchemes } = await import('@features/numbering/scheme-store.js');
    const commercial = UK_PROFILE.numberingPresets.find(p => p.id === 'preset-uk-sg-commercial')!;
    const blob = exportSchemes([commercial]);
    // Import — built-in id collision → fresh id assigned.
    const result = await importSchemes(blob, 'replace');
    expect(result.errors).toHaveLength(0);
    expect(result.imported).toHaveLength(1);
    expect(result.imported[0]?.levels.length).toBe(9);
  });

  it('import rejects malformed JSON', async () => {
    const { importSchemes } = await import('@features/numbering/scheme-store.js');
    const bad = new Blob(['{not valid json}'], { type: 'application/json' });
    const result = await importSchemes(bad);
    expect(result.errors.length).toBeGreaterThan(0);
    expect(result.imported).toHaveLength(0);
  });

  it('import rejects a newer schemaVersion', async () => {
    const { importSchemes } = await import('@features/numbering/scheme-store.js');
    const future = JSON.stringify({ schemaVersion: 99, schemes: [] });
    const blob = new Blob([future], { type: 'application/json' });
    const result = await importSchemes(blob);
    expect(result.errors.length).toBeGreaterThan(0);
  });

  it('user-library collision with keepBoth → fresh id, original preserved', async () => {
    const { exportSchemes, importSchemes, saveScheme, loadScheme } = await import('@features/numbering/scheme-store.js');
    // Create a user scheme (non-built-in id) and save it.
    const userScheme = {
      id: 'user-test-001',
      name: 'My Firm Style',
      jurisdiction: 'SG-UK' as const,
      schemaVersion: 1,
      suppressForSgUk: false,
      levels: [
        { levelIndex: 0, listNumbering: 'arabic' as const, formatString: '%1.',
          prefix: '', suffix: '', indent: { textIndentPt: 0, numberIndentPt: 0 },
          labelStyle: 'decimal-compound' as const, restartAfterLevel: -1 },
      ],
    };
    await saveScheme(userScheme);
    // Export it and re-import with keepBoth.
    const blob = exportSchemes([userScheme]);
    const result = await importSchemes(blob, 'keepBoth');
    expect(result.errors).toHaveLength(0);
    // keepBoth → fresh id, NOT the original 'user-test-001'.
    expect(result.imported.length).toBeGreaterThan(0);
    expect(result.imported[0]?.id).not.toBe('user-test-001');
    // Original still loadable.
    const original = await loadScheme('user-test-001');
    expect(original).toBeDefined();
    expect(original?.name).toBe('My Firm Style');
  });

  it('built-in id collision gets a fresh id on import', async () => {
    const { exportSchemes, importSchemes } = await import('@features/numbering/scheme-store.js');
    const commercial = UK_PROFILE.numberingPresets.find(p => p.id === 'preset-uk-sg-commercial')!;
    const blob = exportSchemes([commercial]);
    const result = await importSchemes(blob, 'replace');
    // The imported scheme must NOT use the reserved built-in id.
    expect(result.imported[0]?.id).not.toBe('preset-uk-sg-commercial');
    expect(result.imported[0]?.builtIn).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// Warnings (T8 — edge cases)
// ---------------------------------------------------------------------------

describe('warnings', () => {
  it('wordOnlineBanner returns empty string when no degraded levels', async () => {
    const { wordOnlineBanner } = await import('@features/numbering/warnings.js');
    expect(wordOnlineBanner([])).toBe('');
  });

  it('wordOnlineBanner lists degraded levels', async () => {
    const { wordOnlineBanner } = await import('@features/numbering/warnings.js');
    const banner = wordOnlineBanner([0, 1]);
    expect(banner).toContain('0');
    expect(banner).toContain('1');
  });

  it('staticSkippedMessage includes count', async () => {
    const { staticSkippedMessage } = await import('@features/numbering/warnings.js');
    const msg = staticSkippedMessage(3);
    expect(msg).toContain('3');
  });

  it('hasMultipleSectionKinds returns false for single-kind index', async () => {
    const { hasMultipleSectionKinds } = await import('@features/numbering/warnings.js');
    const fakeIndex = {
      nodes: [
        { kind: 'clause' }, { kind: 'clause' }, { kind: 'clause' },
      ],
    };
    expect(hasMultipleSectionKinds(fakeIndex as unknown as import('@shared/section-index/index.js').SectionIndex)).toBe(false);
  });

  it('hasMultipleSectionKinds returns true for mixed kinds', async () => {
    const { hasMultipleSectionKinds } = await import('@features/numbering/warnings.js');
    const fakeIndex = {
      nodes: [
        { kind: 'clause' }, { kind: 'schedule' },
      ],
    };
    expect(hasMultipleSectionKinds(fakeIndex as unknown as import('@shared/section-index/index.js').SectionIndex)).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// BUILTIN_IDS
// ---------------------------------------------------------------------------

describe('BUILTIN_IDS', () => {
  it('contains all shipped preset ids', () => {
    expect(BUILTIN_IDS.has('preset-uk-sg-commercial')).toBe(true);
    expect(BUILTIN_IDS.has('preset-schedule')).toBe(true);
    expect(BUILTIN_IDS.has('preset-recitals')).toBe(true);
    expect(BUILTIN_IDS.has('preset-plain-outline')).toBe(true);
    expect(BUILTIN_IDS.has('preset-us-litigation')).toBe(true);
  });
});
