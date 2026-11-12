/**
 * Feature 12 (Compare — additional sources) T2: the `text` CompareSource kind.
 * Spec: spec/features/10-compare-additional-sources.md §4.1 · ADR-0018.
 *
 * Verifies a `{ kind:'text' }` source resolves to a DocxBundle via loadBothSides,
 * and that both sides load together (the input to the unchanged pipeline).
 */
import { describe, it, expect } from 'vitest';

import { loadBothSides } from '@features/compare/io/load.js';
import type { CompareDirection } from '@features/compare/io/load.js';

describe('T2 — text CompareSource kind', () => {
  it('resolves a text source to a DocxBundle with a document part', async () => {
    const direction: CompareDirection = {
      original: { kind: 'text', text: 'First.\nSecond.' },
      revised: { kind: 'text', text: 'First.\nSecond changed.' },
    };
    const { originalBundle, revisedBundle } = await loadBothSides(direction);

    expect(originalBundle.part('word/document.xml')).toContain('First.');
    expect(revisedBundle.part('word/document.xml')).toContain('Second changed.');
    // Both sides carry the styles part the redline needs.
    expect(originalBundle.hasPart('word/styles.xml')).toBe(true);
    expect(revisedBundle.hasPart('word/styles.xml')).toBe(true);
  });

  it('an empty-string text source still yields a valid (empty-body) bundle', async () => {
    const { originalBundle } = await loadBothSides({
      original: { kind: 'text', text: '' },
      revised: { kind: 'text', text: '' },
    });
    expect(originalBundle.part('word/document.xml')).toContain('<w:body>');
  });
});
