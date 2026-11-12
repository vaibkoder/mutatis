/**
 * host-web — end-to-end acceptance.
 *
 * The phase-2 goal in one test: a `.docx` parsed by the web loader, fed to the
 * UNCHANGED `runFactualConsistency` checker, produces the SAME findings the add-in
 * path produces — proving "same checker, same findings, no Word". Covers all three
 * sub-checks (weekend date, words-vs-figures amount mismatch, unfilled placeholder).
 */
import { describe, it, expect } from 'vitest';
import { synthesiseDocx } from '@features/compare/io/text/synthesise-docx.js';
import { SG_PROFILE } from '@shared/config/index.js';
import { runFactualConsistency } from '@features/factual-consistency/index.js';
import { loadDocSnapshotFromOoxml } from '@host-web/ooxml-loader.js';

describe('host-web factual-consistency end-to-end', () => {
  it('web-loaded .docx yields the expected factual-consistency findings', () => {
    // One <w:p> per line (synthesiseDocx) — a realistic mini-document.
    const docxText = [
      'This Agreement is executed on 6 June 2026.', // 6 June 2026 = Saturday → weekend execution
      'The fee is Ten Thousand Dollars (S$1,000).', // words (10,000) ≠ figure (1,000) → mismatch
      'The completion date is [●].', // unfilled placeholder
      'This clause is entirely unremarkable.', // no finding
    ].join('\n');

    const snapshot = loadDocSnapshotFromOoxml(synthesiseDocx(docxText));
    const paras = [...snapshot.paragraphs, ...snapshot.footnotes];
    const findings = runFactualConsistency(paras, SG_PROFILE);

    const categories = findings.map(f => f.category);
    expect(categories).toContain('factual-consistency/weekend-date');
    expect(categories).toContain('factual-consistency/amount-mismatch');
    expect(categories).toContain('factual-consistency/placeholder');

    // Findings carry host-neutral locations anchored to the loader's paragraph ids.
    const weekend = findings.find(f => f.category === 'factual-consistency/weekend-date');
    expect(weekend?.location.paraId).toBe('body-para-0');
    expect(weekend?.title).toMatch(/Saturday/);
  });

  it('a clean document produces no findings', () => {
    const snapshot = loadDocSnapshotFromOoxml(
      synthesiseDocx('This is a perfectly ordinary clause.\nAnd another one.'),
    );
    const findings = runFactualConsistency([...snapshot.paragraphs], SG_PROFILE);
    expect(findings).toHaveLength(0);
  });
});
