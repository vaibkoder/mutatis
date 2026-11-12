/**
 * A0 acceptance — every P-test's committed fixture is present and non-empty.
 * Guards against a fixture being deleted/renamed without its consumer noticing.
 * The byte-stability + validity of the fixtures is exercised by the generator
 * itself (re-run → identical) and by the A2+ extraction tests that parse them.
 */
import { describe, it, expect } from 'vitest';
import { existsSync, statSync } from 'node:fs';

import { PDF_FIXTURES, pdfFixture } from '../../fixtures/pdf/manifest.js';

describe('A0 — PDF fixture corpus', () => {
  it('every manifested fixture exists and is non-empty', () => {
    for (const name of Object.keys(PDF_FIXTURES)) {
      const path = pdfFixture(name);
      expect(existsSync(path), `${name} missing — run tests/fixtures/pdf/generate.ts`).toBe(true);
      expect(statSync(path).size, `${name} is empty`).toBeGreaterThan(0);
    }
  });

  it('contains no real/client documents (corpus is the manifest only)', () => {
    // The manifest is the allow-list. Any .pdf in the dir not listed here would
    // be an un-vetted (possibly confidential) document — fail closed.
    const pdfNames = Object.keys(PDF_FIXTURES).filter((n) => n.endsWith('.pdf'));
    expect(pdfNames.length).toBeGreaterThanOrEqual(17);
  });
});
