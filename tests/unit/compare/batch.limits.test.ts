/**
 * Compare at scale — input caps (ADR-0037; spec §4.4; acceptance B8).
 *
 * A batch multiplies every input exposure by N, and the `.docx` path had no
 * size cap at all before this. The load-bearing property is not that a cap
 * exists but that it REFUSES: silently keeping the first 100 of 140 variants
 * would hand the user a complete-looking batch that is quietly missing
 * documents, which is suppression wearing a different hat (ADR-0008).
 */
import { describe, it, expect } from 'vitest';

import {
  MAX_BATCH_VARIANTS,
  MAX_BUNDLE_DOCUMENTS,
  MAX_DOCX_BYTES,
  MAX_BATCH_TOTAL_BYTES,
  MAX_BATCH_ARTEFACT_BYTES,
  BatchInputTooLargeError,
  checkVariantCap,
  checkBundleCap,
  checkDocxByteCap,
  checkTotalByteCap,
  artefactBudgetAllows,
} from '@features/compare/batch/limits.js';
import { resolvePoolSize, BATCH_POOL_MAX } from '@features/compare/batch/constants.js';

describe('batch caps refuse rather than truncate', () => {
  it('accepts a batch exactly at each cap', () => {
    expect(() => checkVariantCap(MAX_BATCH_VARIANTS)).not.toThrow();
    expect(() => checkBundleCap(MAX_BUNDLE_DOCUMENTS)).not.toThrow();
    expect(() => checkDocxByteCap(MAX_DOCX_BYTES, 'spa.docx')).not.toThrow();
    expect(() => checkTotalByteCap(MAX_BATCH_TOTAL_BYTES)).not.toThrow();
  });

  it('throws a typed error naming the cap and the limit', () => {
    expect(() => checkVariantCap(MAX_BATCH_VARIANTS + 1)).toThrow(BatchInputTooLargeError);
    try {
      checkVariantCap(140);
    } catch (err) {
      const e = err as BatchInputTooLargeError;
      expect(e.cap).toBe('variants');
      expect(e.actual).toBe(140);
      expect(e.limit).toBe(MAX_BATCH_VARIANTS);
      expect(e.message).toContain('no variants were dropped to fit');
    }
  });

  it('names the offending document for a per-document cap', () => {
    try {
      checkDocxByteCap(MAX_DOCX_BYTES + 1, 'facility-agreement-v7.docx');
    } catch (err) {
      const e = err as BatchInputTooLargeError;
      expect(e.cap).toBe('docxBytes');
      expect(e.documentLabel).toBe('facility-agreement-v7.docx');
      expect(e.message).toContain('facility-agreement-v7.docx');
    }
  });

  it('covers every cap kind', () => {
    expect(() => checkBundleCap(MAX_BUNDLE_DOCUMENTS + 1)).toThrow(/documents cap/);
    expect(() => checkTotalByteCap(MAX_BATCH_TOTAL_BYTES + 1)).toThrow(/totalBytes cap/);
  });
});

describe('the artefact budget degrades instead of refusing', () => {
  it('allows an artefact that fits', () => {
    expect(artefactBudgetAllows(0, 1024)).toBe(true);
    expect(artefactBudgetAllows(MAX_BATCH_ARTEFACT_BYTES - 10, 10)).toBe(true);
  });

  it('declines one that does not, without throwing', () => {
    expect(artefactBudgetAllows(MAX_BATCH_ARTEFACT_BYTES, 1)).toBe(false);
  });
});

describe('resolvePoolSize', () => {
  it('leaves a core for the UI thread', () => {
    expect(resolvePoolSize(4)).toBe(3);
    expect(resolvePoolSize(2)).toBe(1);
  });

  it('never drops below one worker', () => {
    expect(resolvePoolSize(1)).toBe(1);
    expect(resolvePoolSize(0)).toBe(1);
  });

  it('caps at BATCH_POOL_MAX however many cores are reported', () => {
    expect(resolvePoolSize(64)).toBe(BATCH_POOL_MAX);
  });

  it('falls back sanely when hardwareConcurrency is unavailable', () => {
    expect(resolvePoolSize(undefined)).toBe(1);
  });

  it('honours an explicit override, still bounded', () => {
    expect(resolvePoolSize(64, 2)).toBe(2);
    expect(resolvePoolSize(2, 64)).toBe(BATCH_POOL_MAX);
  });
});
