/**
 * Feature 12 (Compare — additional sources) T3: text-vs-text driver.
 * Spec: spec/features/10-compare-additional-sources.md §4.1 · ADR-0018.
 *
 * Covers the eager-generate / click-to-open contract (TT5), the empty no-op
 * (TT6), and the oversize typed limit error (TT7). `materialise` reaches the host
 * via getActiveHost().writeBack (12f); we install a stub host and spy on its
 * writeBack to prove it is NOT called during Compare and IS called by openRedline().
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { setActiveHost, _resetActiveHost } from '@shared/host/index.js';
import type { HostPort } from '@shared/host/index.js';

import {
  textCompareDriverState as d,
  checkPasteSize,
  TextPasteTooLargeError,
  MAX_TEXT_PASTE_CHARS,
} from '@features/compare/driver.js';

/** Stub host whose spied writeBack stands in for materialise's host write. */
let writeBack: ReturnType<typeof vi.fn>;

beforeEach(() => {
  // Reset the shared module-level signals between tests.
  d.status.value = 'idle';
  d.progress.value = 0;
  d.errorMessage.value = null;
  d.ops.value = null;
  d.summary.value = null;
  d.originalText.value = '';
  d.revisedText.value = '';
  d.redlineReady.value = false;
  writeBack = vi.fn().mockResolvedValue(undefined);
  setActiveHost({ writeBack } as unknown as HostPort);
});

afterEach(() => {
  _resetActiveHost();
  vi.restoreAllMocks();
});

describe('checkPasteSize — typed limit error (TT7 unit)', () => {
  it('throws TextPasteTooLargeError above the cap', () => {
    const big = 'x'.repeat(MAX_TEXT_PASTE_CHARS + 1);
    expect(() => checkPasteSize('original', big)).toThrow(TextPasteTooLargeError);
  });

  it('does not throw at or below the cap', () => {
    expect(() => checkPasteSize('revised', 'x'.repeat(10))).not.toThrow();
  });
});

describe('runCompare — eager generate, click to open (TT5)', () => {
  it('produces ops + summary + a ready redline WITHOUT calling writeBack', async () => {
    const spy = writeBack;
    d.originalText.value = 'The Seller shall deliver the Goods.';
    d.revisedText.value  = 'The Seller may deliver the Goods.';

    await d.runCompare();

    expect(d.status.value).toBe('done');
    expect(d.ops.value).not.toBeNull();
    expect(d.summary.value).not.toBeNull();
    expect(d.redlineReady.value).toBe(true);
    // Eager-generate, click-to-open: the redline is NOT materialised on Compare.
    expect(spy).not.toHaveBeenCalled();
  });

  it('openRedline() materialises the held redline (newDocument) only when called', async () => {
    const spy = writeBack;
    d.originalText.value = 'Clause 1.\nClause 2.';
    d.revisedText.value  = 'Clause 1.\nClause 2 amended.';

    await d.runCompare();
    expect(spy).not.toHaveBeenCalled();

    await d.openRedline();
    expect(spy).toHaveBeenCalledTimes(1);
    const [, optsArg] = spy.mock.calls[0]!;
    expect(optsArg).toEqual({ mode: 'newDocument' });
  });

  it('openRedline() is a no-op when no redline has been generated', async () => {
    const spy = writeBack;
    await d.openRedline();
    expect(spy).not.toHaveBeenCalled();
  });
});

describe('runCompare — empty input no-op (TT6)', () => {
  it('both sides blank → friendly message, no crash, no redline', async () => {
    const spy = writeBack;
    d.originalText.value = '   ';
    d.revisedText.value  = '\n\t ';

    await d.runCompare();

    expect(d.status.value).toBe('error');
    expect(d.errorMessage.value).toMatch(/paste text/i);
    expect(d.redlineReady.value).toBe(false);
    expect(spy).not.toHaveBeenCalled();
  });
});

describe('runCompare — oversize input (TT7)', () => {
  it('a paste beyond the cap surfaces a friendly limit message, nothing generated', async () => {
    d.originalText.value = 'x'.repeat(MAX_TEXT_PASTE_CHARS + 1);
    d.revisedText.value  = 'y';

    await d.runCompare();

    expect(d.status.value).toBe('error');
    expect(d.errorMessage.value).toMatch(/too large/i);
    expect(d.redlineReady.value).toBe(false);
  });
});
