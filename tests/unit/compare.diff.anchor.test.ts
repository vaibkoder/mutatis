/**
 * Feature 08 — Compare §5.3 · edit-boundary anchoring (ADR-0028).
 *
 * `diff_main` returns *a* minimal edit script, not the most readable one, and
 * the word diff runs it in sentinel space where DMP's own boundary-shifting
 * pass is blind. These tests pin the readable placement, and pin the two
 * invariants the pass must never break: atomic tokens stay whole (§5.3) and
 * nothing is ever added to or dropped from either side (ADR-0008).
 */
import { describe, it, expect } from 'vitest';

import { getCompareConfig, SG_PROFILE, UK_PROFILE } from '@shared/config/index.js';
import { wordDiff, slideEditBoundaries } from '@features/compare/diff/index.js';
import type { WordDiff } from '@features/compare/model.js';

const SG = getCompareConfig(SG_PROFILE);
const UK = getCompareConfig(UK_PROFILE);

/** Rebuild one side's text from the diff — what that side must still say. */
function side(diffs: WordDiff[], which: 'ins' | 'del'): string {
  return diffs.filter((d) => d.op === 'eq' || d.op === which).map((d) => d.text).join('');
}

describe('edit-boundary anchoring — readable placement', () => {
  it('anchors an appended proviso AFTER the clause it follows, not before it', () => {
    // Both sides end "before Completion.", so the insert is equally minimal
    // anchored on either occurrence. Only the later one leaves the original
    // sentence readable and unbroken.
    const original =
      'The Vendor shall indemnify the Purchaser against all Tax Liabilities arising in respect of any period ending on or before Completion.';
    const revised =
      'The Vendor shall indemnify the Purchaser against all Tax Liabilities arising in respect of any period ending on or before Completion, provided that the Vendor shall have no liability to the extent that the Tax Liability has been paid before Completion.';

    const diffs = wordDiff(original, revised, SG);
    const first = diffs[0]!;
    const inserted = diffs.filter((d) => d.op === 'ins');

    // The untouched clause runs unbroken, all the way to its own "Completion".
    expect(first.op).toBe('eq');
    expect(first.text.endsWith('before Completion')).toBe(true);
    // The insert is one span that opens on the comma joining the proviso.
    expect(inserted).toHaveLength(1);
    expect(inserted[0]!.text.startsWith(', provided that')).toBe(true);
    // …and the reader never meets "Completion" twice in a row.
    expect(first.text.endsWith('before Completion, provided')).toBe(false);
  });

  it('anchors an appended list limb after the limb it follows', () => {
    const diffs = wordDiff(
      'indemnify the Purchaser against all Tax Liabilities of the Company.',
      'indemnify the Purchaser against all Tax Liabilities of the Company and all Environmental Liabilities of the Company.',
      UK,
    );
    const inserted = diffs.filter((d) => d.op === 'ins');
    expect(inserted).toHaveLength(1);
    expect(inserted[0]!.text.trimStart().startsWith('and all Environmental')).toBe(true);
    expect(diffs[0]!.text.endsWith('of the Company')).toBe(true);
  });

  it('leaves an already well-placed deletion where it is', () => {
    const diffs = wordDiff(
      'except in the case of fraud, gross negligence or wilful misconduct.',
      'except in the case of fraud, negligence or wilful misconduct.',
      UK,
    );
    const del = diffs.filter((d) => d.op === 'del');
    expect(del).toHaveLength(1);
    expect(del[0]!.text).toBe('gross ');
  });

  it('leaves a two-sided replace alone (no free boundary to slide)', () => {
    const diffs = wordDiff('The Seller shall pay', 'The Seller may pay', UK);
    expect(diffs.filter((d) => d.op === 'del').map((d) => d.text)).toEqual(['shall']);
    expect(diffs.filter((d) => d.op === 'ins').map((d) => d.text)).toEqual(['may']);
  });
});

describe('edit-boundary anchoring — invariants', () => {
  it('never splits an atomic amount to move a boundary', () => {
    const original = 'the Vendor shall pay S$1,250,000 on Completion.';
    const revised = 'the Vendor shall pay S$1,250,000 and a further S$1,250,000 on Completion.';
    const diffs = wordDiff(original, revised, SG);

    // The amount survives whole on both sides, with its atomic kind intact…
    const amounts = diffs.filter((d) => d.text === 'S$1,250,000');
    expect(amounts.length).toBeGreaterThan(0);
    for (const a of amounts) expect(a.atomic).toBe('amount');
    // …and no digit fragment ever leaks out of it.
    for (const d of diffs) expect(/^[\d,]+$/.test(d.text)).toBe(false);
  });

  it('is length-preserving: both sides still say exactly what they said', () => {
    const pairs: [string, string][] = [
      ['ending on or before Completion.',
       'ending on or before Completion, provided that it is paid before Completion.'],
      ['all Tax Liabilities of the Company.',
       'all Tax Liabilities of the Company and all Losses of the Company.'],
      ['No set-off may be made under this clause 9.',
       'No set-off or deduction may be made under this clause 9.'],
      ['pay £40,000 now', 'pay £50,000 now'],
      ['The limitations in clause 9.13 (other than for fraud) apply.',
       'The limitations in clause 9.13 apply.'],
    ];
    for (const [original, revised] of pairs) {
      const diffs = wordDiff(original, revised, UK);
      expect(side(diffs, 'del')).toBe(original);
      expect(side(diffs, 'ins')).toBe(revised);
    }
  });

  it('slideEditBoundaries returns its input untouched when there is nothing to slide', () => {
    const diffs: WordDiff[] = [
      { op: 'eq', text: 'a ' },
      { op: 'ins', text: 'brand new ' },
      { op: 'eq', text: 'clause' },
    ];
    expect(slideEditBoundaries(diffs, UK)).toEqual(diffs);
  });

  it('slideEditBoundaries never drops a token (ADR-0008 never-suppress)', () => {
    const original = 'ending on or before Completion.';
    const revised = 'ending on or before Completion, and in any event before Completion.';
    const diffs = wordDiff(original, revised, UK);
    // Every word of the inserted proviso is still present in the redline.
    for (const word of ['and', 'in', 'any', 'event', 'before']) {
      expect(diffs.some((d) => d.op === 'ins' && d.text.includes(word))).toBe(true);
    }
  });
});
