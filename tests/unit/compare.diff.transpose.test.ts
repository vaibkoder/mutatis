/**
 * Feature 08 — Compare §5.3 · adjacent transposition (ADR-0031).
 *
 * Myers/LCS diffing maximises retained common subsequence, so when two adjacent
 * words trade places it keeps the separator between them and reports two
 * independent substitutions — leaving the reader to notice unaided that the two
 * struck words are the two inserted words in the other order. These tests pin
 * the relocation shape, the guards that keep it narrow, and the invariant it
 * may never break: nothing leaves the redline (ADR-0008).
 */
import { describe, it, expect } from 'vitest';

import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import { wordDiff, detectTransposition } from '@features/compare/diff/index.js';
import type { WordDiff } from '@features/compare/model.js';

const SG = getCompareConfig(SG_PROFILE);

function render(diffs: readonly WordDiff[]): string {
  return diffs
    .map((d) => (d.op === 'eq' ? d.text : d.op === 'del' ? `[-${d.text}-]` : `{+${d.text}+}`))
    .join('');
}

function side(diffs: readonly WordDiff[], which: 'ins' | 'del'): string {
  return diffs
    .filter((d) => d.op === 'eq' || d.op === which)
    .map((d) => d.text)
    .join('');
}

/** The shape DMP produces for a swap, as the pass receives it. */
function quintuple(a: string, b: string, mid: string): WordDiff[] {
  return [
    { op: 'del', text: a },
    { op: 'ins', text: b },
    { op: 'eq', text: mid },
    { op: 'del', text: b },
    { op: 'ins', text: a },
  ];
}

describe('detectTransposition (§5.3, ADR-0031)', () => {
  it('reads a word swap as one relocation, not two replacements', () => {
    const out = wordDiff(
      'The Vendor shall indemnify, defend and hold harmless the Purchaser.',
      'The Vendor shall defend, indemnify and hold harmless the Purchaser.',
      SG,
    );
    expect(render(out)).toBe(
      'The Vendor shall {+defend, +}indemnify[-, defend-] and hold harmless the Purchaser.',
    );
  });

  it('leaves the surviving span unmarked', () => {
    const out = wordDiff('shall indemnify, defend and hold', 'shall defend, indemnify and hold', SG);
    expect(out.filter((d) => d.op === 'eq').map((d) => d.text)).toContain('indemnify');
  });

  it('never adds to or drops from either side (ADR-0008)', () => {
    const pairs: Array<[string, string]> = [
      ['shall indemnify, defend and hold', 'shall defend, indemnify and hold'],
      ['pay S$1,000 or S$2,000 on demand', 'pay S$2,000 or S$1,000 on demand'],
      ['the Vendor and the Purchaser agree', 'the Purchaser and the Vendor agree'],
    ];
    for (const [a, b] of pairs) {
      const out = wordDiff(a, b, SG);
      expect(side(out, 'del')).toBe(a);
      expect(side(out, 'ins')).toBe(b);
    }
  });

  it('carries each span its atomic kind across the relocation', () => {
    // A swap of two amounts must still classify as a numerical-value change.
    const out = detectTransposition(
      [
        { op: 'del', text: 'S$1,000', atomic: 'amount' },
        { op: 'ins', text: 'S$2,000', atomic: 'amount' },
        { op: 'eq', text: ' / ' },
        { op: 'del', text: 'S$2,000', atomic: 'amount' },
        { op: 'ins', text: 'S$1,000', atomic: 'amount' },
      ],
      SG,
    );
    expect(out.find((d) => d.op === 'ins' && d.text === 'S$2,000')?.atomic).toBe('amount');
    expect(out.find((d) => d.op === 'del' && d.text === 'S$2,000')?.atomic).toBe('amount');
    expect(out.find((d) => d.op === 'eq' && d.text === 'S$1,000')?.atomic).toBe('amount');
  });

  it('declines a swap of two long phrases', () => {
    // Long spans already read fine as two clean phrase-for-phrase replacements;
    // re-expressing them as a relocation moves a whole clause across the page.
    const long = 'all Losses arising out of any breach of the Warranties';
    const other = 'all Losses arising from any misrepresentation by the Vendor';
    const input = quintuple(long, other, ', ');
    expect(detectTransposition(input, SG)).toEqual(input);
  });

  it('declines a swap separated by more than a short connective', () => {
    // Punctuation-only, so `isNoise` passes and MAX_MID_CHARS is the sole reason
    // this declines: a dotted leader puts the two spans too far apart to read as
    // one swap.
    const input = quintuple('indemnify', 'defend', ' ................ ');
    expect(detectTransposition(input, SG)).toEqual(input);
  });

  it('declines a connective carrying a word — never strike one to re-insert it', () => {
    // The rewrite duplicates the connective (inserted with B, struck with B's old
    // copy). Duplicating `, ` is invisible; duplicating ` and the ` makes the
    // reader prove two marks identical, which is exactly the P2 failure
    // `refineReplaceBlocks` refuses to commit (ADR-0030 rule 2). The two ADRs
    // have to agree, so this shape stays two clean one-span replacements.
    const input = quintuple('the Vendor', 'the Purchaser', ' and the ');
    expect(detectTransposition(input, SG)).toEqual(input);
    expect(
      render(wordDiff('by the Vendor and the Purchaser here', 'by the Purchaser and the Vendor here', SG)),
    ).toBe('by the [-Vendor-]{+Purchaser+} and the [-Purchaser-]{+Vendor+} here');
  });

  it('declines a window that is only part of a longer deletion', () => {
    // `coalesce` leaves adjacent same-op entries in place wherever one of them
    // is atomic, so a five-entry window can be a slice through the middle of a
    // longer span. Claiming "A moved" for part of what moved would misreport it.
    const inner = quintuple('indemnify', 'defend', ', ');
    const sliced: WordDiff[] = [{ op: 'del', text: 'shall ' }, ...inner];
    expect(detectTransposition(sliced, SG)).toEqual(sliced);
    const trailing: WordDiff[] = [...inner, { op: 'ins', text: ' promptly' }];
    expect(detectTransposition(trailing, SG)).toEqual(trailing);
  });

  it('declines a same-text "swap" and a non-equal middle', () => {
    const same = quintuple('indemnify', 'indemnify', ', ');
    expect(detectTransposition(same, SG)).toEqual(same);
    const notEq: WordDiff[] = [
      { op: 'del', text: 'indemnify' },
      { op: 'ins', text: 'defend' },
      { op: 'del', text: ', ' },
      { op: 'del', text: 'defend' },
      { op: 'ins', text: 'indemnify' },
    ];
    expect(detectTransposition(notEq, SG)).toEqual(notEq);
  });

  it('leaves an ordinary replace untouched', () => {
    const out = wordDiff('the Vendor shall pay', 'the Vendor must pay', SG);
    expect(render(out)).toBe('the Vendor [-shall-]{+must+} pay');
  });
});
