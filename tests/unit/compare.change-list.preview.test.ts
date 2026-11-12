/**
 * §8.3's change-list row preview, and the ONE consequence of ADR-0043 that is
 * not purely presentational.
 *
 * `previewText` walks `wordDiffs` in stream order across BOTH sides and then
 * truncates from the head at 80 characters. Stage 6.5 puts the insertion first,
 * so a long replace's row now leads with the inserted text and can lose the
 * struck text off the end — where before it led with the struck text and lost
 * the insertion.
 *
 * This path had NO test when 08d landed the flip. The plan's task 5 recorded
 * "both UI panels passed unchanged under the flip", which measured that no test
 * broke; nothing covered this function's `wordDiffs` branch, so nothing could.
 * These tests exist so the next change to either the order or the budget has to
 * say what it is doing to the row.
 */
import { describe, it, expect } from 'vitest';

import { previewText } from '@features/compare/components/ChangeList.js';
import type { DiffOp, WordDiff } from '@features/compare/model.js';

const ins = (text: string): WordDiff => ({ op: 'ins', text });
const del = (text: string): WordDiff => ({ op: 'del', text });
const eq = (text: string): WordDiff => ({ op: 'eq', text });

const opWith = (wordDiffs: WordDiff[]): DiffOp =>
  ({ type: 'replace', wordDiffs }) as unknown as DiffOp;

// Long enough that the insertion alone overruns the 80-character budget.
const LONG_INS =
  'shall indemnify and hold harmless each Indemnified Person against all Losses arising from any breach of this Agreement';
const SHORT_DEL = 'shall be liable for loss';

describe('§8.3 change-list preview under ADR-0043', () => {
  it('leads with the insertion, because stage 6.5 put it first', () => {
    const preview = previewText(opWith([eq('The Vendor '), ins('non-exclusive'), del('exclusive')]));
    expect(preview).toBe('+non-exclusive -exclusive');
    // The guard against this test going vacuous: it is only meaningful while the
    // stream really is ins-first, which is stage 6.5's job and is asserted end to
    // end in tests/unit/compare.acceptance.test.ts.
    expect(preview.indexOf('+')).toBeLessThan(preview.indexOf('-'));
  });

  it('drops the struck text when the insertion alone overruns the budget', () => {
    const preview = previewText(opWith([eq('The Vendor '), ins(LONG_INS), del(SHORT_DEL)]));
    expect(preview).toBe(
      '+shall indemnify and hold harmless each Indemnified Person against all Losses ar…',
    );
    expect(preview).not.toContain(SHORT_DEL);
    // And the fact this pins: the OLD order lost the other side just as hard, so
    // this is a change of which side is favoured, not a new loss of information.
    const oldOrder = previewText(opWith([eq('The Vendor '), del(SHORT_DEL), ins(LONG_INS)]));
    expect(oldOrder).toContain(SHORT_DEL);
    expect(oldOrder).not.toContain('breach of this Agreement');
  });

  it('is unaffected when both sides fit inside the budget', () => {
    const short = [eq('pay the '), ins('Revised Fee'), del('Fee')];
    expect(previewText(opWith(short))).toBe('+Revised Fee -Fee');
    expect(previewText(opWith([short[0], short[2], short[1]]))).toBe('-Fee +Revised Fee');
    // Same characters either way — the budget is the only thing that makes the
    // order observable here, which is why the case above is the one that matters.
    expect(previewText(opWith(short)).length).toBe(
      previewText(opWith([short[0], short[2], short[1]])).length,
    );
  });
});
