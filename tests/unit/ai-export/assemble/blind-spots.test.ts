import { describe, it, expect } from 'vitest';
import { buildBlindSpots } from '@features/ai-export/assemble/blind-spots.js';
import type { CheckId } from '@features/ai-export/types.js';

const ALL: CheckId[] = ['defined-terms', 'cross-reference', 'factual-consistency', 'compare', 'party-role'];

const FIXED = [
  'Headers, footers, and text boxes are not scanned.',
  'Image-only content (scanned pages, embedded pictures) is not text-compared.',
  'Cross-document and external references are detected but not resolved (Office JS cannot read other files).',
];

describe('buildBlindSpots', () => {
  it('emits a disclosure for each checker that did not run', () => {
    const spots = buildBlindSpots(['compare']);
    expect(spots).toContain('Defined-terms checker was not run; glossary and usage findings absent.');
    expect(spots).toContain('Cross-reference / section-index checker was not run; clause-tree absent.');
    expect(spots).toContain('Factual-consistency checker was not run; deal dates, amounts, and placeholders absent.');
    expect(spots).toContain('Party-role checker was not run; cast list absent.');
    // The one that ran has no absence disclosure.
    expect(spots).not.toContain('Compare was not run; change set absent.');
  });

  it('always appends the three fixed entries regardless of checks_run', () => {
    for (const spots of [buildBlindSpots([]), buildBlindSpots(ALL)]) {
      for (const fixed of FIXED) expect(spots).toContain(fixed);
    }
  });

  it('with all checks run, only the three fixed entries remain', () => {
    expect(buildBlindSpots(ALL)).toEqual(FIXED);
  });

  it('with nothing run, lists all five absences + three fixed', () => {
    expect(buildBlindSpots([])).toHaveLength(ALL.length + FIXED.length);
  });
});
