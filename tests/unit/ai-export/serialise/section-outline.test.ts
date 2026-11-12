import { describe, it, expect } from 'vitest';
import { serialiseSectionOutline } from '@features/ai-export/serialise/section-outline.js';
import { makeNode, makeIndex, paraTextFrom } from '../_helpers.js';

describe('serialiseSectionOutline', () => {
  const index = makeIndex([
    makeNode({ key: 'clause:1', displayLabel: 'Clause 1', paraId: 'p1', level: 0, kind: 'clause' }),
    makeNode({ key: 'clause:1.1', displayLabel: 'Clause 1.1', paraId: 'p2', level: 1, kind: 'clause', parentKey: 'clause:1' }),
    makeNode({ key: 'schedule:1', displayLabel: 'Schedule 1', paraId: 'p3', level: 0, kind: 'schedule' }),
  ]);
  const paraText = paraTextFrom({ p1: 'Definitions', p2: 'Interpretation', p3: 'Completion Steps' });

  it('rebuilds nesting from parentKey', () => {
    const out = serialiseSectionOutline(index, { paraText });
    expect(out.clauses).toHaveLength(2); // Clause 1 + Schedule 1 are roots
    const clause1 = out.clauses[0]!;
    expect(clause1.number).toBe('Clause 1');
    expect(clause1.children).toHaveLength(1);
    expect(clause1.children![0]!.number).toBe('Clause 1.1');
  });

  it('uses displayLabel (never paraId) and forwards kind + 1-based level', () => {
    const out = serialiseSectionOutline(index, { paraText });
    const json = JSON.stringify(out);
    expect(json).not.toContain('p1');
    expect(json).not.toContain('paraId');
    expect(out.clauses[0]!.kind).toBe('clause');
    expect(out.clauses[0]!.level).toBe(1);          // 0-based node → 1-based export
    expect(out.clauses[0]!.children![0]!.level).toBe(2);
    expect(out.clauses[1]!.kind).toBe('schedule');
    expect(out.clauses[0]!.heading).toBe('Definitions');
  });

  it('omits the children key for leaf nodes', () => {
    const out = serialiseSectionOutline(index, { paraText });
    expect('children' in out.clauses[1]!).toBe(false); // Schedule 1 has no children
  });
});
