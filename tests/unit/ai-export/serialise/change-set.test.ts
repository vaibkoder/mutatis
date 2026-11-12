import { describe, it, expect } from 'vitest';
import { serialiseChangeSet } from '@features/ai-export/serialise/change-set.js';
import type { DiffOp, WordDiff } from '@features/compare/model.js';

function op(p: Partial<DiffOp> & { id: string; type: DiffOp['type'] }): DiffOp {
  return {
    origRange: null, revRange: null, partId: 'document',
    classification: 'substantiveText', classes: ['substantiveText'],
    author: 'A', date: '2026-01-01',
    ...p,
  } as DiffOp;
}

const wd = (op: WordDiff['op'], text: string): WordDiff => ({ op, text });

const deps = {
  baseline: 'v1.docx',
  revised: 'v2.docx',
  clauseFor: (o: DiffOp) => (o.partId === 'document' ? 'Body' : o.partId),
};

describe('serialiseChangeSet', () => {
  it('replace emits both original and revised from wordDiffs', () => {
    const out = serialiseChangeSet([
      op({ id: 'r1', type: 'replace', wordDiffs: [wd('eq', 'The price is '), wd('del', '100'), wd('ins', '200'), wd('eq', '.')] }),
    ], deps);
    expect(out.changes[0]!.original).toBe('The price is 100.');
    expect(out.changes[0]!.revised).toBe('The price is 200.');
  });

  it('delete emits only original; insert emits only revised', () => {
    const out = serialiseChangeSet([
      op({ id: 'd1', type: 'delete', wordDiffs: [wd('del', 'Removed clause.')] }),
      op({ id: 'i1', type: 'insert', wordDiffs: [wd('ins', 'New clause.')] }),
    ], deps);
    const del = out.changes[0]!;
    expect(del.original).toBe('Removed clause.');
    expect(del.revised).toBeUndefined();
    const ins = out.changes[1]!;
    expect(ins.revised).toBe('New clause.');
    expect(ins.original).toBeUndefined();
  });

  it('move carries move_target and both sides from paragraph text', () => {
    const out = serialiseChangeSet([
      op({ id: 'm1', type: 'move', moveTargetId: 'x', partId: 'document' }),
    ], {
      ...deps,
      origParaText: () => 'The moved paragraph.',
      revParaText: () => 'The moved paragraph.',
      moveTarget: () => 'Clause 9',
    });
    const mv = out.changes[0]!;
    expect(mv.type).toBe('move');
    expect(mv.move_target).toBe('Clause 9');
    expect(mv.original).toBe('The moved paragraph.');
    expect(mv.revised).toBe('The moved paragraph.');
  });

  it('delete/insert without wordDiffs fall back to whole-paragraph callbacks (production shape)', () => {
    // Pure delete/insert ops carry NO wordDiffs (those exist only for 'replace'),
    // so original/revised must come from the para-text callbacks — the live
    // ComparePanel path. A serialiser that only read wordDiffs would emit ''.
    const out = serialiseChangeSet([
      op({ id: 'd1', type: 'delete', origRange: { start: 2, end: 2 } }),
      op({ id: 'i1', type: 'insert', revRange: { start: 5, end: 5 } }),
    ], {
      ...deps,
      origParaText: o => (o.id === 'd1' ? 'The deleted paragraph in full.' : ''),
      revParaText: o => (o.id === 'i1' ? 'The inserted paragraph in full.' : ''),
    });
    expect(out.changes[0]!.original).toBe('The deleted paragraph in full.');
    expect(out.changes[0]!.revised).toBeUndefined();
    expect(out.changes[1]!.revised).toBe('The inserted paragraph in full.');
    expect(out.changes[1]!.original).toBeUndefined();
  });

  it('skips equal ops and carries baseline/revised + low_confidence', () => {
    const out = serialiseChangeSet([
      op({ id: 'e1', type: 'equal' }),
      op({ id: 'r2', type: 'replace', lowConfidence: true, wordDiffs: [wd('del', 'a'), wd('ins', 'b')] }),
    ], deps);
    expect(out.baseline).toBe('v1.docx');
    expect(out.revised).toBe('v2.docx');
    expect(out.changes).toHaveLength(1);
    expect(out.changes[0]!.low_confidence).toBe(true);
  });

  it('produces no diff markup characters (F1)', () => {
    const out = serialiseChangeSet([
      op({ id: 'r3', type: 'replace', wordDiffs: [wd('del', 'old'), wd('ins', 'new')] }),
    ], deps);
    const json = JSON.stringify(out);
    expect(json).not.toContain('~~');
    expect(json).not.toContain('<w:');
  });
});
