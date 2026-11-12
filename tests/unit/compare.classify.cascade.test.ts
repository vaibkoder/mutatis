/**
 * Feature 08 — Compare · renumbering-cascade collapse tests (Task B3).
 *
 * Covers T3 (47 +1 Clause shifts → one CascadeAnnotation) and the REQUIRED
 * ADR-0008 suppression-guard: computeCascades never removes a DiffOp.
 */

import { describe, it, expect } from 'vitest';
import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import type { CompareConfig } from '@shared/config/index.js';
import type { DiffOp } from '../../src/features/compare/model.js';
import { computeCascades } from '../../src/features/compare/classify/cascade.js';
import { wordDiff } from '../../src/features/compare/diff/index.js';
import { runComparePipeline } from '../../src/features/compare/pipeline.js';
import { extractTextRecords } from '../../src/features/compare/io/text/extract-text-records.js';

const DEFAULT_CONFIG: CompareConfig = getCompareConfig(SG_PROFILE);

/**
 * Build a crossRefRenumber op: a clauseRef "<kind> <old>" → "<kind> <new>"
 * where new = old + shift. The trailing number drives the shift.
 */
function makeCrossRefOp(i: number, kind: string, shift: number, base = 2): DiffOp {
  const oldN = base + i;
  const newN = oldN + shift;
  return {
    id: `xref-${i}`,
    type: 'replace',
    origRange: { start: i, end: i },
    revRange: { start: i, end: i },
    partId: 'document',
    wordDiffs: [
      { op: 'del', text: `${kind} 4.${oldN}`, atomic: 'clauseRef' },
      { op: 'ins', text: `${kind} 4.${newN}`, atomic: 'clauseRef' },
    ],
    classification: 'crossRefRenumber',
    classes: ['crossRefRenumber'],
    author: 'Test',
    date: '2026-06-08T00:00:00Z',
  };
}

// ---------------------------------------------------------------------------
// T3 — 47 +1 Clause shifts collapse into one annotation
// ---------------------------------------------------------------------------

describe('cascade — T3 renumbering collapse', () => {
  it('47 consistent +1 Clause shifts collapse into one CascadeAnnotation', () => {
    const ops: DiffOp[] = Array.from({ length: 47 }, (_, i) => makeCrossRefOp(i, 'Clause', +1));
    const { ops: result, cascades } = computeCascades(ops, DEFAULT_CONFIG);

    expect(cascades).toHaveLength(1);
    const cascade = cascades[0]!;
    expect(cascade.tokenKind).toBe('Clause');
    expect(cascade.shift).toBe(1);
    expect(cascade.affectedCount).toBe(47);
    expect(cascade.exampleRefs).toHaveLength(3);
    expect(cascade.exampleRefs[0]).toBe('Clause 4.2 → 4.3');
    expect(cascade.id).toBe('cascade-0');

    // All 47 ops still present and tagged.
    expect(result).toHaveLength(47);
    expect(result.every((o) => o.collapsedInto === 'cascade-0')).toBe(true);
  });

  it('a group below threshold is NOT collapsed (no annotation, no collapsedInto)', () => {
    const ops: DiffOp[] = Array.from({ length: 2 }, (_, i) => makeCrossRefOp(i, 'Clause', +1));
    const { ops: result, cascades } = computeCascades(ops, { ...DEFAULT_CONFIG, renumberThreshold: 3 });
    expect(cascades).toHaveLength(0);
    expect(result.every((o) => o.collapsedInto === undefined)).toBe(true);
    // Still present.
    expect(result).toHaveLength(2);
  });

  it('separates groups by tokenKind and by shift', () => {
    const ops: DiffOp[] = [
      ...Array.from({ length: 4 }, (_, i) => makeCrossRefOp(i, 'Clause', +1)),
      ...Array.from({ length: 5 }, (_, i) => makeCrossRefOp(100 + i, 'Schedule', +1)),
      ...Array.from({ length: 3 }, (_, i) => makeCrossRefOp(200 + i, 'Clause', -1)),
    ];
    const { ops: result, cascades } = computeCascades(ops, { ...DEFAULT_CONFIG, renumberThreshold: 3 });

    expect(cascades).toHaveLength(3);
    const byKey = new Map(cascades.map((c) => [`${c.tokenKind} ${c.shift}`, c]));
    expect(byKey.get('Clause 1')?.affectedCount).toBe(4);
    expect(byKey.get('Schedule 1')?.affectedCount).toBe(5);
    expect(byKey.get('Clause -1')?.affectedCount).toBe(3);
    // Every op preserved.
    expect(result).toHaveLength(12);
    expect(result.filter((o) => o.collapsedInto !== undefined)).toHaveLength(12);
  });

  it('ignores ops that are not classified crossRefRenumber', () => {
    const ops: DiffOp[] = [
      ...Array.from({ length: 4 }, (_, i) => makeCrossRefOp(i, 'Clause', +1)),
      {
        id: 'substantive-1',
        type: 'replace',
        origRange: { start: 50, end: 50 },
        revRange: { start: 50, end: 50 },
        partId: 'document',
        wordDiffs: [{ op: 'ins', text: 'some new wording' }],
        classification: 'substantiveText',
        classes: ['substantiveText'],
        author: 'Test',
        date: '2026-06-08T00:00:00Z',
      },
    ];
    const { ops: result, cascades } = computeCascades(ops, { ...DEFAULT_CONFIG, renumberThreshold: 3 });
    expect(cascades).toHaveLength(1);
    expect(result).toHaveLength(5);
    // The substantive op is untouched (not collapsed).
    expect(result.find((o) => o.id === 'substantive-1')?.collapsedInto).toBeUndefined();
  });

  it('emits at most three example refs even for large cascades', () => {
    const ops: DiffOp[] = Array.from({ length: 47 }, (_, i) => makeCrossRefOp(i, 'Clause', +1));
    const { cascades } = computeCascades(ops, DEFAULT_CONFIG);
    expect(cascades[0]!.exampleRefs.length).toBeLessThanOrEqual(3);
  });
});

// ---------------------------------------------------------------------------
// ADR-0008 — suppression guard (REQUIRED, verbatim logic from the task brief)
// ---------------------------------------------------------------------------

describe('cascade — ADR-0008 suppression guard', () => {
  it('computeCascades never removes a DiffOp from the list', () => {
    // Build 50 crossRefRenumber ops.
    const ops: DiffOp[] = Array.from({ length: 50 }, (_, i) => makeCrossRefOp(i, 'Clause', +1));
    const { ops: result } = computeCascades(ops, { ...DEFAULT_CONFIG, renumberThreshold: 3 });
    expect(result.length).toBe(50); // not 1, not 49 — all 50 present
    expect(result.filter((o) => o.collapsedInto !== undefined).length).toBe(50);
  });

  it('every input op id survives, even mixed cascaded and non-cascaded', () => {
    const ops: DiffOp[] = [
      ...Array.from({ length: 5 }, (_, i) => makeCrossRefOp(i, 'Clause', +1)),
      ...Array.from({ length: 2 }, (_, i) => makeCrossRefOp(900 + i, 'paragraph', +1)), // below threshold
    ];
    const inputIds = ops.map((o) => o.id).sort();
    const { ops: result } = computeCascades(ops, { ...DEFAULT_CONFIG, renumberThreshold: 3 });
    expect(result.map((o) => o.id).sort()).toEqual(inputIds);
  });

  it('does not mutate the input ops array or its members', () => {
    const ops: DiffOp[] = Array.from({ length: 4 }, (_, i) => makeCrossRefOp(i, 'Clause', +1));
    const snapshot = JSON.parse(JSON.stringify(ops));
    const { ops: result } = computeCascades(ops, { ...DEFAULT_CONFIG, renumberThreshold: 3 });
    expect(ops).toEqual(snapshot); // inputs unchanged
    expect(result).not.toBe(ops);
    expect(ops.every((o) => o.collapsedInto === undefined)).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// Real-tokenizer shape (ADR-0029): the noun sits OUTSIDE the mark
// ---------------------------------------------------------------------------

describe('cascade — reference kind recovered from the text before the mark', () => {
  /** A crossRefRenumber op whose wordDiffs come from the real Level-2 diff. */
  function realOp(id: string, orig: string, rev: string): DiffOp {
    return {
      id,
      type: 'replace',
      origRange: { start: 0, end: 0 },
      revRange: { start: 0, end: 0 },
      partId: 'document',
      wordDiffs: wordDiff(orig, rev, DEFAULT_CONFIG),
      classification: 'crossRefRenumber',
      classes: ['crossRefRenumber'],
      author: 'Test',
      date: '2026-06-08T00:00:00Z',
    };
  }

  it('collapses shifts whose marks carry only the bare VALUE', () => {
    // Post-ADR-0029 the tokenizer emits the noun as an ordinary word and tags
    // only the reference VALUE, so a renumbering marks `4.2`→`4.3` and leaves
    // `Clause` unmarked. The cascade grouper therefore cannot read the kind off
    // the token: it has to recover it from the unchanged text before the mark.
    const ops: DiffOp[] = [
      realOp('x0', 'subject to Clause 4.2 above', 'subject to Clause 4.3 above'),
      realOp('x1', 'as required by Clause 7.1', 'as required by Clause 7.2'),
      realOp('x2', 'in accordance with Clause 9.4 hereof', 'in accordance with Clause 9.5 hereof'),
    ];
    // Guard the premise: no mark carries the noun, so `tokenKind` is unavailable.
    for (const op of ops) {
      const marked = op.wordDiffs!.filter((d) => d.op !== 'eq');
      expect(marked.every((d) => !/Clause/i.test(d.text))).toBe(true);
    }

    const { ops: result, cascades } = computeCascades(ops, DEFAULT_CONFIG);
    expect(cascades).toHaveLength(1);
    expect(cascades[0]!.tokenKind).toBe('Clause');
    expect(cascades[0]!.shift).toBe(1);
    expect(cascades[0]!.affectedCount).toBe(3);
    expect(result).toHaveLength(3);
    expect(result.every((o) => o.collapsedInto === cascades[0]!.id)).toBe(true);
  });

  it('normalises the recovered noun to the profile spelling, so case does not split the group', () => {
    // Drafters mix "Clause 4.2" and "clause 4.2" freely. Both are the same
    // cross-reference kind and must land in ONE cascade, reported under the
    // profile's own spelling.
    const ops: DiffOp[] = [
      realOp('u0', 'subject to Clause 4.2 above', 'subject to Clause 4.3 above'),
      realOp('u1', 'subject to Clause 5.1 above', 'subject to Clause 5.2 above'),
      realOp('u2', 'subject to Clause 6.7 above', 'subject to Clause 6.8 above'),
      realOp('l0', 'as set out in clause 8.2', 'as set out in clause 8.3'),
      realOp('l1', 'as set out in clause 9.9', 'as set out in clause 9.10'),
      realOp('l2', 'as set out in clause 11.4', 'as set out in clause 11.5'),
    ];
    const { ops: result, cascades } = computeCascades(ops, DEFAULT_CONFIG);

    expect(cascades).toHaveLength(1);
    expect(cascades[0]!.tokenKind).toBe('Clause');
    expect(cascades[0]!.affectedCount).toBe(6);
    // ADR-0008: every op survives, annotated rather than removed.
    expect(result).toHaveLength(6);
    expect(result.map((o) => o.id).sort()).toEqual(['l0', 'l1', 'l2', 'u0', 'u1', 'u2']);
  });

  it('does not group a value whose noun is not a reference token for the profile', () => {
    // "Subclause" is not in `referenceTokens`, so its number is not a tagged
    // cross-reference and must not be swept into a Clause cascade.
    const ops: DiffOp[] = [
      realOp('s0', 'under Subclause 4.2 hereof', 'under Subclause 4.3 hereof'),
      realOp('s1', 'under Subclause 5.2 hereof', 'under Subclause 5.3 hereof'),
      realOp('s2', 'under Subclause 6.2 hereof', 'under Subclause 6.3 hereof'),
    ];
    const { ops: result, cascades } = computeCascades(ops, DEFAULT_CONFIG);
    expect(cascades).toHaveLength(0);
    expect(result).toHaveLength(3);
  });
});

// ---------------------------------------------------------------------------
// The collapse must not depend on prose the renumbering never touched
// ---------------------------------------------------------------------------

describe('cascade — fires on realistic clause prose', () => {
  /**
   * Drives the WHOLE pipeline from text, so the classification under test is the
   * one the classifier produces rather than one the fixture asserts.
   *
   * The first version of this test built its ops by hand with
   * `classification: 'crossRefRenumber'` already set and called
   * `computeCascades` directly. That pinned nothing: reverting `isDefinedTerm`
   * left it green, which is the precise regression its comment claimed to
   * prevent. `computeCascades` only collapses ops whose PRIMARY classification
   * is `crossRefRenumber`, so the classifier is half the behaviour and cannot be
   * stubbed out of the test.
   */
  async function cascadesFor(paras: readonly string[], revised: readonly string[]) {
    const result = await runComparePipeline({
      originalParas: extractTextRecords(paras.join('\n')),
      revisedParas: extractTextRecords(revised.join('\n')),
      config: DEFAULT_CONFIG,
      author: 'Test',
      date: '2026-06-08T00:00:00Z',
    });
    return result;
  }

  it('collapses a +1 shift whatever prose surrounds the reference', async () => {
    // Row 4 of §6 (`definedTerm`) outranks row 5 (`crossRefRenumber`). While the
    // defined-term test looked at the whole PARAGRAPH rather than at the CHANGE,
    // a renumbering inside a clause opening "The Vendor …" — how most clauses in
    // an SG/E&W contract open — classified `definedTerm` and never reached
    // `computeCascades` at all. The renderings were identical; only the triage
    // differed, and it differed on words the edit had never touched. Pinned at
    // the source in `compare.classify.taxonomy.test.ts` and end-to-end here.
    const shapes: Array<[string, (n: number) => string]> = [
      ['neutral', (n) => `subject to Clause 4.${n} above`],
      ['opens with a party', (n) => `The Vendor shall comply with Clause 4.${n} above`],
      ['mentions a term', (n) => `each Permitted Encumbrance under Clause 4.${n} applies`],
    ];
    for (const [label, text] of shapes) {
      const ns = [2, 12, 22, 32];
      const result = await cascadesFor(
        ns.map((n) => text(n)),
        ns.map((n) => text(n + 1)),
      );

      // The classifier really did reach `crossRefRenumber` — without this the
      // assertions below could pass on an empty cascade list for the wrong
      // reason, which is how the previous version of this test went stale.
      expect(
        result.ops.map((o) => o.classification),
        label,
      ).toEqual(['crossRefRenumber', 'crossRefRenumber', 'crossRefRenumber', 'crossRefRenumber']);

      const cascades = result.summary.renumberingCascades;
      expect(cascades, label).toHaveLength(1);
      expect(cascades[0]!.shift, label).toBe(1);
      expect(cascades[0]!.affectedCount, label).toBe(4);
    }
  });
});
