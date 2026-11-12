import { describe, it, expect } from 'vitest';

import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import type { CompareConfig } from '@shared/config/index.js';
import type { ParaRecord, DiffOp } from '@features/compare/model.js';
import {
  structuralDiff,
  demoteDissimilarReplaces,
  paragraphSimilarity,
  contiguousShare,
} from '@features/compare/diff/index.js';
import { fnv1a } from '@features/compare/preprocess/index.js';

const CONFIG: CompareConfig = getCompareConfig(SG_PROFILE);
const AUTHOR = 'Test';
const DATE = '2026-06-17T00:00:00.000Z';

function para(index: number, text: string): ParaRecord {
  return {
    index,
    partId: 'document',
    rawXml: `<w:p>${text}</w:p>`,
    normText: text,
    normTextAggressive: text.toLowerCase(),
    numberingPrefix: '',
    hash: fnv1a(text),
    styleId: null,
    isTableCell: false,
    isSignatureBlock: false,
    isHeaderFooter: false,
    isFootnote: false,
  };
}
const side = (...texts: string[]): ParaRecord[] => texts.map((t, i) => para(i, t));

/** structural → demote, exactly as the Word pipeline runs them (Stage 2 → 2.5). */
function diff(
  original: ParaRecord[],
  revised: ParaRecord[],
  cfg: CompareConfig = CONFIG,
): DiffOp[] {
  return demoteDissimilarReplaces(
    structuralDiff(original, revised, AUTHOR, DATE),
    original,
    revised,
    cfg,
  );
}

describe('paragraphSimilarity — Sørensen–Dice over content tokens', () => {
  it('is 1 for identical text and (near) 0 for disjoint text', () => {
    expect(paragraphSimilarity('The quick brown fox', 'The quick brown fox', CONFIG)).toBe(1);
    expect(paragraphSimilarity('alpha beta gamma', 'delta epsilon zeta', CONFIG)).toBe(0);
  });

  it('stays high for a small in-place edit (one value changed)', () => {
    const sim = paragraphSimilarity(
      'The Borrower shall pay S$40,000 within 30 days.',
      'The Borrower shall pay S$50,000 within 30 days.',
      CONFIG,
    );
    expect(sim).toBeGreaterThan(0.5);
  });

  it('falls below 0.5 for a wholesale rewrite that only shares boilerplate', () => {
    const sim = paragraphSimilarity(
      'The Borrower shall maintain insurance at all times.',
      'Confidentiality survives termination of this Agreement.',
      CONFIG,
    );
    expect(sim).toBeLessThan(0.5);
  });
});

describe('demoteDissimilarReplaces — change-complexity threshold (ADR-0026)', () => {
  const A = 'ANCHOR ONE — a unique opening clause.';
  const B = 'ANCHOR TWO — a unique closing clause.';
  const OLD = 'The Borrower shall maintain insurance at all times.';
  const NEW = 'Confidentiality survives termination of this Agreement.';

  it('demotes a wholesale-rewritten replace pair to a clean delete + insert', () => {
    const original = side(A, OLD, B);
    const revised = side(A, NEW, B);

    const ops = diff(original, revised);

    // No inline replace survives — it became delete + insert.
    expect(ops.some((o) => o.type === 'replace')).toBe(false);

    const del = ops.filter((o) => o.type === 'delete');
    const ins = ops.filter((o) => o.type === 'insert');
    expect(del).toHaveLength(1);
    expect(ins).toHaveLength(1);

    // The whole original paragraph is the deletion; the whole revised is the insertion.
    expect(del[0]!.origRange).toEqual({ start: 1, end: 1 });
    expect(del[0]!.revRange).toBeNull();
    expect(original[del[0]!.origRange!.start]!.normText).toBe(OLD);

    expect(ins[0]!.revRange).toEqual({ start: 1, end: 1 });
    expect(ins[0]!.origRange).toBeNull();
    expect(revised[ins[0]!.revRange!.start]!.normText).toBe(NEW);

    // Both demoted ops are structural and carry unique ids.
    expect(del[0]!.classification).toBe('structural');
    expect(ins[0]!.classification).toBe('structural');
    expect(del[0]!.id).not.toBe(ins[0]!.id);

    // Anchors stay equal — only the rewritten paragraph changed.
    expect(ops.filter((o) => o.type === 'equal')).toHaveLength(2);
  });

  it('NEVER suppresses — both full paragraphs remain represented after demotion', () => {
    const original = side(A, OLD, B);
    const revised = side(A, NEW, B);

    const ops = diff(original, revised);

    const deletedText = ops
      .filter((o) => o.type === 'delete' && o.origRange)
      .map((o) => original[o.origRange!.start]!.normText);
    const insertedText = ops
      .filter((o) => o.type === 'insert' && o.revRange)
      .map((o) => revised[o.revRange!.start]!.normText);

    expect(deletedText).toContain(OLD); // nothing dropped from the original
    expect(insertedText).toContain(NEW); // nothing dropped from the revised
  });

  it('leaves a similar (small-edit) replace pair untouched for inline word diff', () => {
    const original = side(A, 'The Borrower shall pay S$40,000 within 30 days.', B);
    const revised = side(A, 'The Borrower shall pay S$50,000 within 30 days.', B);

    const ops = diff(original, revised);

    expect(ops.filter((o) => o.type === 'replace')).toHaveLength(1);
    expect(ops.some((o) => o.type === 'delete')).toBe(false);
    expect(ops.some((o) => o.type === 'insert')).toBe(false);
  });

  it('threshold 0 disables the pass — a dissimilar replace stays a replace', () => {
    const cfg: CompareConfig = { ...CONFIG, replaceSimilarityThreshold: 0 };
    const original = side(A, OLD, B);
    const revised = side(A, NEW, B);

    const ops = diff(original, revised, cfg);
    expect(ops.filter((o) => o.type === 'replace')).toHaveLength(1);
  });

  it('KEEPS a pair inline when a large CONTIGUOUS block was cut (ADR-0041)', () => {
    // ABLATION: the contiguity escape. This is a known defect class — a
    // paragraph pair with a large contiguous cut can score low on Dice
    // (around 0.4) and be demoted, whereas a conventional redline keeps it
    // inline.
    //
    // Sørensen-Dice is a SET measure, so it reads length asymmetry as
    // dissimilarity — and a large contiguous deletion is nothing but length
    // asymmetry. Here the revised paragraph is the original's first sentence,
    // verbatim and whole: inline that is ONE strike-through with every
    // surviving word untouched. Demoting it strikes the survivor too and makes
    // the reader prove it did not change (P1/P2).
    const kept = 'The Vendor shall maintain insurance at all times.';
    const cut =
      `${kept} The Vendor shall procure that each Subsidiary maintains equivalent cover ` +
      'with a reputable insurer of international standing, and shall produce evidence of ' +
      'renewal to the Purchaser on each anniversary of Completion.';

    // The premise: Dice alone WOULD demote this.
    expect(paragraphSimilarity(cut, kept, CONFIG)).toBeLessThan(
      CONFIG.replaceSimilarityThreshold,
    );
    // And the survivor is one unbroken run of the shorter side.
    expect(contiguousShare(cut, kept, CONFIG)).toBe(1);

    const original = side(A, cut, B);
    const revised = side(A, kept, B);
    const ops = diff(original, revised);

    const replaces = ops.filter((o) => o.type === 'replace');
    expect(replaces).toHaveLength(1);
    expect(replaces[0]!.origRange).toEqual({ start: 1, end: 1 });
    expect(replaces[0]!.revRange).toEqual({ start: 1, end: 1 });
    expect(ops.some((o) => o.type === 'delete' || o.type === 'insert')).toBe(false);
  });

  it('STILL demotes a rewrite whose shared words are scattered, not contiguous', () => {
    // The other side of the escape, and why it is a contiguity test and not a
    // length test. Both paragraphs here share most of their vocabulary — the
    // same parties, the same modal, the same boilerplate tail — but no shared
    // run is long enough to read as an untouched block, which is exactly the
    // "equal islands in a sea of ins/del" salad ADR-0026 exists to prevent.
    const before =
      'The Vendor shall promptly indemnify the Purchaser against all Losses arising out of any breach of the Warranties.';
    const after =
      'The Purchaser may, against the Vendor, recover each Loss suffered in connection with any breach, howsoever arising.';

    expect(paragraphSimilarity(before, after, CONFIG)).toBeLessThan(
      CONFIG.replaceSimilarityThreshold,
    );
    expect(contiguousShare(before, after, CONFIG)).toBeLessThan(
      CONFIG.replaceContiguityThreshold,
    );

    const ops = diff(side(A, before, B), side(A, after, B));
    expect(ops.some((o) => o.type === 'replace')).toBe(false);
    expect(ops.filter((o) => o.type === 'delete')).toHaveLength(1);
    expect(ops.filter((o) => o.type === 'insert')).toHaveLength(1);
  });

  it('passes non-replace ops through unchanged (a pure inserted paragraph)', () => {
    // A new middle paragraph between two anchors is a pure `insert` — no replace
    // pair exists, so the demotion pass must leave the op list as-is.
    const original = side(A, B);
    const revised = side(A, 'a brand-new inserted middle paragraph', B);

    const ops = diff(original, revised);

    expect(ops.filter((o) => o.type === 'insert')).toHaveLength(1);
    expect(ops.filter((o) => o.type === 'equal')).toHaveLength(2);
    expect(ops.some((o) => o.type === 'delete')).toBe(false);
    expect(ops.some((o) => o.type === 'replace')).toBe(false);
  });
});
