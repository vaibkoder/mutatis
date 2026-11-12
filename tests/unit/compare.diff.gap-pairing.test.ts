/**
 * ADR-0039 — `alignGap` and `gapPairingFor`, the two halves of Level-1 gap pairing.
 * Spec: spec/features/08-compare.md §5.1.
 *
 * The behaviour a reader sees is pinned through the pipeline in
 * `compare.pipeline.gap-pairing.test.ts`; this file pins the properties that
 * behaviour rests on and that a pipeline test would only show indirectly.
 */
import { describe, it, expect } from 'vitest';

import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import type { ParaRecord } from '@features/compare/model.js';
import {
  alignGap,
  gapPairingFor,
  MAX_ALIGN_CELLS,
  MAX_ALIGN_WORK,
  MAX_ALIGN_WORK_TOTAL,
  ALIGN_CELL_OVERHEAD,
  paragraphSimilarity,
  survivesDemotion,
} from '@features/compare/diff/index.js';
import { fnv1a } from '@features/compare/preprocess/index.js';

const CONFIG = getCompareConfig(SG_PROFILE);

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

/** A score from a lookup table, so the alignment can be tested apart from Dice. */
function tableScore(table: number[][]): (o: ParaRecord, r: ParaRecord) => number {
  return (o, r) => table[o.index]![r.index]!;
}

describe('alignGap — the order-preserving maximal alignment', () => {
  it('chooses the non-crossing set with the greatest total, not the greatest single pair', () => {
    // (0,1) alone scores 0.9; (0,0)+(1,1) scores 1.4 and wins.
    const o = side('a', 'b');
    const r = side('x', 'y');
    expect(alignGap(o, r, tableScore([[0.7, 0.9], [0, 0.7]]))).toEqual([
      [0, 0],
      [1, 1],
    ]);
  });

  it('never returns crossing pairs, whatever the scores', () => {
    // (0,1) and (1,0) both score high and cross; only one may be chosen.
    expect(alignGap(side('a', 'b'), side('x', 'y'), tableScore([[0, 0.95], [0.9, 0]]))).toEqual([[0, 1]]);
  });

  it('never chooses a zero-score pair', () => {
    expect(alignGap(side('a', 'b'), side('x', 'y'), tableScore([[0, 0], [0, 0]]))).toEqual([]);
  });

});

describe('gapPairingFor — the candidate score', () => {
  const a = para(0, 'Working Day means any day on which clearing banks in Singapore are open for business.');
  const b = para(1, 'Working Day means any day on which clearing banks in Singapore and London are open for business.');
  const unrelated = para(2, 'Approved Supplier means a supplier approved in writing by the Buyer.');
  const pairing = gapPairingFor(CONFIG, [a, unrelated], [b]);

  it('scores a candidate with exactly the Dice the demotion will compute', () => {
    expect(pairing.score(a, b)).toBe(paragraphSimilarity(a.normText, b.normText, CONFIG));
    expect(pairing.score(a, b)).toBeGreaterThanOrEqual(CONFIG.replaceSimilarityThreshold);
  });

  it('agrees with paragraphSimilarity on EVERY pair it scores above zero (the interned fast path)', () => {
    // The fast path computes the multiset intersection over interned, sorted
    // ids; `paragraphSimilarity` over string maps. ADR-0039's promise that a
    // chosen pair is never demoted rests on the two being the same number.
    const words = 'the Buyer shall pay Price within thirty days of invoice and Seller may charge interest on late sums'.split(' ');
    let seed = 11;
    const rnd = (): number => (seed = (seed * 1103515245 + 12345) % 2147483648) / 2147483648;
    const sentence = (): string =>
      Array.from({ length: 4 + Math.floor(rnd() * 14) }, () => words[Math.floor(rnd() * words.length)]).join(' ') + '.';
    const os = Array.from({ length: 60 }, (_, i) => para(i, sentence()));
    const rs = Array.from({ length: 60 }, (_, i) => para(i, sentence()));
    const p = gapPairingFor(CONFIG, os, rs);
    let scored = 0;
    for (const o of os) {
      for (const r of rs) {
        const w = p.score(o, r);
        const dice = paragraphSimilarity(o.normText, r.normText, CONFIG);
        if (w > 0) {
          scored += 1;
          expect(w).toBe(dice);
        } else {
          expect(dice < CONFIG.replaceSimilarityThreshold || o.normText === '' || r.normText === '').toBe(true);
        }
      }
    }
    expect(scored, 'the generator must produce candidates or this proves nothing').toBeGreaterThan(100);
  });

  it('a pair below the demotion threshold is not a candidate', () => {
    expect(pairing.score(a, unrelated)).toBe(0);
  });

  it('a contentless paragraph is never a candidate, though two of them score Dice 1', () => {
    const stars = para(0, '* * *');
    const dashes = para(0, '- - -');
    expect(paragraphSimilarity(stars.normText, dashes.normText, CONFIG)).toBe(1);
    expect(gapPairingFor(CONFIG, [stars], [dashes]).score(stars, dashes)).toBe(0);
  });
});

describe('gapPairingFor — rule 1, the lock runs the real passes', () => {
  it('locks a positional pair that survives ADR-0026 on its own', () => {
    const o = side('Working Day means any day on which clearing banks in Singapore are open for business.');
    const r = side('Working Day means any day on which clearing banks in Singapore and London are open for business.');
    expect(gapPairingFor(CONFIG, o, r).lockPositional(o, r)).toEqual([true]);
  });

  it('does not lock a positional pair ADR-0026 would demote', () => {
    const o = side('Approved Supplier means a supplier approved in writing by the Buyer.');
    const r = side('Working Day means any day on which clearing banks in Singapore are open for business.');
    expect(gapPairingFor(CONFIG, o, r).lockPositional(o, r)).toEqual([false]);
  });

  it('locks a merge HOST whose pair fails the gate alone — ADR-0032 renders it inline (the task-7 defect)', () => {
    const head = 'The Purchaser shall notify the Vendor promptly.';
    const tail =
      'The notice shall specify the nature of the claim in reasonable detail and attach copies of all relevant correspondence received.';
    const joined =
      'The Purchaser shall within five Business Days notify the Vendor in writing, and the notice shall specify the nature of the claim in reasonable detail and attach copies of all relevant correspondence received.';
    const o = side(head, tail);
    const r = side(joined);
    // Alone, the positional pair would be demoted …
    expect(survivesDemotion(head, joined, CONFIG)).toBe(false);
    // … but the merge pass claims it, so it renders inline today and is locked.
    expect(gapPairingFor(CONFIG, o, r).lockPositional(o, r)).toEqual([true]);
  });
});

describe('gapPairingFor — rule 2, stretches it leaves positional', () => {
  const w = para(0, 'Working Day means any day on which clearing banks in Singapore are open for business.');
  const wEdited = para(0, 'Working Day means any day on which clearing banks in Singapore and London are open for business.');
  const other = para(1, 'Approved Supplier means a supplier approved in writing by the Buyer.');

  it('aligns an ordinary stretch', () => {
    const o = [other, { ...w, index: 1 }];
    const r = [wEdited];
    expect(gapPairingFor(CONFIG, o, r).align(o, r)).toEqual([[1, 0]]);
  });

  it('leaves a stretch positional when it holds a table unit (ADR-0027)', () => {
    const o = [{ ...other, table: [] }, { ...w, index: 1 }];
    const r = [wEdited];
    expect(gapPairingFor(CONFIG, o, r).align(o, r)).toEqual([]);
  });

  it('never pairs a record whose text appears verbatim on the other side (held for a Level-3 move)', () => {
    const o = [other, { ...w, index: 1 }];
    const r = [wEdited];
    // `w` itself also appears elsewhere in the revised document: it is headed
    // for a verbatim move, which re-pairing would take away (08f task 7).
    const revisedDoc = [wEdited, para(1, w.normText)];
    const p = gapPairingFor(CONFIG, o, revisedDoc);
    expect(p.heldForMove(o[1]!, 'original')).toBe(true);
    expect(p.heldForMove(other, 'original')).toBe(false);
    expect(p.align(o, r)).toEqual([]);
  });

  it('holds a REVISED record whose text appears verbatim in the original (the mirror)', () => {
    const o = [w];
    const r = [{ ...other, index: 0 }, { ...wEdited, index: 1 }];
    const originalDoc = [w, para(1, wEdited.normText)];
    const p = gapPairingFor(CONFIG, originalDoc, r);
    expect(p.heldForMove(r[1]!, 'revised')).toBe(true);
    expect(p.heldForMove(r[0]!, 'revised')).toBe(false);
    expect(p.align(o, r)).toEqual([]);
  });

  it('leaves a stretch positional above MAX_ALIGN_WORK (under the cell cap), without running the DP', () => {
    // 800 × 800 paragraphs of 200 distinct tokens: 6.4·10⁵ cells, under the
    // cell cap, but 6.4·10⁵ × (32 + 200) ≈ 1.5·10⁸ cost units, over the budget.
    const long = (i: number, tag: string): ParaRecord =>
      para(i, Array.from({ length: 200 }, (_, k) => `${tag}${i}t${k}`).join(' '));
    const o = Array.from({ length: 800 }, (_, i) => long(i, 'o'));
    const r = o.map((x) => ({ ...x, hash: `${x.hash}-r` }));
    expect(800 * 800).toBeLessThan(MAX_ALIGN_CELLS);
    expect(800 * 800 * (ALIGN_CELL_OVERHEAD + 200)).toBeGreaterThan(MAX_ALIGN_WORK);
    const p = gapPairingFor(CONFIG, o, r);
    const t0 = performance.now();
    expect(p.align(o, r)).toEqual([]);
    // Loose on purpose: most of this is tokenising 1,600 × 200 tokens, which
    // comes before the check (a 1,500 ms bound flaked under load, task 11).
    // The DP itself would take tens of seconds.
    expect(performance.now() - t0, 'the budget must be checked before the DP runs').toBeLessThan(5000);
  });

  it('leaves a stretch positional above MAX_ALIGN_CELLS even when the work budget is not reached', () => {
    // Two-token paragraphs: the token-weighted budget would admit millions of
    // cells, and the DP's two Float64Arrays are m·n each (08f task 9: 8,600² took
    // 7.2 s and ~0.9 GB). 1,001 × 1,000 is just over the cell cap.
    const o = Array.from({ length: 1001 }, (_, i) => para(i, `alpha${i} beta${i}`));
    const r = Array.from({ length: 1000 }, (_, i) => para(i, `alpha${i} beta${i} gamma`));
    expect(1001 * 1000).toBeGreaterThan(MAX_ALIGN_CELLS);
    expect(1001 * 1000 * (ALIGN_CELL_OVERHEAD + 2)).toBeLessThan(MAX_ALIGN_WORK);
    const p = gapPairingFor(CONFIG, o, r);
    const t0 = performance.now();
    expect(p.align(o, r)).toEqual([]);
    expect(performance.now() - t0, 'the cap must be checked before the DP runs').toBeLessThan(1000);
  });

  it('stops aligning once the comparison has spent MAX_ALIGN_WORK_TOTAL, each stretch under its own budget', () => {
    // Many large gaps paid the per-stretch budget once each (ten 999 × 1,000
    // gaps took 6.3 s of Stage 2, 08f task 11).
    const o = Array.from({ length: 999 }, (_, i) => para(i, `alpha${i} beta${i}`));
    const r = Array.from({ length: 1000 }, (_, i) => para(i, `alpha${i} beta${i} gamma`));
    const work = 999 * 1000 * (ALIGN_CELL_OVERHEAD + 2.5);
    expect(work).toBeLessThan(MAX_ALIGN_WORK);
    const affordable = Math.floor(MAX_ALIGN_WORK_TOTAL / work);
    const p = gapPairingFor(CONFIG, o, r);
    for (let k = 0; k < affordable; k++) expect(p.align(o, r), `stretch ${k}`).toHaveLength(999);
    expect(p.align(o, r)).toEqual([]);
    // A fresh comparison has its whole budget again.
    expect(gapPairingFor(CONFIG, o, r).align(o, r)).toHaveLength(999);
  });

  it('a whole-document gap of 400 × 401 aligns in well under a few seconds (the token cache is load-bearing)', () => {
    // The prototype without the per-record cache took 86 s at 950 × 951; this
    // size takes ~15 s uncached and well under a second cached.
    const vocab = 'party agreement clause notice payment completion warranty indemnity schedule period'.split(' ');
    const make = (i: number, suffix: string): string =>
      Array.from({ length: 40 }, (_, k) => `${vocab[(i * 7 + k) % vocab.length]}${i}x${k % 5}`).join(' ') + suffix;
    const o = Array.from({ length: 400 }, (_, i) => para(i, make(i, '.')));
    const r = [para(0, 'An inserted opening paragraph.'), ...o.map((_, i) => para(i + 1, make(i, ' amended.')))];
    const fresh = gapPairingFor(CONFIG, o, r);
    const t0 = performance.now();
    const pairs = fresh.align(o, r);
    const ms = performance.now() - t0;
    expect(pairs).toHaveLength(400);
    expect(pairs[0]).toEqual([0, 1]);
    expect(ms).toBeLessThan(5000);
  });
});
