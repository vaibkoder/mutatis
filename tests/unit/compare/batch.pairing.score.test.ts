/**
 * Pair scoring (ADR-0037; spec §6.3, §6.4; acceptance M5).
 *
 * The claim under test is narrow and load-bearing: the cast signal is consulted
 * ONLY where content and filename have come out level. Inside one deal bundle
 * every document shares the same parties and most defined terms, so cast
 * overlap is near 1 for every candidate pair — a signal that discriminates
 * nothing but, folded into the weighted sum, would still be able to push a pair
 * across the auto-pair boundary. M5 is that property stated as a test.
 */
import { describe, it, expect } from 'vitest';

import {
  isTie,
  rankCandidates,
  scoreAll,
  scorePair,
  type PairScore,
} from '@features/compare/batch/pairing/score.js';
import { DEFAULT_PAIRING_SETTINGS } from '@features/compare/batch/config.js';
import {
  PAIRING_AUTO_PAIR_SCORE,
  PAIRING_NO_MATCH_SCORE,
  PAIRING_TIE_EPSILON,
  PAIRING_WEIGHT_CONTENT,
  PAIRING_WEIGHT_FILENAME,
} from '@features/compare/batch/constants.js';
import type { PairingSketch } from '@features/compare/batch/pairing/sketch.js';

/** A sketch with only the fields scoring reads. */
function sketch(
  docId: string,
  fileName: string,
  hashes: readonly string[],
  cast: readonly string[] = ['seller', 'buyer'],
): PairingSketch {
  return {
    docId,
    side: docId.startsWith('a') ? 'a' : 'b',
    sourceKind: 'docx',
    fileName,
    byteLength: 1024,
    paraHashes: new Set(hashes),
    definedTerms: new Set(cast),
    partyNames: new Set(cast),
    quarantinedCount: 0,
  };
}

const shared = (n: number, prefix = 'p'): string[] =>
  Array.from({ length: n }, (_, i) => `${prefix}-${i}`);

const unique = (prefix: string, n: number): string[] =>
  Array.from({ length: n }, (_, i) => `${prefix}-only-${i}`);

/**
 * A candidate score for the ranking tests.
 *
 * `contentAvailable` is set even though ranking never reads it: a `PairScore`
 * missing a required field type-checks nowhere and fails nothing here, because
 * `tests/**` is outside `tsconfig`. A fixture that is not a valid instance of
 * the type it claims to be is how a missing field becomes `undefined` in a
 * later test that does read it.
 */
function candidate(b: string, score: number, cast: number): PairScore {
  return { a: 'a1', b, content: score, filename: score, cast, contentAvailable: true, score };
}

describe('scorePair', () => {
  it('weights content at 0.8 and filename at 0.2, and nothing else', () => {
    const a = sketch('a1', 'Project Falcon SPA v6.docx', shared(10));
    const b = sketch('b1', 'Project Falcon SPA v7.docx', shared(10));

    const s = scorePair(a, b);
    expect(s.content).toBe(1);
    expect(s.filename).toBe(1);
    expect(s.score).toBeCloseTo(PAIRING_WEIGHT_CONTENT + PAIRING_WEIGHT_FILENAME, 10);
    expect(s.score).toBeCloseTo(1, 10);
  });

  it('leaves the cast out of the total, however high it is', () => {
    // Two unrelated documents from one deal bundle: identical cast, no shared
    // content, no shared name. The cast must not lift the score at all.
    const a = sketch('a1', 'Project Falcon SPA.docx', shared(10, 'spa'));
    const b = sketch('b1', 'Tax Deed.docx', shared(10, 'deed'));

    const s = scorePair(a, b);
    expect(s.cast).toBe(1);
    expect(s.content).toBe(0);
    expect(s.score).toBeCloseTo(PAIRING_WEIGHT_FILENAME * s.filename, 10);
    expect(s.score).toBeLessThan(0.2);
  });

  it('keeps its workings, so the grid can explain a pair rather than assert it', () => {
    const s = scorePair(
      sketch('a1', 'SPA v6.docx', shared(10)),
      sketch('b1', 'SPA v7.docx', [...shared(5), 'x', 'y']),
    );
    expect(s).toMatchObject({ a: 'a1', b: 'b1' });
    for (const key of ['content', 'filename', 'cast', 'score'] as const) {
      expect(typeof s[key]).toBe('number');
      expect(s[key]).toBeGreaterThanOrEqual(0);
      expect(s[key]).toBeLessThanOrEqual(1);
    }
  });

  it('is symmetric', () => {
    const a = sketch('a1', 'Project Falcon SPA v6.docx', shared(10));
    const b = sketch('b1', 'falcon-spa-clean.docx', [...shared(6), 'q']);
    const forward = scorePair(a, b);
    const back = scorePair(b, a);
    expect(back.score).toBeCloseTo(forward.score, 12);
    expect(back.content).toBeCloseTo(forward.content, 12);
  });
});

describe('an absent content signal is neither evidence of absence nor of presence', () => {
  it('falls back onto the filename when one side has no hashes', () => {
    // A PDF's text is deliberately not extracted for pairing, so its hash set
    // is empty. Scoring that as content 0 would cap the pair at 0.2 — below the
    // no-match threshold — and a Word and a PDF of the SAME agreement could
    // never be paired at all, which is the pairing M8 requires to be shown.
    const docx = sketch('a1', 'Project Falcon SPA.docx', shared(40));
    const pdf = sketch('b1', 'Project Falcon SPA.pdf', []);

    const s = scorePair(docx, pdf);
    expect(s.contentAvailable).toBe(false);
    expect(s.content).toBe(0);
    expect(s.filename).toBe(1);
    // Enough to clear the no-match threshold, and no more.
    expect(s.score).toBeCloseTo(DEFAULT_PAIRING_SETTINGS.noContentCeiling, 12);
    expect(s.score).toBeGreaterThan(PAIRING_NO_MATCH_SCORE);
  });

  it('cannot outrank a pairing that content actually corroborates', () => {
    // The defect this ceiling exists for. The first version renormalised
    // outright, so a perfectly-named PDF scored 1.0 — higher than any real
    // content-verified pair short of an exact match — won the global
    // assignment, and displaced the genuine `.docx` counterpart, which was then
    // shown as "no counterpart in the other bundle". Nothing was hidden, and
    // the proposal was still wrong.
    const docx = sketch('a1', 'Project Falcon SPA.docx', shared(40));
    const pdf = sketch('b1', 'Project Falcon SPA.pdf', []);

    // A real pair, only moderately corroborated: three fifths of its paragraphs
    // match and the names disagree entirely.
    const real = scorePair(
      sketch('a2', 'Project Falcon SPA v6.docx', [...shared(30), ...unique('a', 10)]),
      sketch('b2', 'Vendor draft 2026.docx', [...shared(30), ...unique('b', 10)]),
    );

    expect(real.contentAvailable).toBe(true);
    expect(scorePair(docx, pdf).score).toBeLessThan(real.score);
  });

  it('cannot stop a real pair auto-pairing just by being in the bundle', () => {
    // What the ceiling is actually FOR, once formation moved to `assign.ts`'s
    // second phase. `bandFor` measures its margin against the best rival, so an
    // uncapped filename-only rival at 1.000 would suppress auto-pair for a
    // perfect content-verified pair — adding an unrelated PDF to a bundle would
    // change the band of a pairing it has nothing to do with.
    const real = scorePair(
      sketch('a1', 'SPA v6.docx', shared(40)),
      sketch('b1', 'SPA v7.docx', shared(40)),
    );
    const pdfRival = scorePair(
      sketch('a1', 'SPA v6.docx', shared(40)),
      sketch('b2', 'SPA v6.pdf', []),
    );

    expect(pdfRival.filename).toBe(1);
    expect(real.score - pdfRival.score).toBeGreaterThan(0.15);
    expect(DEFAULT_PAIRING_SETTINGS.noContentCeiling).toBeLessThan(PAIRING_AUTO_PAIR_SCORE);
  });

  it('scores an unrelated name at nothing, whatever the formats', () => {
    // Whether such a pair FORMS is `assign.ts`'s second phase to decide, not
    // this number's — but the number still has to be honest, and two documents
    // with nothing in common in either signal have a score of nothing.
    const s = scorePair(
      sketch('a1', 'Project Falcon SPA.docx', shared(40)),
      sketch('b1', 'Vendor disclosure letter.pdf', []),
    );
    expect(s.contentAvailable).toBe(false);
    expect(s.score).toBeLessThan(PAIRING_NO_MATCH_SCORE);
  });

  it('marks the ordinary case as available, so the renormalisation is not the default', () => {
    const s = scorePair(
      sketch('a1', 'SPA v6.docx', shared(20)),
      sketch('b1', 'SPA v7.docx', shared(20)),
    );
    expect(s.contentAvailable).toBe(true);
    expect(s.score).toBeCloseTo(PAIRING_WEIGHT_CONTENT * 1 + PAIRING_WEIGHT_FILENAME * 1, 10);
  });

  it('still scores two nameless, contentless documents at 0', () => {
    // Nothing to go on at all must not become a confident match.
    const s = scorePair(sketch('a1', 'final.docx', []), sketch('b1', 'draft.docx', []));
    expect(s.contentAvailable).toBe(false);
    expect(s.score).toBe(0);
  });
});

describe('M5 — one shared cast must not decide a pairing', () => {
  it('pairs correctly across a bundle where every document shares the cast', () => {
    // The realistic bundle: SPA, disclosure letter and tax deed, all with the
    // same parties and defined terms, each present in two versions.
    const CAST = ['seller', 'buyer', 'completion date', 'purchase price'];
    const sideA = [
      sketch('a-spa', 'Project Falcon SPA v6.docx', shared(40, 'spa'), CAST),
      sketch('a-dl', 'Project Falcon Disclosure Letter v3.docx', shared(20, 'dl'), CAST),
      sketch('a-deed', 'Project Falcon Tax Deed v2.docx', shared(15, 'deed'), CAST),
    ];
    const sideB = [
      sketch('b-deed', 'Project Falcon Tax Deed (execution).docx', shared(15, 'deed'), CAST),
      sketch('b-spa', 'Project Falcon SPA (execution).docx', shared(38, 'spa'), CAST),
      sketch('b-dl', 'Project Falcon Disclosure Letter (execution).docx', shared(20, 'dl'), CAST),
    ];

    // Every candidate pair has cast overlap 1 — the signal is useless here.
    const all = scoreAll(sideA, sideB);
    expect(all.every(s => s.cast === 1)).toBe(true);

    // And every document's best candidate is still its own counterpart.
    for (const a of sideA) {
      const best = rankCandidates(all.filter(s => s.a === a.docId))[0];
      expect(best?.b).toBe(a.docId.replace(/^a-/, 'b-'));
    }
  });

  it('scores a right pair far above a wrong one even at cast 1', () => {
    const CAST = ['seller', 'buyer'];
    const a = sketch('a-spa', 'Project Falcon SPA v6.docx', shared(40, 'spa'), CAST);
    const right = sketch('b-spa', 'Project Falcon SPA v7.docx', shared(38, 'spa'), CAST);
    const wrong = sketch('b-dl', 'Project Falcon Disclosure Letter v3.docx', shared(20, 'dl'), CAST);

    expect(scorePair(a, right).score).toBeGreaterThan(scorePair(a, wrong).score + 0.5);
  });
});

describe('rankCandidates — the cast is consulted only inside the tie band', () => {
  it('ignores the cast entirely when the scores are apart', () => {
    // The lower-scoring candidate has the better cast. It must still lose:
    // outside the epsilon the cast is not consulted at all.
    const ranked = rankCandidates([
      candidate('low-but-familiar', 0.60, 1.0),
      candidate('high', 0.90, 0.1),
    ]);
    expect(ranked.map(c => c.b)).toEqual(['high', 'low-but-familiar']);
  });

  it('uses the cast to order two candidates that are level', () => {
    const ranked = rankCandidates([
      candidate('level-weak-cast', 0.80, 0.2),
      candidate('level-strong-cast', 0.801, 0.9),
      candidate('level-mid-cast', 0.799, 0.5),
    ]);
    // All three are within PAIRING_TIE_EPSILON of the best, so the cast orders
    // them — and it is the ONLY thing that could have.
    expect(ranked.map(c => c.b))
      .toEqual(['level-strong-cast', 'level-mid-cast', 'level-weak-cast']);
  });

  it('applies the cast to the leading band only, never below it', () => {
    const ranked = rankCandidates([
      candidate('band-a', 0.90, 0.1),
      candidate('band-b', 0.895, 0.9),
      candidate('below-with-best-cast', 0.50, 1.0),
    ]);
    expect(ranked.map(c => c.b)).toEqual(['band-b', 'band-a', 'below-with-best-cast']);
  });

  it('is a function of the candidate SET, not of input order', () => {
    // The property M4 demands of the assignment, asserted one level down. A
    // pairwise "is this neighbour within epsilon" comparator is NOT transitive
    // — 0.10/0.11/0.12 at epsilon 0.02 ties the adjacent pairs but not the ends
    // — and would sort differently for different permutations.
    const set = [
      candidate('p', 0.100, 0.4),
      candidate('q', 0.110, 0.9),
      candidate('r', 0.120, 0.1),
      candidate('s', 0.400, 0.5),
    ];

    const first = rankCandidates(set).map(c => c.b);
    for (const permutation of [
      [3, 2, 1, 0], [1, 3, 0, 2], [2, 0, 3, 1], [0, 2, 1, 3],
    ]) {
      const shuffled = permutation.map(i => set[i] as PairScore);
      expect(rankCandidates(shuffled).map(c => c.b)).toEqual(first);
    }
  });

  it('handles no candidates and one candidate', () => {
    expect(rankCandidates([])).toEqual([]);
    expect(rankCandidates([candidate('only', 0.3, 0.2)]).map(c => c.b)).toEqual(['only']);
  });

  it('drops nothing — confidence orders the grid, it never filters it', () => {
    const set = [
      candidate('strong', 0.95, 0.5),
      candidate('weak', 0.02, 0.5),
      candidate('hopeless', 0.0, 0.0),
    ];
    expect(rankCandidates(set)).toHaveLength(3);
  });
});

describe('isTie', () => {
  it('is exactly the epsilon band, exclusive at the boundary', () => {
    expect(isTie(candidate('x', 0.80, 0), candidate('y', 0.80, 0))).toBe(true);
    expect(isTie(candidate('x', 0.80, 0), candidate('y', 0.80 + PAIRING_TIE_EPSILON / 2, 0)))
      .toBe(true);
    expect(isTie(candidate('x', 0.80, 0), candidate('y', 0.80 + PAIRING_TIE_EPSILON, 0)))
      .toBe(false);
  });
});
