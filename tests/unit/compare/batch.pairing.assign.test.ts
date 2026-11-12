/**
 * Hungarian assignment with no-match padding (ADR-0037; spec §6.4;
 * acceptance M1, M3, M4).
 *
 * This is the module where a wrong answer is least visible. A mis-pairing does
 * not throw and does not look broken — it produces a complete, well-formed,
 * fully-populated redline of one document against an unrelated one, which a
 * reader has no cue to distrust. So the tests here are about the three
 * properties that separate an optimal assignment from a plausible one:
 *
 *  - M1 — it recovers a known pairing exactly;
 *  - M3 — surplus documents come out unmatched rather than force-matched;
 *  - M4 — shuffling both bundles changes nothing, which is the property a
 *    greedy "everyone takes their favourite" pass fails.
 */
import { describe, it, expect } from 'vitest';

import { hungarian } from '@features/compare/batch/pairing/hungarian.js';
import { assignPairs, noMatchCost } from '@features/compare/batch/pairing/assign.js';
import { PAIRING_NO_MATCH_SCORE } from '@features/compare/batch/constants.js';
import { DEFAULT_PAIRING_SETTINGS } from '@features/compare/batch/config.js';
import type { PairingSketch } from '@features/compare/batch/pairing/sketch.js';

function sketch(
  docId: string,
  fileName: string,
  hashes: readonly string[],
): PairingSketch {
  return {
    docId,
    side: docId.startsWith('a') ? 'a' : 'b',
    sourceKind: 'docx',
    fileName,
    byteLength: 1024,
    paraHashes: new Set(hashes),
    definedTerms: new Set(['seller', 'buyer']),
    partyNames: new Set(['seller', 'buyer']),
    quarantinedCount: 0,
  };
}

/** `n` paragraph hashes for one document, `drift` of them replaced. */
function body(prefix: string, n: number, drift = 0): string[] {
  return Array.from({ length: n }, (_, i) =>
    i < drift ? `${prefix}-new-${i}` : `${prefix}-${i}`);
}

const totalCost = (cost: readonly (readonly number[])[], rowToCol: readonly number[]): number =>
  rowToCol.reduce((sum, j, i) => sum + ((cost[i] as readonly number[])[j] as number), 0);

describe('hungarian — the solver itself', () => {
  it('finds the minimum, not a good-enough greedy answer', () => {
    // The trap: the globally cheapest cell is (0,0) at 1, and taking it strands
    // row 1 on the 100. Greedy pays 1 + 100 = 101; the optimum gives up the
    // cheapest cell and pays 4 + 2 = 6. This is the pairing case where two
    // documents both rank the same counterpart first.
    const cost = [
      [1, 4],
      [2, 100],
    ];
    const assignment = hungarian(cost);
    expect(totalCost(cost, assignment)).toBe(6);
    expect(assignment).toEqual([1, 0]);
  });

  it('assigns every row exactly once and no column twice', () => {
    const cost = [
      [7, 2, 9, 4],
      [3, 8, 1, 6],
      [5, 4, 6, 2],
      [9, 7, 3, 8],
    ];
    const assignment = hungarian(cost);
    expect(assignment).toHaveLength(4);
    expect(new Set(assignment).size).toBe(4);
    expect(assignment.every(j => j >= 0 && j < 4)).toBe(true);
  });

  it('beats every permutation on a matrix small enough to brute-force', () => {
    const cost = [
      [8, 4, 7, 3, 6],
      [2, 9, 5, 8, 1],
      [6, 3, 9, 4, 7],
      [5, 7, 2, 6, 9],
      [9, 1, 8, 5, 4],
    ];
    const best = totalCost(cost, hungarian(cost));

    const permute = (rest: number[], acc: number[] = []): number[][] =>
      rest.length === 0
        ? [acc]
        : rest.flatMap((x, i) =>
            permute([...rest.slice(0, i), ...rest.slice(i + 1)], [...acc, x]));

    for (const p of permute([0, 1, 2, 3, 4])) {
      expect(best).toBeLessThanOrEqual(totalCost(cost, p));
    }
  });

  it('handles a rectangular matrix wider than it is tall', () => {
    const cost = [
      [5, 1, 9],
      [8, 6, 2],
    ];
    const assignment = hungarian(cost);
    expect(totalCost(cost, assignment)).toBe(3);
  });

  it('handles the empty matrix and refuses one too narrow to solve', () => {
    expect(hungarian([])).toEqual([]);
    expect(() => hungarian([[1, 2], [3, 4], [5, 6]])).toThrow(/columns/i);
  });
});

describe('the no-match cost is derived, not chosen', () => {
  it('is exactly the value that makes the pairing threshold PAIRING_NO_MATCH_SCORE', () => {
    // Pairing costs (1 − score); leaving both unmatched costs 2·NO_MATCH_COST.
    // The solver prefers the pair iff 1 − score < 2·NO_MATCH_COST, and that
    // must come out as score > PAIRING_NO_MATCH_SCORE.
    expect(1 - 2 * noMatchCost()).toBeCloseTo(PAIRING_NO_MATCH_SCORE, 12);
  });

  it('pairs just above the threshold and drops just below it', () => {
    // Content-only, so the score is the Jaccard exactly: 10 shared of 20
    // distinct is 0.333, comfortably above 0.25; 2 of 20 is 0.1, below it.
    const above = assignPairs(
      [sketch('a1', 'x.docx', body('p', 20))],
      [sketch('b1', 'y.docx', [...body('p', 20).slice(0, 10), ...body('q', 10)])],
    );
    expect(above.pairs).toHaveLength(1);

    const below = assignPairs(
      [sketch('a1', 'x.docx', body('p', 20))],
      [sketch('b1', 'y.docx', [...body('p', 20).slice(0, 2), ...body('q', 18)])],
    );
    expect(below.pairs).toHaveLength(0);
    expect(below.unmatchedA).toEqual(['a1']);
    expect(below.unmatchedB).toEqual(['b1']);
  });
});

describe('M1 — recovers a known pairing exactly', () => {
  const sideA = [
    sketch('a-spa', 'Project Falcon SPA v6.docx', body('spa', 60)),
    sketch('a-dl', 'Project Falcon Disclosure Letter v3.docx', body('dl', 30)),
    sketch('a-deed', 'Project Falcon Tax Deed v2.docx', body('deed', 25)),
  ];
  const sideB = [
    sketch('b-deed', 'Project Falcon Tax Deed (execution).docx', body('deed', 25, 2)),
    sketch('b-spa', 'Project Falcon SPA (execution).docx', body('spa', 60, 5)),
    sketch('b-dl', 'Project Falcon Disclosure Letter (execution).docx', body('dl', 30, 3)),
  ];

  it('pairs each document with its own counterpart', () => {
    const { pairs, unmatchedA, unmatchedB } = assignPairs(sideA, sideB);

    expect(pairs).toHaveLength(3);
    expect(unmatchedA).toEqual([]);
    expect(unmatchedB).toEqual([]);
    expect(pairs.map(p => `${p.a}→${p.b}`).sort())
      .toEqual(['a-deed→b-deed', 'a-dl→b-dl', 'a-spa→b-spa']);
  });

  it('keeps every candidate score, not just the winners', () => {
    // The grid shows rivals. An assignment that discarded the losing scores
    // could propose a pair but not explain why it beat the alternatives.
    const { scores } = assignPairs(sideA, sideB);
    expect(scores).toHaveLength(9);
  });

  it('recovers the pairing even where a greedy pass would not', () => {
    // Both A documents rank `b-shared` first, but only one can have it. Greedy
    // gives it to whichever it visits first and strands the other; the optimal
    // assignment trades it for the better total.
    const a1 = sketch('a-1', 'one.docx', [...body('common', 18), ...body('one', 6)]);
    const a2 = sketch('a-2', 'two.docx', [...body('common', 18), ...body('two', 2)]);
    const shared = sketch('b-shared', 'shared.docx', [...body('common', 18), ...body('one', 5)]);
    const other = sketch('b-other', 'other.docx', [...body('common', 18), ...body('two', 2)]);

    const { pairs } = assignPairs([a1, a2], [shared, other]);
    expect(pairs).toHaveLength(2);
    expect(new Set(pairs.map(p => p.b)).size).toBe(2);
  });
});

describe('M3 — surplus documents fall out instead of being force-matched', () => {
  it('leaves an added document unmatched when B is larger', () => {
    const sideA = [sketch('a-spa', 'SPA v6.docx', body('spa', 40))];
    const sideB = [
      sketch('b-spa', 'SPA v7.docx', body('spa', 40, 3)),
      sketch('b-side', 'Side Letter.docx', body('side', 12)),
    ];

    const { pairs, unmatchedA, unmatchedB } = assignPairs(sideA, sideB);
    expect(pairs.map(p => `${p.a}→${p.b}`)).toEqual(['a-spa→b-spa']);
    expect(unmatchedA).toEqual([]);
    expect(unmatchedB).toEqual(['b-side']);
  });

  it('leaves a removed document unmatched when A is larger', () => {
    const sideA = [
      sketch('a-spa', 'SPA v6.docx', body('spa', 40)),
      sketch('a-sched', 'Schedule 4.docx', body('sched', 12)),
    ];
    const sideB = [sketch('b-spa', 'SPA v7.docx', body('spa', 40, 3))];

    const { pairs, unmatchedA, unmatchedB } = assignPairs(sideA, sideB);
    expect(pairs.map(p => `${p.a}→${p.b}`)).toEqual(['a-spa→b-spa']);
    expect(unmatchedA).toEqual(['a-sched']);
    expect(unmatchedB).toEqual([]);
  });

  it('leaves BOTH unmatched when a document was replaced by a different one', () => {
    // Equal-sized bundles. Without the dummy padding this is the case that
    // silently forces a pairing, because there is nothing else left to take.
    const sideA = [
      sketch('a-spa', 'SPA v6.docx', body('spa', 40)),
      sketch('a-old', 'Escrow Deed.docx', body('escrow', 20)),
    ];
    const sideB = [
      sketch('b-spa', 'SPA v7.docx', body('spa', 40, 3)),
      sketch('b-new', 'Tax Deed.docx', body('tax', 20)),
    ];

    const { pairs, unmatchedA, unmatchedB } = assignPairs(sideA, sideB);
    expect(pairs.map(p => `${p.a}→${p.b}`)).toEqual(['a-spa→b-spa']);
    expect(unmatchedA).toEqual(['a-old']);
    expect(unmatchedB).toEqual(['b-new']);
  });

  it('matches nothing when one side is empty', () => {
    const only = [sketch('a-spa', 'SPA.docx', body('spa', 10))];
    expect(assignPairs(only, [])).toMatchObject({
      pairs: [], unmatchedA: ['a-spa'], unmatchedB: [],
    });
    expect(assignPairs([], only)).toMatchObject({
      pairs: [], unmatchedA: [], unmatchedB: ['a-spa'],
    });
    expect(assignPairs([], [])).toMatchObject({ pairs: [], unmatchedA: [], unmatchedB: [] });
  });
});

describe('the two phases, and what separates them', () => {
  it('a pair with no content evidence is REFUSED by phase 1, not merely outbid', () => {
    // Phase 1 prices a no-content cell strictly above the cost of leaving both
    // its documents unmatched — `2·dummy + 1`, not `2·dummy`. At exactly the
    // dummy price the two options tie, and a tie is decided by whatever the
    // solver happens to do, which for a pairing means "sometimes force-matches
    // two unrelated documents". Here the names are unrelated too, so phase 2
    // will not take the pair either: the only correct answer is unmatched.
    const result = assignPairs(
      [sketch('a1', 'Escrow Deed.docx', body('escrow', 40))],
      [sketch('b1', 'Zoning Permit Schedule.docx', [])],
    );

    expect(result.pairs).toEqual([]);
    expect(result.unmatchedA).toEqual(['a1']);
    expect(result.unmatchedB).toEqual(['b1']);
  });

  it('a filename-only candidate cannot take a document from a pair content supports', () => {
    // The claim the phase split exists for, made where the split lives. The
    // pair is deliberately WEAK — 14 shared hashes out of 40 — so its
    // score sits in the narrow window above the no-match price where a
    // perfectly-named no-content candidate (capped at `noContentCeiling`) would
    // outscore it on a single scale. Phase 1 does not admit that candidate at
    // all, so the question never arises.
    const overlap = body('shared', 14);
    const result = assignPairs(
      [sketch('a1', 'facility agreement 2026.docx', [...overlap, ...body('a', 13)])],
      [
        sketch('b1', 'revised banking terms.docx', [...overlap, ...body('b', 13)]),
        sketch('b2', 'facility agreement 2026.pdf', []),
      ],
    );

    // The fixture asserts its own score first: drifted out of the window, this
    // test would pass while proving nothing.
    const real = result.scores.find(x => x.b === 'b1')?.score as number;
    expect(real).toBeGreaterThan(PAIRING_NO_MATCH_SCORE);
    expect(real).toBeLessThanOrEqual(DEFAULT_PAIRING_SETTINGS.noContentCeiling);

    expect(result.pairs.map(p => `${p.a}::${p.b}`)).toEqual(['a1::b1']);
    expect(result.unmatchedB).toEqual(['b2']);
  });

  it('emits pairs in docId order, whichever phase placed them', () => {
    // `a1` can only be placed by phase 2 (its counterpart has no content) and
    // `a2` only by phase 1, so phase order and docId order disagree. M4 wants
    // the output to be a function of the two input SETS, and "which phase got
    // there first" is not part of either set.
    const result = assignPairs(
      [
        sketch('a1', 'Tax Deed.docx', body('tax', 40)),
        sketch('a2', 'Share Purchase Agreement.docx', body('spa', 40)),
      ],
      [
        sketch('b1', 'Share Purchase Agreement v2.docx', body('spa', 40, 4)),
        sketch('b2', 'Tax Deed of Covenant.docx', []),
      ],
    );

    expect(result.pairs.map(p => `${p.a}::${p.b}`)).toEqual(['a1::b2', 'a2::b1']);
  });
});

describe('M4 — shuffling both bundles gives an identical assignment', () => {
  const sideA = [
    sketch('a-spa', 'Project Falcon SPA v6.docx', body('spa', 60)),
    sketch('a-dl', 'Disclosure Letter v3.docx', body('dl', 30)),
    sketch('a-deed', 'Tax Deed v2.docx', body('deed', 25)),
    sketch('a-sched', 'Schedule 4.docx', body('sched', 12)),
  ];
  const sideB = [
    sketch('b-spa', 'Project Falcon SPA (execution).docx', body('spa', 60, 5)),
    sketch('b-dl', 'Disclosure Letter (execution).docx', body('dl', 30, 3)),
    sketch('b-deed', 'Tax Deed (execution).docx', body('deed', 25, 2)),
    sketch('b-side', 'Side Letter.docx', body('side', 15)),
  ];

  const shuffle = <T,>(xs: readonly T[], seed: number): T[] => {
    const out = [...xs];
    let state = seed;
    for (let i = out.length - 1; i > 0; i -= 1) {
      state = (state * 1103515245 + 12345) % 2147483648;
      const j = state % (i + 1);
      [out[i], out[j]] = [out[j] as T, out[i] as T];
    }
    return out;
  };

  it('is a function of the two SETS, not of their order', () => {
    const baseline = assignPairs(sideA, sideB);
    const key = (r: ReturnType<typeof assignPairs>) => ({
      pairs: [...r.pairs.map(p => `${p.a}→${p.b}`)].sort(),
      unmatchedA: [...r.unmatchedA].sort(),
      unmatchedB: [...r.unmatchedB].sort(),
    });

    for (let seed = 1; seed <= 12; seed += 1) {
      const shuffled = assignPairs(shuffle(sideA, seed), shuffle(sideB, seed * 7));
      expect(key(shuffled)).toEqual(key(baseline));
    }
  });

  it('emits its pairs in a stable order too, not merely a stable set', () => {
    // The grid renders this list. A stable set with an unstable order would
    // still reshuffle the rows under the user between two identical runs.
    const first = assignPairs(sideA, sideB).pairs.map(p => p.a);
    for (let seed = 1; seed <= 6; seed += 1) {
      expect(assignPairs(shuffle(sideA, seed), shuffle(sideB, seed + 3)).pairs.map(p => p.a))
        .toEqual(first);
    }
  });
});
