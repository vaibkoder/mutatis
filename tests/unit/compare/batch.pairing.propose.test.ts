/**
 * The pairing proposal (ADR-0037; spec §6.5, §6.6; acceptance M2, M8).
 *
 * Two rules are under test, and both are about what the user is allowed NOT to
 * see — which is nothing:
 *
 *  - **M2** — a near-identical rival must stop a pair being pre-ticked. Two
 *    consecutive drafts of one agreement both score ~0.97 against a third, so a
 *    score-only rule would auto-pair a coin toss. The margin is what catches it.
 *  - **M8** — a pair that cannot be compared is still paired, still scored, and
 *    still shown with its rivals. Dropping it would hide a document the user
 *    handed us behind a limitation of this build.
 *
 * And underneath both: a band is a badge. No threshold anywhere in this module
 * may remove a row from the grid or a rival from a row.
 */
import { describe, it, expect } from 'vitest';

import { bandFor, proposePairing } from '@features/compare/batch/pairing/propose.js';
import { tallyPairing } from '@features/compare/batch/pairing/model.js';
import { MAX_DOCX_BYTES } from '@features/compare/batch/limits.js';
import {
  PAIRING_AUTO_PAIR_MARGIN,
  PAIRING_AUTO_PAIR_SCORE,
  PAIRING_NO_MATCH_SCORE,
} from '@features/compare/batch/constants.js';
import { DEFAULT_PAIRING_SETTINGS } from '@features/compare/batch/config.js';
import { getActiveProfile, getCompareConfig } from '@shared/config/index.js';
import type { PairingSketch } from '@features/compare/batch/pairing/sketch.js';
import type { PairScore } from '@features/compare/batch/pairing/score.js';

const AT = '2026-09-10T09:00:00.000Z';
const BASE_CONFIG = getCompareConfig(getActiveProfile());

function sketch(
  docId: string,
  fileName: string,
  hashes: readonly string[],
  over: Partial<PairingSketch> = {},
): PairingSketch {
  return {
    docId,
    side: docId.startsWith('a') ? 'a' : 'b',
    sourceKind: 'docx',
    fileName,
    byteLength: 4096,
    paraHashes: new Set(hashes),
    definedTerms: new Set(['seller', 'buyer']),
    partyNames: new Set(['seller', 'buyer']),
    quarantinedCount: 0,
    ...over,
  };
}

const body = (prefix: string, n: number, drift = 0): string[] =>
  Array.from({ length: n }, (_, i) => (i < drift ? `${prefix}-new-${i}` : `${prefix}-${i}`));

const score = (s: number, contentAvailable = true): PairScore =>
  ({ a: 'a', b: 'b', content: s, filename: s, cast: 0.5, contentAvailable, score: s });

describe('bandFor — the two conditions, and why both are needed', () => {
  it('auto-pairs only on a high score AND a clear lead', () => {
    const high = score(PAIRING_AUTO_PAIR_SCORE + 0.05);
    const distantRival = score(high.score - PAIRING_AUTO_PAIR_MARGIN - 0.01);
    expect(bandFor(high, distantRival.score)).toBe('auto-pair');
  });

  it('refuses to auto-pair a high score with a close rival — M2 in miniature', () => {
    // The whole point. Two drafts of one agreement both score ~0.97 against a
    // third; a score-only rule pre-ticks whichever happened to win by 0.001.
    const high = score(0.97);
    const closeRival = score(0.96);
    expect(bandFor(high, closeRival.score)).toBe('propose-with-rivals');
  });

  it('refuses to auto-pair a clear lead that is not actually a good match', () => {
    const mediocre = score(0.40);
    expect(bandFor(mediocre, 0.05)).toBe('propose-with-rivals');
  });

  it('auto-pairs an uncontested strong match', () => {
    expect(bandFor(score(0.99), null)).toBe('auto-pair');
  });

  it('never auto-pairs on a filename alone, however perfect the name match', () => {
    // With no content signal there is nothing to catch a mislabelled file, and
    // a mislabelled file is the ordinary case filenames fail on.
    expect(bandFor(score(1, false), null)).toBe('propose-with-rivals');
  });

  it('treats the boundaries as inclusive, so the constants mean what they say', () => {
    const at = score(PAIRING_AUTO_PAIR_SCORE);
    expect(bandFor(at, PAIRING_AUTO_PAIR_SCORE - PAIRING_AUTO_PAIR_MARGIN))
      .toBe('auto-pair');
    expect(bandFor(score(PAIRING_AUTO_PAIR_SCORE - 0.001), null)).toBe('propose-with-rivals');
  });
});

describe('the cast cannot decide a band — guardrail 3, as an invariant', () => {
  // The M5 test proves the cast is absent from the weighted SUM. That is not
  // the same claim as this one, and two reviews have now found the gap by
  // exploiting it: `rankCandidates` re-orders rivals by cast inside the tie
  // band, so which rivals a row DISPLAYS moves with the cast.
  //
  // Inside one deal bundle the SPA, the disclosure letter and the side letter
  // share the same parties and most defined terms. A signal that near-useless
  // must not decide whether a pairing arrives pre-ticked.
  it('measures the margin from the strongest rival, not the first one listed', () => {
    // `bandFor` takes a NUMBER, so the only way the cast could reach it is via
    // the call site. That is what the proposal-level test below covers.
    const pair = score(0.90);
    expect(bandFor(pair, 0.80)).toBe('propose-with-rivals');
  });

  it('still auto-pairs when every rival is genuinely far behind', () => {
    // The control: the guard above must not have made auto-pair unreachable.
    expect(bandFor(score(0.95), 0.20)).toBe('auto-pair');
  });

  it('gives the same band whichever way the cast orders THREE level rivals', () => {
    // Through the real composition, and with three rivals rather than two.
    //
    // The first fix for this measured the margin over the rival list — but the
    // list the call site had was the two rivals the ROW SHOWS. With three
    // candidates inside the tie band, the cast decides which two those are, so
    // the strongest rival can be pushed off the row and the margin measured
    // from the second-strongest. Two rivals hid it: with only two, the
    // displayed pair IS the whole list however the cast orders it.
    //
    // Here the pair scores 0.9200 and the three rivals 0.4133 / 0.4067 /
    // 0.4000, all inside the default tie epsilon. The margin is set between
    // 0.9200 − 0.4133 and 0.9200 − 0.4067, so measuring from the wrong rival
    // flips the band.
    const CAST = ['seller', 'buyer'];
    const withCast = (names: readonly string[]): Partial<PairingSketch> =>
      ({ definedTerms: new Set(names), partyNames: new Set(names) });

    const proposalWith = (b2: readonly string[], b3: readonly string[], b4: readonly string[]) =>
      proposePairing({
        sideA: [sketch('a1', 'Share Purchase Agreement.docx', body('spa', 100), withCast(CAST))],
        sideB: [
          sketch('b1', 'Share Purchase Agreement v2.docx', body('spa', 90)),
          sketch('b2', 'Tax Deed.docx', [...body('spa', 62), ...body('z', 20)], withCast(b2)),
          sketch('b3', 'Escrow Deed.docx', [...body('spa', 61), ...body('y', 20)], withCast(b3)),
          sketch('b4', 'Warranty Schedule.docx', [...body('spa', 60), ...body('x', 20)], withCast(b4)),
        ],
        generatedAt: AT,
        config: { ...BASE_CONFIG, batch: { pairing: { autoPairMargin: 0.51 } } },
      }).rows.find(r => r.a === 'a1');

    const strongestShown = proposalWith(CAST, [], []);
    const strongestHidden = proposalWith([], CAST, CAST);

    // The cast really did change which rivals are shown — otherwise this
    // proves nothing.
    expect(strongestShown?.rivals.map(r => r.b))
      .not.toEqual(strongestHidden?.rivals.map(r => r.b));
    expect(strongestHidden?.rivals.map(r => r.b)).not.toContain('b2');

    expect(strongestHidden?.band).toBe(strongestShown?.band);
    expect(strongestShown?.band).toBe('propose-with-rivals');
  });
});

describe('M2 — the deliberately near-identical pair', () => {
  // One agreement on side A; side B holds BOTH v6 and v7 of it. Whichever the
  // solver picks, the other is a near-identical rival and the pair must not be
  // pre-ticked.
  const spa = body('spa', 60);
  const sideA = [sketch('a-spa', 'Project Falcon SPA.docx', spa)];
  const sideB = [
    sketch('b-v6', 'Project Falcon SPA v6.docx', body('spa', 60, 2)),
    sketch('b-v7', 'Project Falcon SPA v7.docx', body('spa', 60, 3)),
  ];

  it('lands in propose-with-rivals, not auto-pair', () => {
    const proposal = proposePairing({ sideA, sideB, generatedAt: AT });
    const paired = proposal.rows.filter(r => r.band !== 'unmatched');

    expect(paired).toHaveLength(1);
    expect(paired[0]?.band).toBe('propose-with-rivals');
  });

  it('carries the rival WITH its score, so the choice can be checked', () => {
    const row = proposePairing({ sideA, sideB, generatedAt: AT })
      .rows.find(r => r.band === 'propose-with-rivals');

    expect(row?.rivals.length).toBeGreaterThanOrEqual(1);
    const rival = row?.rivals[0] as PairScore;
    expect(typeof rival.score).toBe('number');
    expect(rival.score).toBeGreaterThan(0.8);
    // The rival is the OTHER version, not the one that was taken.
    expect(rival.b).not.toBe(row?.b);
    // And it really is close — which is precisely why this is not auto-pair.
    expect(Math.abs((row?.score as PairScore).score - rival.score)).toBeLessThan(0.15);
  });

  it('shows the version that lost as an unmatched row rather than dropping it', () => {
    const proposal = proposePairing({ sideA, sideB, generatedAt: AT });
    const unmatched = proposal.rows.filter(r => r.band === 'unmatched');

    expect(unmatched).toHaveLength(1);
    expect(unmatched[0]?.a).toBeNull();
    expect(unmatched[0]?.reason).toMatch(/added, or replacing/i);
    // Every document from both bundles appears exactly once.
    expect(proposal.rows).toHaveLength(2);
    expect(proposal.totals.documentsA).toBe(1);
    expect(proposal.totals.documentsB).toBe(2);
  });

  it('auto-pairs the same document when the near-identical rival is absent', () => {
    // The control. Without the rival the pair is unambiguous, which proves the
    // `propose` above is caused by the margin and not by a weak score.
    const proposal = proposePairing({
      sideA,
      sideB: [sketch('b-v7', 'Project Falcon SPA v7.docx', body('spa', 60, 3))],
      generatedAt: AT,
    });
    expect(proposal.rows.find(r => r.band !== 'unmatched')?.band).toBe('auto-pair');
  });
});

describe('M8 — a pair that cannot be compared is still shown, with its pairing', () => {
  it('surfaces a cross-format pair as not-runnable, never as an absence', () => {
    const proposal = proposePairing({
      sideA: [sketch('a-spa', 'Project Falcon SPA.docx', body('spa', 40))],
      sideB: [sketch('b-spa', 'Project Falcon SPA.pdf', [], {
        sourceKind: 'pdf',
        paraHashes: new Set<string>(),
      })],
      generatedAt: AT,
    });

    const row = proposal.rows.find(r => r.a !== null && r.b !== null);
    expect(row).toBeDefined();
    expect(row?.runnable).toBe(false);
    expect(row?.blocked).toBe('cross-format');
    expect(row?.reason).toMatch(/out of scope/i);
    // Paired and SCORED — the pairing is shown, not withheld.
    expect(row?.score).not.toBeNull();
    expect(proposal.totals.unrunnable).toBe(1);
    expect(proposal.disclosures.some(d => /mix formats/i.test(d))).toBe(true);
  });

  it('surfaces an oversize pair with the cap named', () => {
    const proposal = proposePairing({
      sideA: [sketch('a-spa', 'huge SPA.docx', body('spa', 40), {
        byteLength: MAX_DOCX_BYTES + 1,
      })],
      sideB: [sketch('b-spa', 'huge SPA v2.docx', body('spa', 40, 2))],
      generatedAt: AT,
    });

    const row = proposal.rows.find(r => r.a !== null && r.b !== null);
    expect(row?.runnable).toBe(false);
    expect(row?.blocked).toBe('cap');
    expect(row?.reason).toMatch(/MiB/);
    expect(row?.reason).toMatch(/huge SPA\.docx/);
    expect(proposal.disclosures.some(d => /larger than this build/i.test(d))).toBe(true);
  });

  it('reports cross-format ahead of a size cap', () => {
    // Naming a size cap on a pair that was never comparable anyway would send
    // the user off to shrink a file that would still not be compared.
    const proposal = proposePairing({
      sideA: [sketch('a-spa', 'SPA.pdf', [], {
        sourceKind: 'pdf', byteLength: MAX_DOCX_BYTES + 1, paraHashes: new Set<string>(),
      })],
      sideB: [sketch('b-spa', 'SPA.docx', [])],
      generatedAt: AT,
    });
    expect(proposal.rows[0]?.blocked).toBe('cross-format');
  });

  it('leaves a comparable pair runnable and undisclosed', () => {
    const proposal = proposePairing({
      sideA: [sketch('a-spa', 'SPA v6.docx', body('spa', 40))],
      sideB: [sketch('b-spa', 'SPA v7.docx', body('spa', 40, 2))],
      generatedAt: AT,
    });
    expect(proposal.rows[0]?.runnable).toBe(true);
    expect(proposal.rows[0]?.blocked).toBeUndefined();
    expect(proposal.disclosures).toEqual([]);
  });
});

describe('a filename-only pair does not displace a real one (X-3)', () => {
  // Found by review, on exactly the mixed bundle the fallback was written for.
  // A `.pdf` carries no paragraph hashes, so its score rests on its filename
  // alone; when that fallback renormalised onto the full range, a perfectly
  // named PDF scored 1.000, won the GLOBAL assignment, and took the side-A
  // document away from its genuine `.docx` counterpart. The displaced document
  // was then shown as "no counterpart in the other bundle — added, or
  // replacing", which was not true.
  //
  // The band guard was not enough: it correctly refuses to pre-tick a
  // filename-only pair, but it runs downstream of the assignment, and the
  // assignment is where the harm was done.
  const spa = body('spa', 40);
  const spaDrifted = body('spa', 40, 4);

  function mixed() {
    return proposePairing({
      sideA: [sketch('a1', 'Project Falcon SPA v6.docx', spa)],
      sideB: [
        sketch('b1', 'Project Falcon SPA v7.docx', spaDrifted),
        sketch('b2', 'Project Falcon SPA.pdf', [], { sourceKind: 'pdf' }),
      ],
      generatedAt: AT,
    });
  }

  it('pairs the docx with the docx, not with the identically-named pdf', () => {
    const paired = mixed().rows.find(r => r.a === 'a1' && r.b !== null);
    expect(paired?.b).toBe('b1');
    expect(paired?.runnable).toBe(true);
  });

  it('leaves the pdf as its own row, with the pairing it lost', () => {
    // Not dropped — M8's substance survives the fix. The PDF is still on the
    // grid, and its row still names the document it was a candidate for.
    const pdf = mixed().rows.find(r => r.b === 'b2');
    expect(pdf).toBeDefined();
    expect(pdf?.band).toBe('unmatched');
    expect(pdf?.rivals.map(r => r.a)).toContain('a1');
  });

  it('still forms the cross-format pair when nothing else wants either document', () => {
    // The other half, and the one a first attempt at this fix broke. Scoring
    // filename-only evidence low enough not to displace also stopped it forming
    // any pair unless the normalised names were near-identical — which fails M8
    // from the opposite direction. Formation is now a second assignment phase,
    // so it does not depend on where the score sits at all.
    const only = proposePairing({
      sideA: [sketch('a1', 'Project Falcon SPA v6.docx', spa)],
      sideB: [sketch('b2', 'Project Falcon SPA.pdf', [], { sourceKind: 'pdf' })],
      generatedAt: AT,
    });

    const row = only.rows.find(r => r.a === 'a1');
    expect(row?.b).toBe('b2');
    expect(row?.runnable).toBe(false);
    expect(row?.reason).toMatch(/cross-format/i);
    expect(only.totals.unrunnable).toBe(1);
  });

  it('forms it even when the two filenames are only loosely alike', () => {
    // Spec §1's own motivating example, which a review used to show the first
    // attempt was too strict: these normalise to "project falcon spa" against
    // "spa" — a Dice of 0.26. Weak, but the alternative is telling the user
    // there is no counterpart while the counterpart sits in the other bundle.
    const loose = proposePairing({
      sideA: [sketch('a1', 'Project_Falcon_SPA_v7_exec.docx', spa)],
      sideB: [sketch('b1', 'SPA (clean) 2026-08-30.pdf', [], { sourceKind: 'pdf' })],
      generatedAt: AT,
    });

    const row = loose.rows.find(r => r.a === 'a1');
    expect(row?.b).toBe('b1');
    expect(row?.band).toBe('propose-with-rivals');
    expect(row?.runnable).toBe(false);
    expect(loose.totals.unrunnable).toBe(1);
  });

  it('does NOT invent a pair out of two unrelated names', () => {
    // The control for the loose bar above. The second phase asks whether there
    // is any name evidence at all, and here there is none.
    const unrelated = proposePairing({
      sideA: [sketch('a1', 'Project Falcon SPA v6.docx', spa)],
      sideB: [sketch('b1', 'Board minutes March.pdf', [], { sourceKind: 'pdf' })],
      generatedAt: AT,
    });

    expect(unrelated.totals.pairs).toBe(0);
    expect(unrelated.totals.unmatchedA).toBe(1);
    expect(unrelated.totals.unmatchedB).toBe(1);
  });

  it('says why an unreadable document has no counterpart, instead of claiming it was added', () => {
    // "Added, or replacing" is a claim about the BUNDLE, and one the algorithm
    // cannot make about a document whose content it never read — all it had was
    // a name. Asserting an absence it has no evidence for is the same defect as
    // calling two PDFs cross-format, one layer along.
    const unrelated = proposePairing({
      sideA: [sketch('a1', 'Project Falcon SPA v6.docx', spa)],
      sideB: [sketch('b1', 'Board minutes March.pdf', [], { sourceKind: 'pdf' })],
      generatedAt: AT,
    });

    const pdf = unrelated.rows.find(r => r.b === 'b1');
    expect(pdf?.reason).toMatch(/content could not be read/i);
    expect(pdf?.reason).not.toMatch(/added, or replacing/i);

    // The readable one keeps the ordinary wording, which for it IS supported.
    expect(unrelated.rows.find(r => r.a === 'a1')?.reason).toMatch(/removed, or replaced/i);
  });

  it('keeps a weak but REAL pair that a filename-only rival would outscore', () => {
    // The case that makes the first phase's exclusion load-bearing rather than
    // decorative. A filename-only score is capped at 0.30 and a corroborated
    // pair can legitimately sit below that — the shipped corpus has one at
    // 0.303 — so on a single scale a perfectly-named PDF would take the
    // document away from a pair that content supports AND that could actually
    // be compared. Preferring an unrunnable pairing to a runnable one is not a
    // trade-off, it is a mistake.
    //
    // The fixture asserts its own score first: if it drifted out of the window
    // this test would pass while proving nothing.
    const overlap = body('shared', 14);
    const weak = proposePairing({
      sideA: [sketch('a1', 'facility agreement 2026.docx', [...overlap, ...body('a', 13)])],
      sideB: [
        sketch('b1', 'revised banking terms.docx', [...overlap, ...body('b', 13)]),
        sketch('b2', 'facility agreement 2026.pdf', [], { sourceKind: 'pdf' }),
      ],
      generatedAt: AT,
    });

    const row = weak.rows.find(r => r.a === 'a1');
    const score = row?.score?.score as number;
    expect(score).toBeGreaterThan(PAIRING_NO_MATCH_SCORE);
    expect(score).toBeLessThanOrEqual(DEFAULT_PAIRING_SETTINGS.noContentCeiling);

    // The runnable pair wins, and the PDF falls out.
    expect(row?.b).toBe('b1');
    expect(row?.runnable).toBe(true);
    expect(weak.rows.find(r => r.b === 'b2')?.band).toBe('unmatched');
  });

  it('lets a content-verified pair the first phase rejected be picked up in the second', () => {
    // The known cost of the two-phase rule, tested so it is a decision rather
    // than a surprise. Phase 1 judged `a1` and `b1` not a pair — their score is
    // under the no-match threshold — so `a1` is a leftover, and the identically
    // named PDF is the only candidate left for it.
    //
    // Tolerable because phase 1 had already said these two are not a pair, the
    // result cannot auto-pair, and both documents keep a row with their rivals.
    const weak = proposePairing({
      sideA: [sketch('a1', 'facility original.docx', body('f', 40))],
      sideB: [
        sketch('b1', 'facility revised.docx', body('g', 40)),
        sketch('b2', 'facility original.pdf', [], { sourceKind: 'pdf' }),
      ],
      generatedAt: AT,
    });

    const row = weak.rows.find(r => r.a === 'a1');
    expect(row?.b).toBe('b2');
    expect(row?.band).toBe('propose-with-rivals');
    expect(weak.rows.find(r => r.b === 'b1')?.band).toBe('unmatched');
  });
});

describe('two PDFs are not "cross-format" (D-3)', () => {
  // The reason text is the only thing the user gets, so it has to be true.
  // `blockFor` used to route ANY non-docx pair through CROSS_FORMAT_REASON,
  // including a PDF against a PDF — which agree perfectly on a format this
  // build does not batch. Telling someone their two PDFs mismatch sends them
  // looking for a mismatch that does not exist, and that is worse than saying
  // nothing.
  const spa = body('spa', 40);

  it('says a Word-against-PDF pair is cross-format', () => {
    const p = proposePairing({
      sideA: [sketch('a1', 'SPA v6.docx', spa)],
      sideB: [sketch('b1', 'SPA v7.pdf', spa, { sourceKind: 'pdf' })],
      generatedAt: AT,
    });
    const row = p.rows.find(r => r.a === 'a1' && r.b === 'b1');
    expect(row?.blocked).toBe('cross-format');
    expect(row?.reason).toMatch(/cross-format/i);
  });

  it('says a PDF-against-PDF pair is a format this build does not batch', () => {
    const p = proposePairing({
      sideA: [sketch('a1', 'SPA v6.pdf', spa, { sourceKind: 'pdf' })],
      sideB: [sketch('b1', 'SPA v7.pdf', spa, { sourceKind: 'pdf' })],
      generatedAt: AT,
    });
    const row = p.rows.find(r => r.a === 'a1' && r.b === 'b1');
    expect(row?.blocked).toBe('unsupported-format');
    expect(row?.reason).toMatch(/only Word documents/i);
    expect(row?.reason).not.toMatch(/cross-format/i);
  });

  it('is still shown, with its pairing and its own disclosure', () => {
    // The refusal changed; M8's requirement did not.
    const p = proposePairing({
      sideA: [sketch('a1', 'SPA v6.pdf', spa, { sourceKind: 'pdf' })],
      sideB: [sketch('b1', 'SPA v7.pdf', spa, { sourceKind: 'pdf' })],
      generatedAt: AT,
    });
    expect(p.totals.unrunnable).toBe(1);
    expect(p.disclosures.join(' ')).toMatch(/only Word documents/i);
    expect(p.disclosures.join(' ')).not.toMatch(/Word-against-PDF/i);
  });
});

describe('the proposal is a proposal — nothing is hidden and nothing has run', () => {
  const sideA = [
    sketch('a-spa', 'SPA v6.docx', body('spa', 60)),
    sketch('a-dl', 'Disclosure Letter v3.docx', body('dl', 30)),
    sketch('a-old', 'Escrow Deed.docx', body('escrow', 20)),
  ];
  const sideB = [
    sketch('b-spa', 'SPA v7.docx', body('spa', 60, 4)),
    sketch('b-dl', 'Disclosure Letter v4.docx', body('dl', 30, 2)),
    sketch('b-new', 'Tax Deed.docx', body('tax', 20)),
  ];

  it('gives every document from both bundles exactly one row', () => {
    const proposal = proposePairing({ sideA, sideB, generatedAt: AT });
    const seen = [
      ...proposal.rows.map(r => r.a).filter(Boolean),
      ...proposal.rows.map(r => r.b).filter(Boolean),
    ];
    expect(new Set(seen).size).toBe(6);
    expect(seen).toHaveLength(6);
  });

  it('shows an unmatched document with the pairings that were rejected', () => {
    const proposal = proposePairing({ sideA, sideB, generatedAt: AT });
    const orphan = proposal.rows.find(r => r.a === 'a-old');

    expect(orphan?.band).toBe('unmatched');
    expect(orphan?.score).toBeNull();
    // "Unmatched" is checkable, not asserted: the rivals are what the solver
    // considered and turned down.
    expect(orphan?.rivals.length).toBeGreaterThan(0);
  });

  it('derives its totals from its rows, so they cannot disagree', () => {
    const proposal = proposePairing({ sideA, sideB, generatedAt: AT });
    expect(proposal.totals).toEqual(tallyPairing(proposal.rows));
    expect(proposal.totals.pairs + proposal.totals.unmatchedA + proposal.totals.unmatchedB)
      .toBe(proposal.rows.length + proposal.totals.pairs - proposal.totals.pairs);
    expect(proposal.totals.autoPair + proposal.totals.needsReview).toBe(proposal.totals.pairs);
  });

  it('injects the date rather than reading the clock', () => {
    const a = proposePairing({ sideA, sideB, generatedAt: AT });
    const b = proposePairing({ sideA, sideB, generatedAt: AT });
    expect(a.generatedAt).toBe(AT);
    expect(a).toEqual(b);
  });

  it('takes no filter parameter — a band cannot remove a row', () => {
    // The structural half of the wall, matching `batch.index-model.test.ts`.
    // Even the weakest possible bundle produces a full grid.
    const weak = proposePairing({
      sideA: [sketch('a-1', 'one.docx', body('one', 10))],
      sideB: [sketch('b-1', 'two.docx', body('two', 10))],
      generatedAt: AT,
    });
    expect(weak.rows).toHaveLength(2);
    expect(weak.rows.every(r => r.band === 'unmatched')).toBe(true);
  });

  it('handles empty bundles on either side', () => {
    expect(proposePairing({ sideA: [], sideB: [], generatedAt: AT }).rows).toEqual([]);
    expect(proposePairing({ sideA, sideB: [], generatedAt: AT }).totals.unmatchedA).toBe(3);
    expect(proposePairing({ sideA: [], sideB, generatedAt: AT }).totals.unmatchedB).toBe(3);
  });
});
