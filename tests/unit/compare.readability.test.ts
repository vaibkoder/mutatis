/**
 * Redline readability conformance — the objective gate on how a redline READS.
 * Spec: spec/features/08-compare.md §5.3–5.5. Principles: docs/compare-readability-principles.md.
 *
 * Two kinds of assertion:
 *
 *  - **Invariants** hold for every case, always. An atomic value is never split
 *    (P5) and a changed paragraph never loses its clause number (P9). These are
 *    not tuning targets; a failure is a defect.
 *  - **A ratchet** on `redundancy` — the share of a struck span that is
 *    word-for-word identical to an inserted span elsewhere in the same change,
 *    i.e. text the reader reads twice to learn nothing happened to it (P1/P2). It may only ever go
 *    DOWN. `marks` is pinned exactly, so any change in rendering shape has to be
 *    acknowledged here rather than slipping through: fixing a case often RAISES
 *    its mark count (three honest deletions beat one confusing block), so this is
 *    an intent gate, not a smaller-is-better one.
 *
 * Ceiling to beat: C07's 0.11 is the floor the corpus can reach — a genuine
 * four-swap sentence where one shared space lands inside a mark. The corpus sits
 * at that floor as of ADR-0030; driving it to 0 means splitting C07's one
 * rewritten phrase into two marks over a space, which reads worse (P3).
 *
 * A third assertion applies to the nine cases where the reviewer wrote down the
 * rendering they wanted: `idealMatch` is pinned, so a case that reaches the
 * wanted rendering can never quietly fall back off it.
 */
import { describe, it, expect } from 'vitest';

import { READABILITY_CASES } from '../fixtures/readability-corpus.js';
import type { ReadabilityCase } from '../fixtures/readability-corpus.js';
import { renderCase, measure } from '../fixtures/readability-metrics.js';

/**
 * Per-case gate. `redundancy` is a ceiling; `marks` and `idealMatch` are pinned.
 *
 * `charRedundancy` is a second ceiling, recorded only where a case exists
 * BECAUSE the token yardstick is blind to its defect (C31–C33). Those three
 * score `redundancy` 0 and `charRedundancy` 0.70–1.00, so without it the fixes
 * in Phase 0.5 would be invisible in every number the suite pins — which is the
 * failure mode ADR-0038 exists to close. It is not recorded for the other cases:
 * a ceiling nobody has read is a number that will be lowered to make a test
 * pass rather than because a redline improved.
 */
const BASELINE: Record<
  string,
  { marks: number; redundancy: number; idealMatch?: boolean; charRedundancy?: number }
> = {
  C01: { marks: 2, redundancy: 0 },
  C02: { marks: 2, redundancy: 0 },
  C03: { marks: 4, redundancy: 0, idealMatch: true },
  C04: { marks: 2, redundancy: 0 },
  C05: { marks: 2, redundancy: 0 },
  C06: { marks: 1, redundancy: 0 },
  C07: { marks: 6, redundancy: 0.11 },
  C08: { marks: 2, redundancy: 0, idealMatch: true },
  C09: { marks: 4, redundancy: 0 },
  C10: { marks: 2, redundancy: 0 },
  C11: { marks: 3, redundancy: 0, idealMatch: true },
  C12: { marks: 2, redundancy: 0, idealMatch: true },
  C13: { marks: 4, redundancy: 0 },
  C14: { marks: 1, redundancy: 0 },
  C15: { marks: 1, redundancy: 0 },
  C16: { marks: 4, redundancy: 0, idealMatch: true },
  C17: { marks: 2, redundancy: 0 },
  C18: { marks: 2, redundancy: 0 },
  C19: { marks: 2, redundancy: 0 },
  C20: { marks: 1, redundancy: 0 },
  // ADR-0032: the merge is now ONE range-spanning replace. Three small marks
  // (`[-.-]`, `[-The-]`, `{+, and the+}`) replace two large ones, and the twelve
  // words that did not change are no longer struck in one paragraph and
  // re-inserted in the next — cross-paragraph redundancy 0.94 → 0.
  // `idealMatch` stays false: the reviewer wrote the wanted rendering as a
  // single paragraph, but a merge that can still be REJECTED must keep two
  // <w:p> until it is accepted (see the ADR).
  C21: { marks: 3, redundancy: 0, idealMatch: false },
  C22: { marks: 12, redundancy: 0 },
  C23: { marks: 2, redundancy: 0, idealMatch: true },
  C24: { marks: 2, redundancy: 0, idealMatch: true },
  C25: { marks: 2, redundancy: 0 },
  C26: { marks: 2, redundancy: 0 },
  C27: { marks: 5, redundancy: 0 },
  C28: { marks: 1, redundancy: 0 },
  C29: { marks: 5, redundancy: 0, idealMatch: true },
  C30: { marks: 0, redundancy: 0 },
  // ── The two word-diff defects . Added BEFORE either fix, so the
  // fix has something to prove. Each one renders a change the reader can see is
  // wrong, and each one scores `redundancy` ZERO.
  //
  // C31 strikes and re-inserts 41 characters to change `shall` to `may`;
  // C32 restates a whole word to say it gained an `s`;
  // C33 strikes 156 characters and re-inserts 162 to change one word.
  //
  // These ceilings drop as Phase 0.5 lands. C31 0.70 → 0.33 and C33 0.81 → 0.25
  // with the possessive-apostrophe guard (ADR-0029's 2026-09-11 correction);
  // reverting that guard puts them back over these ceilings, so this gate is
  // itself an ablation of the fix. C32 is defect A and is untouched by it.
  C31: { marks: 2, redundancy: 0, charRedundancy: 0.33 },
  C32: { marks: 2, redundancy: 0, charRedundancy: 1 },
  C33: { marks: 2, redundancy: 0, charRedundancy: 0.25 },
  // ADR-0039. Before it: 3 marks — the deleted definition, then the edited one
  // inserted whole and struck whole — and `redundancy` 0, because the measure
  // exempts the whole-paragraph shape (ADR-0038). The mark count is the only
  // pinned number here that can see the fix, so it is the gate: 3 → 2.
  C34: { marks: 2, redundancy: 0 },
};

/** Sum of the per-case ceilings. Drops as the readability work lands. */
const TOTAL_REDUNDANCY_CEILING = 0.11;

describe('redline readability corpus', () => {
  for (const c of READABILITY_CASES) {
    it(`${c.id} — ${c.title}`, async () => {
      const m = measure(c, await renderCase(c));
      const base = BASELINE[c.id];
      expect(base, `no baseline recorded for ${c.id}`).toBeDefined();

      // P5 — atomic values are indivisible.
      expect(m.atomicIntact, 'a segment boundary fell inside an atomic value').toBe(true);
      // P9 — never lose the clause number.
      expect(m.numbered, 'a changed paragraph lost its manual clause number').toBe(1);

      // P1/P2 — the ratchet.
      expect(m.redundancy).toBeLessThanOrEqual(base!.redundancy);
      expect(m.marks).toBe(base!.marks);

      // The same ratchet on the character yardstick, where a case was added
      // because the token one reads 0 on a defect the reader can see (ADR-0038).
      if (base!.charRedundancy !== undefined) {
        expect(m.charRedundancy).toBeLessThanOrEqual(base!.charRedundancy);
      }

      // The reviewer's own wanted rendering, where they wrote one down.
      if (base!.idealMatch !== undefined) {
        expect(m.idealMatch, "rendering no longer matches the reviewer's `ideal`").toBe(
          base!.idealMatch,
        );
      }
    });
  }

  it('corpus-wide redundancy stays under the ceiling', async () => {
    let total = 0;
    for (const c of READABILITY_CASES) total += measure(c, await renderCase(c)).redundancy;
    expect(Math.round(total * 100) / 100).toBeLessThanOrEqual(TOTAL_REDUNDANCY_CEILING);
  });
});

/**
 * The yardstick itself. A metric that flags honest renderings is as damaging as
 * one that misses bad ones: it makes every future improvement look like a
 * regression, and it is the number every ADR in this workstream quotes.
 */
describe('the cross-paragraph redundancy measure', () => {
  // ADR-0032 added this measure to catch the "insert the clause here, strike it
  // there" rendering of a paragraph merge, which is clean on every
  // within-paragraph test and reads terribly. It was written first as a token
  // LCS over every insertion in the change, and then narrowed TWICE. Each
  // fixture below pins one of those decisions by failing under the version it
  // replaced — the first draft of this suite pinned neither, because its one
  // case scored 0 under both measures.
  const mk = (
    id: string,
    title: string,
    original: readonly string[],
    revised: readonly string[],
  ): ReadabilityCase => ({
    id,
    title,
    act: 'probe',
    original: original.join('\n'),
    revised: revised.join('\n'),
    expect: '',
    grade: 'good',
    cites: ['P2'],
  });

  it('counts a CONTIGUOUS run, not a scattered subsequence', async () => {
    // Honest: one insertion of new drafting into 1.1, and one whole-paragraph
    // deletion of 1.2. Nothing is read twice. The insertion merely reuses the
    // defined terms `Purchaser`, `Loss` and `Deed` and the adverb `promptly`,
    // which ordinary legal prose does constantly — so a subsequence measure
    // finds them scattered through the struck paragraph and reports a defect.
    // Under the token LCS this scores 0.47 — over the 0.30 line
    // `scripts/readability-report.ts` calls out as the worst cases, and far over
    // every per-case ceiling above. The longest run actually shared is one
    // defined term.
    const honest = mk(
      'FP-scatter',
      'honest insertion reusing defined terms, beside a whole-paragraph deletion',
      [
        '1.1 The Company shall at all times during the Term maintain in full force and effect the insurance policies specified in Schedule 3, and shall on written request produce to any person the certificates evidencing them.',
        '1.2 The Vendor shall promptly indemnify the Purchaser against all Losses arising under the Deed.',
      ],
      [
        '1.1 The Company shall at all times during the Term maintain in full force and effect the insurance policies specified in Schedule 3, and shall promptly, for each Purchaser, upon every Loss beneath any Deed, on written request produce to any person the certificates evidencing them.',
      ],
    );

    const r = await renderCase(honest);
    // The rendering really is the honest shape this fixture claims: 1.2 struck
    // whole, and no part of it re-inserted anywhere.
    expect(r.paras[1]!.segs.every((s) => s.t === 'del' || s.s.trim() === '')).toBe(true);
    expect(measure(honest, r).redundancy).toBeLessThan(0.3);
  });

  it('does not care how far away the re-insertion lands', async () => {
    // The duplicating shape, with the merge's two halves TWO paragraphs apart
    // because the paragraph between them is itself deleted. A reader meets the
    // notice-content clause twice. An intermediate version of this measure
    // looked only at immediately adjacent paragraphs and scored exactly this
    // rendering 0.11 — trading ADR-0032's false positive for a false negative
    // on the very shape it exists to detect.
    const duplicating = mk(
      'TP-distant',
      'merge whose re-insertion lands two paragraphs away',
      [
        '1.1 The Vendor shall give notice of any Claim within 30 days.',
        '1.2 Completion shall take place at the offices of the Vendor\u2019s solicitors.',
        '1.3 The notice shall specify the nature of the Claim in reasonable detail.',
      ],
      [
        '1.1 The Vendor shall give notice of any Claim within 30 days and the notice shall specify the nature of the Claim in reasonable detail.',
      ],
    );

    expect(measure(duplicating, await renderCase(duplicating)).redundancy).toBeGreaterThan(0.6);
  });

  it('does not score a struck span that carries no readable content', async () => {
    // `floor` is `min(struck, inserted)`, so a one-character struck span whose
    // character also appears in the insertion beside it scores 1.00 on nothing.
    // A paragraph join renders the lost sentence break as `[-.-]`; if the
    // inserted joining words END in a full stop, the rendering this workstream
    // is trying to REACH scores the maximum while the bad one scores 0. Built by
    // construction because no corpus case hits it — which is what makes it a
    // trap for the next person to move C21 rather than a visible failure.
    const join = mk(
      'FP-punct',
      'sentence extended past its full stop, the inserted words ending in one',
      ['1.1 The Vendor shall give notice of any Claim within 30 days.'],
      [
        '1.1 The Vendor shall give notice of any Claim within 30 days, provided that the notice is given in writing.',
      ],
    );

    const r = await renderCase(join);
    // The `[-.-]` mark really is present and really is beside an insertion
    // containing a full stop — i.e. the trap is armed in this fixture.
    //
    // The insertion sits BEFORE the struck span (§5.5, ADR-0043), so this looks
    // backwards from the contentless `del`. Deliberately not written
    // order-agnostically: `replaceBlocks` in `readability-metrics.ts` already
    // matches both orders, and if the authored order is ever flipped again this
    // assertion should be the thing that says so.
    const segs = r.paras[0]!.segs;
    const at = segs.findIndex((sg) => sg.t === 'del' && !/[\p{L}\p{N}]/u.test(sg.s));
    expect(at, 'fixture no longer produces a contentless struck span').toBeGreaterThan(0);
    expect(segs[at - 1]?.t).toBe('ins');
    expect(segs[at - 1]!.s).toContain('.');

    expect(measure(join, r).redundancy).toBe(0);
  });

  it('exempts the whole-for-whole swap, which is the other honest rendering', async () => {
    // A wholly-inserted paragraph beside a wholly-struck one is what ADR-0026
    // demotes a dissimilar replace TO, on purpose, and P2 admits it as honest.
    // Counting its insertion would make the demotion score itself as a defect.
    // This is the exemption whose cost ADR-0032 discloses rather than fixes.
    const swap = mk(
      'FP-swap',
      'whole-for-whole paragraph swap',
      ['1.1 The Purchaser shall pay the Consideration.'],
      ['1.1 The Purchaser shall settle every sum owing under the Completion Statement.'],
    );

    expect(measure(swap, await renderCase(swap)).redundancy).toBe(0);
  });
});
