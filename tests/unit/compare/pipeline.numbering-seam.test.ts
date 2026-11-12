/**
 * ADR-0042 at the SINGLE-PAIR seam — `runPipeline`, the Word-vs-Word path every
 * add-in user takes.
 *
 * `pipeline.ts` builds a `NumberingContext` for each side and hands it to
 * `preprocessSide`. That wiring is two arguments, and until this file existed
 * either one could be deleted with the whole suite still green — while making an
 * UNCHANGED document report four changes, because one side's paragraphs resolve
 * to numbering and the other's resolve to none, so every matched pair differs in
 * identity. That is the shape of the pass-2 `numId` finding and of the two batch
 * seams that shipped unguarded: a change reported that did not happen.
 *
 * The batch seams (`batch/runner.ts`, `batch/driver.ts`, `batch/pairing/sketch.ts`)
 * exist to MIRROR this one, and each has had a guard since the eighth review
 * pass. This is the one they mirror.
 *
 * Two blind spots make these guards easy to get wrong, both learned the hard way
 * (the 08b wiring lesson):
 *
 *  - A `w:pStyle` re-levelling alone proves nothing. The two style ids differ
 *    with or without a numbering context, so the op exists either way. Assert
 *    the numbering half — `numberingIdentity` and the two RENDERED labels — or a
 *    bystander paragraph that only changes when numbering resolves.
 *  - A `property` op authors byte-identically to `equal` (that is the whole
 *    point of ADR-0042), so `authoredBodyXml` can never see one. Assert on
 *    `ops` and `summary`.
 *
 * Fixture: `tests/fixtures/style-numbered-docx.ts` — `document.xml` carries no
 * `w:numPr` at all, which is how a firm precedent is actually built and the only
 * shape in which dropping the context changes the answer.
 */
import { describe, it, expect } from 'vitest';
import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import { runPipeline } from '@features/compare/pipeline.js';
import { extractParts } from '@features/compare/io/parts.js';
import type { DiffOp } from '@features/compare/model.js';
import {
  styleNumberedBundle, styleNumberedNumbering, HEAD, SUB_A, SUB_B, INSERTED,
  type StyledPara,
} from '../../fixtures/style-numbered-docx.js';

const CONFIG = getCompareConfig(SG_PROFILE);
const AUTHOR = 'Mutatis Compare';
const DATE = '2026-09-12T00:00:00Z';

async function compare(
  original: readonly StyledPara[],
  revised: readonly StyledPara[],
  schemes?: { origNumbering: string; revNumbering: string },
) {
  return runPipeline({
    originalParts: extractParts(styleNumberedBundle(original, schemes?.origNumbering)),
    revisedParts: extractParts(styleNumberedBundle(revised, schemes?.revNumbering)),
    config: CONFIG,
    author: AUTHOR,
    date: DATE,
  });
}

const props = (ops: readonly DiffOp[]) => ops.filter((op) => op.type === 'property');

describe('ADR-0042 — the numbering context reaches runPipeline, on both sides', () => {
  it('reports NOTHING when the two sides are the same document', async () => {
    // The ablation guard, and the reason it is stated in this direction: drop
    // `origNumbering` OR `revNumbering` from `runPipeline` and this document —
    // byte-identical on both sides — reports four changes it does not have.
    // ADR-0008 is about never hiding a change; this is its mirror, and a
    // reviewer who cannot trust an unchanged document to be silent cannot trust
    // any row in the list.
    const result = await compare([HEAD, SUB_A, SUB_B], [HEAD, SUB_A, SUB_B]);

    expect(props(result.ops)).toHaveLength(0);
    expect(result.summary.total).toBe(0);
  });

  it('resolves numbering on BOTH sides, so a pure renumbering reads as one', async () => {
    // A sub-clause inserted at the head of the list: `(a)`→`(b)`, `(b)`→`(c)`.
    // The two labels are the assertion — they exist only when BOTH sides
    // resolve, and they are compared from the two rendered labels, never from
    // the key.
    const result = await compare([HEAD, SUB_A, SUB_B], [HEAD, INSERTED, SUB_A, SUB_B]);
    const moved = props(result.ops);

    expect(moved).toHaveLength(2);
    expect(moved.map((op) => [op.propertyChange?.from.numberLabel, op.propertyChange?.to.numberLabel]))
      .toEqual([['(a)', '(b)'], ['(b)', '(c)']]);
    // Identity unchanged — nothing about these paragraphs moved, only their
    // place in the count — so §9.4's answer is `crossRefRenumber`, not
    // `formatting` (which §8.2’s *Hide formatting-only* filter hides by default).
    for (const op of moved) {
      expect(op.propertyChange?.numberingIdentity).toBe(false);
      expect(op.propertyChange?.styleId).toBe(false);
      expect(op.classification).toBe('crossRefRenumber');
    }
  });

  it('reads each side\'s OWN list definitions, not one side\'s twice', async () => {
    // The three tests above prove a numbering context reaches `preprocessSide`.
    // They cannot prove it is the RIGHT one: the fixture gave both sides the
    // same `numbering.xml`, so passing `origNumbering` where `revNumbering`
    // belongs — one token — produced identical labels, zero ops and a green
    // suite, while a list REDEFINITION went unreported. That is ADR-0008
    // suppression on the primary path, and no count assertion can see it.
    //
    // The edit is what a counterparty's house template does: level 0 redefined
    // from `decimal` to `upperRoman`, one token, everything else equal. Under
    // `%1.%2` — the `1.1` style most SG/E&W drafting uses — that ancestor
    // format reaches every descendant's label too, so the redefinition is
    // visible on all three paragraphs and not only on the one that owns it.
    const paras = [HEAD, SUB_A, SUB_B];
    const result = await compare(paras, paras, {
      origNumbering: styleNumberedNumbering({ subText: '%1.%2' }),
      revNumbering: styleNumberedNumbering({ headFmt: 'upperRoman', subText: '%1.%2' }),
    });
    const redefined = props(result.ops);

    expect(redefined).toHaveLength(3);
    expect(redefined.map((op) => [op.propertyChange?.from.numberLabel, op.propertyChange?.to.numberLabel]))
      .toEqual([['1.', 'I.'], ['1.a', 'I.a'], ['1.b', 'I.b']]);
    for (const op of redefined) {
      // The DEFINITION changed, not the paragraph's place in the count — so
      // `formatting`, which §8.2’s filter hides by default but the same
      // section still counts in the total.
      expect(op.propertyChange?.numberingIdentity).toBe(true);
      expect(op.propertyChange?.styleId).toBe(false);
      expect(op.classification).toBe('formatting');
    }
    expect(result.summary.byCategory.formatting).toBe(3);
  });

  it('reports a style-inherited re-levelling AND the bystander it drags', async () => {
    // Outdenting `SUB_A` to level 0 also renumbers `SUB_B` from `(b)` to `(a)`
    // — a paragraph whose own style never moved. The bystander is the half that
    // cannot exist without the numbering context; the re-levelled paragraph
    // alone would pass with the wiring deleted.
    const result = await compare(
      [HEAD, SUB_A, SUB_B],
      [HEAD, ['Clause1', SUB_A[1]], SUB_B],
    );
    const changed = props(result.ops);

    expect(changed).toHaveLength(2);

    const outdented = changed[0]!.propertyChange!;
    expect(outdented.styleId).toBe(true);
    expect(outdented.numberingIdentity).toBe(true);
    expect([outdented.from.numberLabel, outdented.to.numberLabel]).toEqual(['(a)', '2.']);
    expect([outdented.from.ilvl, outdented.to.ilvl]).toEqual([1, 0]);
    expect(changed[0]!.classification).toBe('formatting');

    const bystander = changed[1]!.propertyChange!;
    expect(bystander.styleId).toBe(false);
    expect(bystander.numberingIdentity).toBe(false);
    expect([bystander.from.numberLabel, bystander.to.numberLabel]).toEqual(['(b)', '(a)']);
    expect(changed[1]!.classification).toBe('crossRefRenumber');
  });
});
