/**
 * The companion metrics themselves — the yardstick, not the thing measured.
 * ADR-0038.
 *
 * This workstream has already been burned twice by an instrument that could not
 * see the defect it existed to catch — once when `redundancy` was computed only
 * within a paragraph and scored a paragraph merge 0 (ADR-0032), and again when
 * `readerTokens` inherited the production tokenizer and scored
 * `[-Statement-]{+Statements+}` 0 on a rendering a reviewer
 * inspected. Both times the number said health and the redline said
 * otherwise.
 *
 * So every companion gets a test that fails when THAT measure alone is broken,
 * and several get a hand-built `Rendered` as a negative control — proof the
 * assertion can fail at all, rather than passing because the fixture happens to
 * be clean. Each mutation below was run and watched to fail before this file
 * was committed.
 */
import { describe, it, expect } from 'vitest';

import type { TextReadabilityCase } from '../fixtures/readability-corpus.js';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import type { PizZipBundle } from '@shared/ooxml/bundle.js';
import { styleNumberedBundle, styleNumberedNumbering } from '../fixtures/style-numbered-docx.js';
import { renderCase, measure, companions, formatRendered } from '../fixtures/readability-metrics.js';
import type { Rendered } from '../fixtures/readability-metrics.js';

/** The whole rendering as one reader-visible string. */
const formatted = (r: Rendered): string => formatRendered(r).join(' ');

const mk = (
  id: string,
  original: readonly string[],
  revised: readonly string[],
): TextReadabilityCase => ({
  id,
  title: id,
  act: 'probe',
  original: original.join('\n'),
  revised: revised.join('\n'),
  expect: '',
  grade: 'unreviewed',
  cites: [],
});

describe("the instrument's own numbering seam (docx cases)", () => {
  // `recordsFor`'s docx branch passes each side's styles + numbering into
  // `preprocessSide`, because *"the instrument must run what `runPipeline`
  // runs, or it measures a pipeline that no user has"* — its own doc comment.
  // That argument had no guard: a review deleted the
  // `buildNumberingContext(...)` argument and the WHOLE SUITE stayed green,
  // because every metric the instrument reports is blind
  // to it — a `property` op authors byte-identically to `equal`, so
  // `authoredBodyXml`, `wholeParaChurn`, `alignmentRate`, `redundancy`, `marks`
  // and `charOffendingBlocks` are all unchanged with the wiring gone.
  //
  // So this asserts the one output that CAN see it, `opShapes`, and it does so
  // on a SYNTHETIC pair written to a temp directory rather than on a
  // local-only corpus — a guard that only runs where such files exist is the gap
  // `tests/fixtures/smoke-docx.ts` exists to close, not a guard.
  //
  // The two sides carry DIFFERENT `numbering.xml` (the seam lesson: a
  // fixture whose sides share one definition cannot tell the two contexts
  // apart), and the paragraphs are word-for-word identical, so every op here is
  // `equal` unless the numbering resolved.
  const write = (bundle: PizZipBundle, name: string): string => {
    const dir = mkdtempSync(join(tmpdir(), 'dp-seam-'));
    const path = join(dir, name);
    writeFileSync(path, bundle.toUint8Array());
    return path;
  };

  it('reports a property op, which only exists when the numbering resolved', async () => {
    const paras: Array<[string, string]> = [
      ['Clause1', 'The Vendor shall procure that the Company complies with this Clause.'],
      ['Clause2', 'Settlement shall be made in cleared funds within ten Business Days.'],
      ['Clause2', 'Tax treatment of each such sum shall be determined by the Accountant.'],
    ];
    const original = write(
      styleNumberedBundle(paras, styleNumberedNumbering({ subText: '%1.%2' })), 'original.docx');
    const revised = write(
      styleNumberedBundle(paras, styleNumberedNumbering({ headFmt: 'upperRoman', subText: '%1.%2' })),
      'revised.docx');

    const rendered = await renderCase({
      id: 'SEAM', title: 'numbering seam', act: 'list redefinition',
      expect: 'every paragraph reports a numbering property change',
      grade: 'good', cites: [],
      kind: 'docx', originalPath: original, revisedPath: revised,
    });

    // Identical words on both sides — so without the numbering context every
    // one of these is `equal` and the instrument sees a document that did not
    // change at all.
    expect(rendered.opShapes.map((o) => o.type)).toEqual(['property', 'property', 'property']);
  });
});

describe('charRedundancy — the character yardstick', () => {
  // The three clean reproductions recorded in the plan's § The two word-diff
  // defects. Every one of them is a whole word struck and re-inserted to say
  // that it gained a suffix or a possessive, every one of them is a
  // rendering a reviewer flagged, and every one of them scores `redundancy` ZERO because
  // the production tokenizer emits the two sides as unrelated tokens and the
  // token yardstick is derived from that same tokenizer.
  const repros: Array<[string, string, string]> = [
    [
      'a word gaining a plural suffix',
      '9.1 The Financial Statement shall be prepared in accordance with GAAP.',
      '9.1 The Financial Statements shall be prepared in accordance with GAAP.',
    ],
    [
      'a second word gaining a plural suffix',
      '9.1 The Corporation shall maintain proper books of account.',
      '9.1 The Corporations shall maintain proper books of account.',
    ],
  ];

  for (const [title, original, revised] of repros) {
    it(`scores ${title} non-zero, where the token yardstick scores 0`, async () => {
      const c = mk('repro', [original], [revised]);
      const m = measure(c, await renderCase(c));
      // If this ever reads > 0, the tokenizer has changed and the case no
      // longer demonstrates the blindness — re-derive it rather than relaxing
      // the assertion below.
      expect(m.redundancy, 'the token yardstick is supposed to be blind here').toBe(0);
      expect(m.charRedundancy).toBe(1);
    });
  }

  it('reads 0 once the defect it exposed is fixed — the yardstick, not a ratchet', async () => {
    // `Purchaser` → `Purchaser's agent` was the third of the plan's clean
    // reproductions and scored token 0 / character 1.00 while the engine
    // rendered it as a whole-word swap. ADR-0040 now renders it as the pure
    // insertion it is, so there is no replace block left to score and the
    // character yardstick reads 0. That is the measure doing its job: it made
    // the defect visible, and it goes quiet when the defect goes away.
    const c = mk(
      'possessive-agent',
      ['9.1 The Purchaser shall pay the Consideration on Completion.'],
      ["9.1 The Purchaser's agent shall pay the Consideration on Completion."],
    );
    const r = await renderCase(c);
    expect(formatted(r)).toContain("Purchaser{+'s agent+}");
    expect(measure(c, r).charRedundancy).toBe(0);
  });

  it('scores a genuine rewrite sharing no word 0 on both yardsticks', async () => {
    // The negative control. Without one, "charRedundancy is 1" would be
    // satisfied by a measure that returned 1 unconditionally.
    const c = mk(
      'unrelated',
      ['9.1 The Vendor shall indemnify the Purchaser.'],
      ['9.1 Each Party bears its own costs.'],
    );
    expect(measure(c, await renderCase(c)).charRedundancy).toBe(0);
  });

  it('does not score a shared run that carries no letter or digit', async () => {
    // A genuine rewrite whose two sides happen to share a SPACE and nothing
    // else — `the Vendor` and `a Claim` have no letter in common. `floor` is
    // min(struck, inserted), so without the guard that one space scores 0.14 on
    // a block where the reader re-reads nothing at all. This is the artefact
    // that accounted for 162 of the 390 blocks the corrected sweep first
    // flagged, in its character-yardstick form.
    //
    // Note the struck side DOES carry readable content, so the block is scored
    // rather than skipped — an earlier version of this fixture struck ` . `,
    // which `scoredBlocks` filters out before the guard is ever reached, and
    // the test passed with the guard removed.
    const r: Rendered = {
      paras: [
        {
          label: '',
          text: 'It shall indemnify the Vendor forthwith.',
          segs: [
            { t: 'eq', s: 'It shall indemnify ' },
            { t: 'del', s: 'the Vendor' },
            { t: 'ins', s: 'a Claim' },
            { t: 'eq', s: ' forthwith.' },
          ],
        },
      ],
      text: 'It shall indemnify the Vendor forthwith.',
      ops: [],
      opShapes: [],
      source: { originalText: '', revisedText: '', originalParaCount: 0, revisedParaCount: 0 },
    };
    const k = companions(r);
    expect(k.blocks, 'the block must be SCORED, not skipped, or the guard is untested').toBe(1);
    expect(k.charRedundancyMean).toBe(0);
  });
});

describe('reconstruction — ADR-0008 as a number', () => {
  it('passes when the redline really does rebuild both sides', async () => {
    const c = mk(
      'honest',
      ['1.1 The Vendor shall give notice of any Claim within 30 days.'],
      ['1.1 The Vendor shall give notice of any Claim within 45 days.'],
    );
    const k = companions(await renderCase(c));
    expect(k.reconstruction).toEqual({ original: true, revised: true });
  });

  it('FAILS when a word is missing from the struck side', async () => {
    // The negative control that makes the assertion above mean something. A
    // hand-built rendering whose `del` side has lost one word: this is what a
    // real suppression would look like, and the metric must report it.
    const r: Rendered = {
      paras: [
        {
          label: '',
          text: 'The Vendor shall indemnify the Purchaser.',
          segs: [{ t: 'eq', s: 'The Vendor shall indemnify the Purchaser.' }],
        },
      ],
      text: 'The Vendor shall indemnify the Purchaser.',
      ops: [],
      opShapes: [],
      source: {
        originalText: 'The Vendor shall promptly indemnify the Purchaser.',
        revisedText: 'The Vendor shall indemnify the Purchaser.',
        originalParaCount: 1,
        revisedParaCount: 1,
      },
    };
    expect(companions(r).reconstruction).toEqual({ original: false, revised: true });
  });

  it('is not defeated by §5.2 normalisation or by a restored clause number', async () => {
    // Two things legitimately differ between the authored redline and the
    // records the pipeline matched on, and a reconstruction check that did not
    // account for them would cry wolf on most real documents:
    //
    //  - the redline is authored from the source `rawXml`, so a CURLY
    //    apostrophe survives where `normText` folded it to a straight one;
    //  - ADR-0028 restores the REVISED clause number onto an edited paragraph,
    //    so a paragraph renumbered 1.1 → 1.2 shows only `1.2` and the original
    //    number is legitimately absent.
    //
    // Both are present in this one fixture. Neither is a lost word.
    const c = mk(
      'normalisation-and-renumber',
      ['1.1 Completion shall take place at the offices of the Vendor’s solicitors.'],
      ['1.2 Completion shall take place at the offices of the Purchaser’s solicitors.'],
    );
    expect(companions(await renderCase(c)).reconstruction).toEqual({
      original: true,
      revised: true,
    });
  });
});

describe('wholeParaChurn — the alignment-failure signature', () => {
  // The shape: a paragraph struck whole sitting immediately beside a paragraph
  // inserted whole. `redundancy` exempts it twice over (`isCleanSwap` within a
  // paragraph; "not wholly inserted" across paragraphs), so it scores 0 — the
  // objective function improves as the alignment degrades, and this is the only
  // number that can see it.
  it('counts a whole-paragraph deletion sitting beside a whole-paragraph insertion', async () => {
    const c = mk(
      'churn',
      [
        '1.1 The Vendor shall give notice of any Claim within 30 days.',
        '1.2 Completion shall take place at the offices of the solicitors.',
      ],
      [
        '1.1 The Vendor shall give notice of any Claim within 30 days.',
        '1.2 Each Party shall bear its own costs of and incidental to this Deed.',
      ],
    );
    const r = await renderCase(c);
    // The rendering really is the churn shape: the paragraph was NOT paired.
    expect(r.opShapes.map((o) => o.type)).toEqual(['equal', 'delete', 'insert']);
    const k = companions(r);
    expect(k.wholeParaChurn).toBe(1);
    // And the existing objective function cannot see it. This is the whole
    // reason the companion exists; if it ever stops being true, the blindness
    // has been fixed elsewhere and this comment is stale.
    expect(measure(c, r).redundancy).toBe(0);
  });

  it('does not count a deletion with no insertion beside it', async () => {
    // Same whole-paragraph deletion, nothing inserted. An honest deletion is
    // not churn, and a measure that merely counted deletions would score this 1.
    const c = mk(
      'no-churn',
      [
        '1.1 The Vendor shall give notice of any Claim within 30 days.',
        '1.2 Completion shall take place at the offices of the solicitors.',
      ],
      ['1.1 The Vendor shall give notice of any Claim within 30 days.'],
    );
    const r = await renderCase(c);
    expect(r.opShapes.map((o) => o.type)).toEqual(['equal', 'delete']);
    expect(companions(r).wholeParaChurn).toBe(0);
  });

  it('reads the ops, not the painted paragraphs', async () => {
    // Deliberate: a struck paragraph does not always render as an all-`del`
    // paragraph, because ADR-0035 revises the paragraph MARK as well. Counting
    // the rendering would measure the renderer; counting the ops measures the
    // pairing, which is what the number is about. A hand-built `Rendered` with
    // NO paragraphs at all still scores the churn in its ops.
    const r: Rendered = {
      paras: [],
      text: '',
      ops: [],
      opShapes: [
        { type: 'delete', origParas: 1, revParas: 0 },
        { type: 'insert', origParas: 0, revParas: 1 },
      ],
      source: { originalText: '', revisedText: '', originalParaCount: 1, revisedParaCount: 1 },
    };
    expect(companions(r).wholeParaChurn).toBe(1);
  });
});

describe('alignmentRate — how much of the original found a counterpart', () => {
  it('counts equal, replace and move as matched, and delete as not', async () => {
    const c = mk(
      'align',
      [
        '1.1 The Vendor shall give notice of any Claim within 30 days.',
        '1.2 Completion shall take place at the offices of the solicitors.',
        '1.3 The Purchaser shall pay the Consideration in cash.',
        '1.4 Each Party shall bear its own costs.',
      ],
      [
        '1.1 The Vendor shall give notice of any Claim within 45 days.',
        '1.3 The Purchaser shall pay the Consideration in cash.',
        '1.4 Each Party shall bear its own costs.',
      ],
    );
    const r = await renderCase(c);
    expect(r.opShapes.map((o) => o.type)).toEqual(['replace', 'delete', 'equal', 'equal']);
    // Three of four original paragraphs found a counterpart; 1.2 did not.
    expect(companions(r).alignmentRate).toBe(0.75);
  });

  it('counts a `property` op as MATCHED — it is an equal op wearing a new name', () => {
    // ADR-0042 converts an `equal` op in place, keeping both ranges and
    // re-pairing nothing. When the type was added this function did not know it
    // and counted it as unmatched: `alignmentRate` on a representative pair fell
    // 1 → 0.94 for a change that provably altered no pairing (the authored XML
    // was byte-identical, 16 of 16 paragraphs still matched). That is the
    // instrument misreading a new op type as a regression — the failure mode
    // this file exists to keep out. The real-document test that also covers it
    // is `describe.skipIf`-gated on fixtures that are never committed, so it
    // does not run in CI; this one does.
    const r: Rendered = {
      paras: [],
      text: '',
      ops: [],
      opShapes: [
        { type: 'property', origParas: 1, revParas: 1 },
        { type: 'property', origParas: 1, revParas: 1 },
        { type: 'equal', origParas: 1, revParas: 1 },
        { type: 'delete', origParas: 1, revParas: 0 },
      ],
      source: { originalText: '', revisedText: '', originalParaCount: 4, revisedParaCount: 3 },
    };
    // Three of four matched. Counting `property` as unmatched gives 0.25.
    expect(companions(r).alignmentRate).toBe(0.75);
  });

  it('is 0 when every original paragraph is deleted outright', () => {
    const r: Rendered = {
      paras: [],
      text: '',
      ops: [],
      opShapes: [
        { type: 'delete', origParas: 1, revParas: 0 },
        { type: 'delete', origParas: 1, revParas: 0 },
        { type: 'insert', origParas: 0, revParas: 1 },
      ],
      source: { originalText: '', revisedText: '', originalParaCount: 2, revisedParaCount: 1 },
    };
    expect(companions(r).alignmentRate).toBe(0);
  });
});

describe('the redundancy distribution', () => {
  it('reports a mean and a p95, not the max twice over', async () => {
    // One bad block pins a max at ~1.0 on a 26-paragraph document and the max
    // then stops discriminating. The distribution has to be able to disagree
    // with the max, so this fixture has two blocks scoring differently.
    const c = mk(
      'spread',
      [
        '1.1 The Financial Statement shall be prepared in accordance with GAAP.',
        '1.2 Completion shall occur on 1 January 2026.',
      ],
      [
        '1.1 The Financial Statements shall be prepared in accordance with GAAP.',
        '1.2 Completion shall occur on 1 February 2026.',
      ],
    );
    const r = await renderCase(c);
    const k = companions(r);
    expect(k.blocks).toBe(2);
    // `Statement`→`Statements` is 1.00; `1 January 2026`→`1 February 2026`
    // shares only `uary 2026`. A mean implemented as a max would read 1.
    expect(k.charRedundancyP95).toBe(1);
    expect(k.charRedundancyMean).toBeLessThan(1);
    expect(k.charRedundancyMean).toBeGreaterThan(0.5);
  });

  it('counts offenders on BOTH yardsticks, because the token one can read 0 on a real defect', async () => {
    const c = mk(
      'offenders',
      ['9.1 The Financial Statement shall be prepared in accordance with GAAP.'],
      ['9.1 The Financial Statements shall be prepared in accordance with GAAP.'],
    );
    const k = companions(await renderCase(c));
    // The same single block: invisible to the token floor, caught by the
    // character one. Reporting only the first would publish a clean sheet.
    expect(k.offendingBlocks).toBe(0);
    expect(k.charOffendingBlocks).toBe(1);
    expect(k.charOffendingBlocksPerKWord).toBeGreaterThan(0);
  });
});
