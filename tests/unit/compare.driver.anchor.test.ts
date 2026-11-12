/**
 * Feature 08 — Compare · the "go to change" search anchor.
 *
 * `findAndReveal` selects the **first** match in the document
 * (`host-office/office-host.ts` → `results.items[0].select()`), so an anchor
 * that is not distinctive does not fail loudly — it silently reveals the wrong
 * place. Two shapes make the obvious choice wrong, and both are produced by the
 * paragraph merge/split path (ADR-0032):
 *
 *  - the first mark spans a paragraph break, which the host search cannot cross;
 *  - the first mark is very often bare punctuation — for a paragraph join it is
 *    `"."`, which selects the first full stop in the document.
 *
 * Both were live: the original code took the first mark whole (so the anchor
 * contained U+E000 and could never match), and the first correction took its
 * first line (so the anchor became `"."`). Neither was tested.
 *
 * ## The anchor's SIDE, and why this file's hand-built inputs are not enough
 *
 * `searchAnchorFor` takes the first *changed* line without regard to which side
 * it came from, so ADR-0043's mark order decides that side. The tests below
 * build their own `WordDiff[]`, so they pin the picker's rules and say nothing
 * about which side the pipeline hands it. That is exactly the 08b blind spot —
 * a file describing a shape the pipeline no longer produces — so
 * `'the anchor SIDE follows the authored mark order (ADR-0043)'` at the bottom
 * drives `runComparePipeline` end to end instead.
 */
import { describe, it, expect } from 'vitest';

import { searchAnchorFor } from '@features/compare/driver.js';
import type { WordDiff } from '@features/compare/model.js';
import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import { runComparePipeline } from '@features/compare/pipeline.js';
import { extractTextRecords } from '@features/compare/io/text/extract-text-records.js';

const eq = (text: string): WordDiff => ({ op: 'eq', text });
const del = (text: string): WordDiff => ({ op: 'del', text });
const ins = (text: string): WordDiff => ({ op: 'ins', text });

describe('searchAnchorFor', () => {
  it('takes a changed span when it is distinctive enough to find', () => {
    // NB the input here is written INS-first, as the pipeline now authors it
    // (ADR-0043); the picker takes the first changed line either way, so this
    // pins the distinctiveness rule and not the side. It used to be written
    // del-first — a stream the pipeline no longer produces — which is why the
    // end-to-end side guard below exists.
    expect(
      searchAnchorFor([
        eq('The Vendor '),
        ins('must'),
        del('shall'),
        eq(' indemnify the Purchaser.'),
      ]),
    ).toBe('must');
    // …and the mirror, so the rule is pinned as side-agnostic rather than by
    // accident of which order the fixture happened to use.
    expect(
      searchAnchorFor([
        eq('The Vendor '),
        del('shall'),
        ins('must'),
        eq(' indemnify the Purchaser.'),
      ]),
    ).toBe('shall');
  });

  it('never anchors on the lost full stop of a paragraph join', () => {
    // The reviewer's C21 shape. `"."` would select the first full stop in the
    // document — a wrong jump, which is worse than no jump.
    const anchor = searchAnchorFor([
      eq('within 30 days'),
      del('.\nThe'),
      ins(', and the'),
      eq(' notice shall specify the nature of the Claim.'),
    ]);
    expect(anchor).not.toBe('.');
    expect(anchor).not.toContain('\n');
    expect(anchor.replace(/[^\p{L}\p{N}]/gu, '').length).toBeGreaterThanOrEqual(4);
  });

  it('never anchors on a seam whose only mark is punctuation', () => {
    // The bare seam shape: `del(".\n")` and nothing else changed. There is no
    // distinctive CHANGED text at all, so it must fall back to the op's own
    // unchanged text rather than to `"."`.
    const anchor = searchAnchorFor([
      eq('The Vendor shall give notice of any Claim within 30 days'),
      del('.\n'),
      eq('the notice shall specify the nature of the Claim.'),
    ]);
    expect(anchor).not.toBe('.');
    expect(anchor).not.toContain('\n');
    // The longest unchanged line is the likeliest to be unique in the document.
    expect(anchor).toBe('The Vendor shall give notice of any Claim within 30 days');
  });

  it('never returns an anchor containing the paragraph-break newline', () => {
    for (const diffs of [
      [del('.\nThe Vendor'), eq(' shall pay.')],
      [ins('; and\nthe Purchaser'), eq(' shall pay.')],
      [eq('a'), del('.\n'), eq('b')],
    ]) {
      expect(searchAnchorFor(diffs)).not.toContain('\n');
    }
  });

  it('returns empty rather than aiming at an arbitrary character', () => {
    // A punctuation-only change to a very short paragraph: nothing in the op is
    // distinctive. Doing nothing beats selecting the wrong place.
    expect(searchAnchorFor([eq('a'), del(','), ins(';'), eq('b')])).toBe('');
    expect(searchAnchorFor([])).toBe('');
  });

  it('the anchor SIDE follows the authored mark order (ADR-0043)', async () => {
    // A CONSEQUENCE of the mark order, recorded because ADR-0043 first claimed
    // "nothing else moves" and this moved. Found in review.
    //
    // `searchAnchorFor` takes the first CHANGED line, so flipping which mark
    // comes first flips which document's text the anchor is drawn from:
    //
    //   before ADR-0043   anchor "an exclusive"      → in the ORIGINAL only
    //   after  ADR-0043   anchor "a non-exclusive"   → in the REVISED only
    //
    // **It was never in both, in either direction** — so `findAndReveal` has
    // always missed for whichever single input document is not the anchor's
    // side; ADR-0043 swapped which one, it did not create the miss. The authored
    // REDLINE contains both sides' runs, so navigation over the redline (the
    // default output, §8.1) is unaffected either way.
    //
    // Pinned end to end, not on a hand-built stream: this file's other tests
    // construct their own `WordDiff[]` and cannot see which side the pipeline
    // produces — the 08b lesson.
    //
    // Whether the anchor should instead be SIDE-AWARE (drawn from whichever
    // document is live, or preferring unchanged text present in both) is a
    // design question ADR-0043 does not settle and was not opened earlier.
    const ORIGINAL = 'The Vendor grants the Purchaser an exclusive licence over the Business IP.';
    const REVISED = 'The Vendor grants the Purchaser a non-exclusive licence over the Business IP.';

    const result = await runComparePipeline({
      originalParas: extractTextRecords(ORIGINAL),
      revisedParas: extractTextRecords(REVISED),
      config: getCompareConfig(SG_PROFILE),
      author: 'Mutatis Compare',
      date: '2026-09-17T00:00:00.000Z',
    });

    const op = result.ops.find((o) => o.wordDiffs !== undefined);
    expect(op, 'fixture no longer produces a word-diffed op').toBeDefined();
    // The shape, so the assertion below cannot go vacuous.
    expect(op!.wordDiffs!.map((d) => d.op).join(',')).toBe('eq,ins,del,eq');

    const anchor = searchAnchorFor(op!.wordDiffs!);
    expect(anchor).toBe('a non-exclusive');
    expect(REVISED).toContain(anchor);
    expect(ORIGINAL).not.toContain(anchor);
  });
});
