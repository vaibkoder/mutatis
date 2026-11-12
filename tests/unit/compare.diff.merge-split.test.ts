/**
 * Feature 08 — Compare §5.1–5.3 · paragraph merge / split re-pairing (ADR-0032).
 *
 * Level-1 matches paragraphs one-to-one, so a drafter joining two paragraphs
 * leaves the pipeline showing the same words twice — once struck in the orphaned
 * paragraph, once inserted into its neighbour — and the reader has to diff the
 * redline to see that nothing happened to them (P2). These tests pin the
 * re-pairing, every guard that keeps it from gluing unrelated paragraphs
 * together, and the two invariants it may never break: nothing leaves the
 * redline (ADR-0008), and an ordinary one-to-one `replace` is completely
 * unaffected.
 */
import { describe, it, expect } from 'vitest';
import { XMLValidator } from 'fast-xml-parser';

import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import type { CompareConfig } from '@shared/config/index.js';
import type { ParaRecord, DiffOp } from '@features/compare/model.js';
import {
  structuralDiff,
  detectMergesAndSplits,
  rangeDiffText,
  wordDiff,
  PARA_SEAM,
} from '@features/compare/diff/index.js';
// White-box: the guard test asserts against the constant itself, so it tracks
// the code rather than a number copied into the test.
import { MERGE_MIN_TOKENS } from '@features/compare/diff/merge-split.js';
import { markParagraphMark, wrapParaRuns } from '@features/compare/output/revision-markup.js';
import { splitParagraph } from '@features/compare/output/author.js';
import { fnv1a } from '@features/compare/preprocess/index.js';
import { runComparePipeline } from '@features/compare/pipeline.js';
import { extractTextRecords } from '@features/compare/io/text/extract-text-records.js';

const CONFIG: CompareConfig = getCompareConfig(SG_PROFILE);
const AUTHOR = 'Test';
const DATE = '2026-06-17T00:00:00.000Z';

function para(index: number, text: string, extra: Partial<ParaRecord> = {}): ParaRecord {
  return {
    index,
    partId: 'document',
    rawXml: `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`,
    normText: text,
    normTextAggressive: text.toLowerCase(),
    numberingPrefix: '',
    hash: fnv1a(text),
    styleId: null,
    isTableCell: false,
    isSignatureBlock: false,
    isHeaderFooter: false,
    isFootnote: false,
    ...extra,
  };
}
const side = (...texts: string[]): ParaRecord[] => texts.map((t, i) => para(i, t));

/** Stage 2 → 2.25, exactly as the pipeline runs them. */
function repair(original: ParaRecord[], revised: ParaRecord[], cfg = CONFIG): DiffOp[] {
  return detectMergesAndSplits(
    structuralDiff(original, revised, AUTHOR, DATE),
    original,
    revised,
    cfg,
  );
}

/** Every record index each side's ops claim, in order — the never-suppress check. */
function claimed(ops: readonly DiffOp[], which: 'origRange' | 'revRange'): number[] {
  const out: number[] = [];
  for (const op of ops) {
    const range = op[which];
    if (range === null) continue;
    for (let i = range.start; i <= range.end; i++) out.push(i);
  }
  return out;
}

// The reviewer's C21: a full stop becomes "; and" and the following paragraph is
// pulled up into the first. Written long enough to clear MERGE_MIN_TOKENS.
const P1 = 'The Vendor shall give notice of any Claim within 30 days.';
const P1_JOINED = 'The Vendor shall give notice of any Claim within 30 days; and';
const P2 = 'the notice shall specify the nature of the Claim in reasonable detail.';

describe('rangeDiffText (§5.3, ADR-0032)', () => {
  it('is byte-identical to normText for a single-record range', () => {
    // The invariant every existing one-to-one `replace` depends on: joining a
    // range of one must not touch the text, or every word diff in the product
    // changes the day this pass lands.
    const recs = side(P1, P2, 'A third paragraph with several words in it.');
    for (const [i, rec] of recs.entries()) {
      expect(rangeDiffText(recs, { start: i, end: i })).toBe(rec.normText);
    }
  });

  it('joins a multi-record range with one seam per paragraph mark', () => {
    const recs = side(P1, P2);
    expect(rangeDiffText(recs, { start: 0, end: 1 })).toBe(`${P1}${PARA_SEAM}${P2}`);
    expect(rangeDiffText(recs, null)).toBe('');
  });

  it('carries the paragraph mark into the word diff as its own deletion', () => {
    // The seam is neither whitespace nor a word character, so the tokenizer
    // emits it alone and the diff can strike it as a unit — which is what the
    // author turns into a rejectable revision on the paragraph mark.
    const diffs = wordDiff(`${P1}${PARA_SEAM}${P2}`, `${P1_JOINED} ${P2}`, CONFIG);
    const struck = diffs.filter((d) => d.op === 'del').map((d) => d.text);
    expect(struck.some((t) => t.includes(PARA_SEAM))).toBe(true);
    // …and nothing leaves either side.
    const rebuild = (s: 'ins' | 'del'): string =>
      diffs.filter((d) => d.op === 'eq' || d.op === s).map((d) => d.text).join('');
    expect(rebuild('del')).toBe(`${P1}${PARA_SEAM}${P2}`);
    expect(rebuild('ins')).toBe(`${P1_JOINED} ${P2}`);
  });
});

describe('detectMergesAndSplits — merge (§5.1, ADR-0032)', () => {
  it('re-pairs replace + delete into one range-spanning replace', () => {
    const ops = repair(side(P1, P2), side(`${P1_JOINED} ${P2}`));
    expect(ops).toHaveLength(1);
    expect(ops[0]!.type).toBe('replace');
    expect(ops[0]!.origRange).toEqual({ start: 0, end: 1 });
    expect(ops[0]!.revRange).toEqual({ start: 0, end: 0 });
  });

  it('re-pairs a merge joined on a semicolon, with no punctuation change', () => {
    // A second content shape, NOT a second adjacency: `emitGap` emits a
    // segment's `replace` before its orphans, so this is `replace + delete`
    // like the case above. (It was once named for a "mirror adjacency" the
    // pipeline cannot produce — see the guards block.)
    const first = 'Any Claim shall be notified to the Vendor in writing;';
    const second = 'and the notice shall specify the nature of that Claim.';
    const ops = repair(side(first, second), side(`${first} ${second}`));
    expect(ops).toHaveLength(1);
    expect(ops[0]!.type).toBe('replace');
    expect(ops[0]!.origRange).toEqual({ start: 0, end: 1 });
  });

  it('preserves every record on both sides (ADR-0008)', () => {
    const ops = repair(side(P1, P2), side(`${P1_JOINED} ${P2}`));
    expect(claimed(ops, 'origRange')).toEqual([0, 1]);
    expect(claimed(ops, 'revRange')).toEqual([0]);
  });
});

describe('detectMergesAndSplits — split (§5.1, ADR-0032)', () => {
  it('re-pairs replace + insert into one range-spanning replace', () => {
    const ops = repair(side(`${P1_JOINED} ${P2}`), side(P1, P2));
    expect(ops).toHaveLength(1);
    expect(ops[0]!.type).toBe('replace');
    expect(ops[0]!.origRange).toEqual({ start: 0, end: 0 });
    expect(ops[0]!.revRange).toEqual({ start: 0, end: 1 });
  });

  it('re-pairs a split at a semicolon, with no punctuation change', () => {
    // As above: `replace + insert`, a content variant rather than a new ordering.
    const first = 'Any Claim shall be notified to the Vendor in writing;';
    const second = 'and the notice shall specify the nature of that Claim.';
    const ops = repair(side(`${first} ${second}`), side(first, second));
    expect(ops).toHaveLength(1);
    expect(ops[0]!.revRange).toEqual({ start: 0, end: 1 });
    expect(claimed(ops, 'revRange')).toEqual([0, 1]);
    expect(claimed(ops, 'origRange')).toEqual([0]);
  });
});

describe('detectMergesAndSplits — guards (ADR-0032)', () => {
  it('declines a fragment shorter than MERGE_MIN_TOKENS', () => {
    // A limb stub matches too easily by chance, and the reading gain on four
    // words is nil. Proven against the constant so the test tracks the code.
    const stub = '; and';
    expect(stub.split(/\s+/).filter(Boolean).length).toBeLessThan(MERGE_MIN_TOKENS);
    const ops = repair(side(P1, stub), side(`${P1} ${stub}`));
    expect(ops.length).toBeGreaterThan(1);
  });

  it('declines when the absorbed words are present but out of order', () => {
    // Containment is POSITIONAL. A paragraph that merely shares vocabulary with
    // its neighbour has not been absorbed into it, and gluing the two together
    // would misreport the edit.
    const scrambled = 'detail reasonable in Claim the of nature the specify shall notice the';
    const ops = repair(side(P1, P2), side(`${P1_JOINED} ${scrambled}`));
    expect(ops.length).toBeGreaterThan(1);
  });

  it('declines when what is left of the host no longer resembles its pair', () => {
    // The host ends on the absorbed paragraph, but its opening half is a
    // wholesale rewrite — that is two unrelated edits, not a merge. Reuses the
    // ADR-0026 similarity threshold so the two gates cannot disagree.
    const unrelated = 'Nothing in this Deed excludes liability for fraudulent misrepresentation.';
    const ops = repair(side(P1, P2), side(`${unrelated} ${P2}`));
    expect(ops.length).toBeGreaterThan(1);
  });

  it('does not re-pair a whole-paragraph delete beside an unchanged neighbour', () => {
    // Scope note: this pins the BEHAVIOUR (a plain delete stays a plain delete)
    // but does not isolate any one guard — Level-1 hash-matches the surviving
    // paragraph, so it emits `delete + equal` and no `replace` exists for the
    // pass to consider. The guard itself is isolated by the test below.
    const ops = repair(side(P1, P2), side(P2));
    expect(ops.every((op) => op.origRange === null || op.origRange.start === op.origRange.end)).toBe(
      true,
    );
  });

  it('declines a host that is entirely the absorbed paragraph', () => {
    // This isolates `host.length <= part.length`, and it took an outside review
    // to find the fixture. I had recorded the guard as "not separately
    // observable", reasoning that firing it forces an empty `remainder` so the
    // Dice check declines anyway. That is wrong: `dice()` returns 1 when BOTH
    // sides are empty, so when the head paragraph carries no content tokens the
    // Dice check PASSES and this line is the only brake left.
    //
    // A bare "." paragraph above the clause is the shape — it survives §5.2
    // normalisation and has no content tokens. Without the guard the pass
    // re-pairs these into a merge and authors a struck paragraph mark, i.e. it
    // reports a paragraph join that did not happen.
    const original = side('.', 'The Vendor shall indemnify the Purchaser against all Losses.');
    const revised = side('The Vendor shall indemnify the Purchaser against all Losses;');
    const level1 = structuralDiff(original, revised, AUTHOR, DATE);
    expect(level1.map((o) => o.type)).toEqual(['replace', 'delete']);

    const out = detectMergesAndSplits(level1, original, revised, CONFIG);
    expect(out).toHaveLength(2);
    expect(out.every((o) => o.origRange === null || o.origRange.start === o.origRange.end)).toBe(
      true,
    );
  });

  it('declines a table unit, which is one alignment unit (ADR-0027)', () => {
    const original = [para(0, P1), para(1, P2, { table: { rows: 2, cols: 2, cells: [] } as never })];
    const ops = repair(original, side(`${P1_JOINED} ${P2}`));
    expect(ops.length).toBeGreaterThan(1);
  });

  it('declines a replace and an orphan that are adjacent in the op list but not in the document', () => {
    // WHITE-BOX: `emitGap` emits every `replace` in a segment before its
    // orphans and separates segments with an `equal` anchor, so the pipeline
    // cannot put a replace next to a NON-adjacent orphan. Built by hand so the
    // behaviour is pinned if Level-1 emission ever changes.
    //
    // Honest scope note: this asserts the OUTCOME, and the outcome is defended
    // twice over, so ablating `b.origRange.start === a.origRange.end + 1` alone
    // does not fail it. Widening across a skipped record necessarily changes the
    // set of records the op list claims, which the never-suppress guard
    // (ADR-0008) detects independently — so the adjacency precondition, like
    // `host.length <= part.length`, states intent rather than being separately
    // observable. Defence in depth is the point; the assertion below covers
    // both layers.
    const original = side(P1, 'An unrelated intervening clause applies here.', P2);
    const revised = side(`${P1_JOINED} ${P2}`);
    const all = structuralDiff(original, revised, AUTHOR, DATE);
    const replaceOp = all.find((o) => o.type === 'replace')!;
    const deleteP2 = all.find((o) => o.type === 'delete' && o.origRange?.start === 2)!;
    expect(replaceOp.origRange).toEqual({ start: 0, end: 0 });

    // Record 2 IS absorbed at the host's tail, so `absorbs` itself would say yes.
    const ops = [replaceOp, deleteP2];
    const out = detectMergesAndSplits(ops, original, revised, CONFIG);
    expect(out).toEqual(ops);
    // The intervening record 1 must never be swallowed into the merge.
    expect(claimed(out, 'origRange')).toEqual([0, 2]);
  });

  it('declines an orphan that spans more than one paragraph', () => {
    // Also white-box, and load-bearing: without this precondition the pass
    // widens the replace over BOTH orphaned records while `absorbs` has only
    // inspected the first, and the never-suppress guard cannot see it — the set
    // of records claimed is identical either way. The result would be a
    // three-paragraph op silently diffed against one.
    const original = side(P1, P2, 'A further paragraph of some length here.');
    const revised = side(`${P1_JOINED} ${P2}`);
    const all = structuralDiff(original, revised, AUTHOR, DATE);
    const replaceOp = all.find((o) => o.type === 'replace')!;
    const firstDelete = all.find((o) => o.type === 'delete')!;
    const spanning = { ...firstDelete, origRange: { start: 1, end: 2 } };

    const ops = [replaceOp, spanning];
    expect(detectMergesAndSplits(ops, original, revised, CONFIG)).toEqual(ops);
  });

  it('declines a record whose own text contains the seam sentinel (ADR-0008)', () => {
    // `normaliseText` does not strip Private Use Area code points, and PDF
    // extraction with a broken ToUnicode map — or verbatim pasted text — yields
    // them. Re-pairing on such a record makes the sentinel ambiguous: the author
    // cuts at the document's own U+E000, emitting an extra <w:p> with no
    // revision on it and dropping the character with no <w:del>.
    const dirty = `the notice shall specify${PARA_SEAM} the nature of the Claim in reasonable detail.`;
    const ops = repair(side(P1, dirty), side(`${P1_JOINED} ${dirty}`));
    expect(ops.length).toBeGreaterThan(1);
    // and nothing is lost either way
    expect(claimed(ops, 'origRange')).toEqual([0, 1]);
  });

  it('leaves an ordinary one-to-one replace exactly as Level-1 produced it', () => {
    const before = structuralDiff(
      side('The Vendor shall pay the Purchaser on demand.', P2),
      side('The Vendor must pay the Purchaser on demand.', P2),
      AUTHOR,
      DATE,
    );
    const after = detectMergesAndSplits(
      before,
      side('The Vendor shall pay the Purchaser on demand.', P2),
      side('The Vendor must pay the Purchaser on demand.', P2),
      CONFIG,
    );
    expect(after).toEqual(before);
  });

  it('returns its input unchanged when there is nothing to pair', () => {
    const single = structuralDiff(side(P1), side(P1_JOINED), AUTHOR, DATE);
    expect(detectMergesAndSplits(single, side(P1), side(P1_JOINED), CONFIG)).toBe(single);
  });
});

describe('markParagraphMark (§5.5, ADR-0032)', () => {
  it('appends the mark rPr to an existing w:pPr', () => {
    const out = markParagraphMark(
      '<w:p><w:pPr><w:pStyle w:val="Body"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
      'del',
      7,
      AUTHOR,
      DATE,
    );
    // w:rPr sits near the end of CT_PPr's sequence, so it goes last inside w:pPr.
    expect(out).toContain('<w:pStyle w:val="Body"/><w:rPr><w:del ');
    expect(out).toContain('</w:rPr></w:pPr><w:r>');
    // Exactly one mark, and it is on the PARAGRAPH mark — not a run-level
    // deletion, which is what would wrongly strike the paragraph's text.
    expect(out.match(/<w:del /g)).toHaveLength(1);
    expect(out.indexOf('<w:rPr>')).toBeGreaterThan(out.indexOf('<w:pStyle'));
  });

  it('expands a self-closed w:pPr', () => {
    const out = markParagraphMark('<w:p><w:pPr/><w:r><w:t>x</w:t></w:r></w:p>', 'ins', 7, AUTHOR, DATE);
    expect(out).toContain('<w:pPr><w:rPr><w:ins ');
    expect(out).not.toContain('<w:pPr/>');
  });

  it('adds properties to a paragraph that has none, and keeps attributes on w:p', () => {
    const out = markParagraphMark('<w:p w:rsidR="00AB"><w:r><w:t>x</w:t></w:r></w:p>', 'del', 7, AUTHOR, DATE);
    expect(out.startsWith('<w:p w:rsidR="00AB"><w:pPr><w:rPr><w:del ')).toBe(true);
    expect(out).toContain('</w:rPr></w:pPr><w:r>');
  });

  it('leaves anything that is not a paragraph untouched', () => {
    const notAPara = '<w:tbl><w:tr/></w:tbl>';
    expect(markParagraphMark(notAPara, 'del', 7, AUTHOR, DATE)).toBe(notAPara);
  });

  // ── CT_PPr conformance. `CT_PPr` permits ONE w:rPr and orders it
  // `rPr → sectPr → pPrChange`; these three shapes are all written by Word
  // routinely and each one produced invalid document.xml before this suite.
  it('merges into an rPr the paragraph mark already has, instead of adding a second', () => {
    // What Word writes whenever the paragraph mark carries direct formatting.
    const out = markParagraphMark(
      '<w:p><w:pPr><w:pStyle w:val="Body"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
      'del',
      7,
      AUTHOR,
      DATE,
    );
    expect(out.match(/<w:rPr[\s/>]/g)).toHaveLength(1);
    expect(out.match(/<w:del /g)).toHaveLength(1);
    // Inside CT_ParaRPr, w:ins/w:del come first.
    expect(out).toContain('<w:rPr><w:del ');
    expect(out).toContain('/><w:b/></w:rPr>');
  });

  it('merges into a quarantined empty rPr (§9.2 strips a pre-existing w:del)', () => {
    const out = markParagraphMark(
      '<w:p><w:pPr><w:rPr/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
      'del',
      7,
      AUTHOR,
      DATE,
    );
    expect(out.match(/<w:rPr[\s/>]/g)).toHaveLength(1);
    expect(out).toContain('<w:pPr><w:rPr><w:del ');
  });

  it('inserts the mark BEFORE a sectPr, not after it', () => {
    // The last paragraph of any section carries sectPr, which must follow rPr.
    const out = markParagraphMark(
      '<w:p><w:pPr><w:pStyle w:val="Body"/><w:sectPr><w:pgSz w:w="11906"/></w:sectPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
      'del',
      7,
      AUTHOR,
      DATE,
    );
    expect(out.indexOf('<w:del ')).toBeLessThan(out.indexOf('<w:sectPr'));
    expect(out.match(/<w:rPr[\s/>]/g)).toHaveLength(1);
  });

  it('expands a self-closed rPr that carries attributes, rather than reading it as an open tag', () => {
    // `<w:rPr(?:\s[^>]*)?>` matches `<w:rPr w:x="1"/>` — `[^>]*` swallows the
    // slash — so the revision was appended AFTER the self-closing tag, making
    // `w:del` a direct child of CT_PPr and leaving the paragraph mark unmarked.
    // The merge is then unrejectable: the failure this function exists to
    // prevent, reached through the one rPr shape the code did not expect.
    const out = markParagraphMark(
      '<w:p><w:pPr><w:rPr w:x="1"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
      'del',
      7,
      AUTHOR,
      DATE,
    );
    expect(XMLValidator.validate(out), `not well-formed: ${out}`).toBe(true);
    expect(out.match(/<w:rPr[\s/>]/g)).toHaveLength(1);
    // The mark is INSIDE the rPr, not a sibling of it.
    expect(out).toContain('<w:rPr w:x="1"><w:del ');
    expect(out).not.toContain('/><w:del ');
  });

  it('does not inject the mark into the nested w:pPr inside w:pPrChange', () => {
    // `w:pPrChange` carries a complete <w:pPr> of its own — a CT_PPrBase, which
    // has no rPr at all. A lazy `</w:pPr>` match lands the revision inside it,
    // and then the paragraph mark is NOT marked: the merge becomes unrejectable,
    // the exact failure markParagraphMark exists to prevent.
    const out = markParagraphMark(
      '<w:p><w:pPr><w:pStyle w:val="Body"/><w:pPrChange w:id="9" w:author="A" w:date="D">' +
        '<w:pPr><w:pStyle w:val="Old"/></w:pPr></w:pPrChange></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
      'del',
      7,
      AUTHOR,
      DATE,
    );
    expect(out.match(/<w:del /g)).toHaveLength(1);
    expect(out.indexOf('<w:del ')).toBeLessThan(out.indexOf('<w:pPrChange'));
    // the pPrChange's own properties survive untouched
    expect(out).toContain('<w:pPrChange w:id="9" w:author="A" w:date="D"><w:pPr><w:pStyle w:val="Old"/></w:pPr></w:pPrChange>');
  });
});

describe('splitting off a leading w:pPr (ADR-0032)', () => {
  // `w:pPr` NESTS: `w:pPrChange` carries a complete `<w:pPr>` of its own. All
  // three call sites used a lazy `</w:pPr>` match, which stops at the INNER
  // close, so the split lands mid-element. `markParagraphMark`'s share of that
  // bug is covered above; these are the other two, and neither was covered when
  // `matchLeadingPPr` was introduced — the whole suite stayed green with the
  // lazy regex put back.
  //
  // A paragraph whose formatting was changed under track-changes carries
  // `w:pPrChange`, so this is not an exotic shape: it is any paragraph a
  // reviewer restyled before the comparison.
  const RESTYLED =
    '<w:p><w:pPr><w:pStyle w:val="Body"/>' +
    '<w:pPrChange w:id="9" w:author="A" w:date="D"><w:pPr><w:pStyle w:val="Old"/></w:pPr></w:pPrChange>' +
    '</w:pPr><w:r><w:t>x</w:t></w:r></w:p>';

  const wellFormed = (xml: string): boolean => XMLValidator.validate(xml) === true;

  it('wrapParaRuns keeps the whole pPr outside the revision element', () => {
    const out = wrapParaRuns(RESTYLED, (runs) => `<w:del w:id="1">${runs}</w:del>`);

    // Under the lazy match the wrapper opened between `</w:pPr>` and
    // `</w:pPrChange>`, giving `<w:del>` a close tag it did not open.
    expect(wellFormed(out), `not well-formed: ${out}`).toBe(true);
    expect(out).toContain('</w:pPrChange></w:pPr><w:del w:id="1">');
    expect(out).toContain('<w:del w:id="1"><w:r><w:t>x</w:t></w:r></w:del>');
    // Nothing from the paragraph properties was swept into the revision.
    expect(out).not.toContain('<w:del w:id="1"><w:pPr');
    expect(out).not.toContain('<w:del w:id="1"></w:pPrChange>');
  });

  it('splitParagraph returns the pPr whole, leaving only runs in `inner`', () => {
    const { open, pPr, inner, close } = splitParagraph(RESTYLED);

    expect(open).toBe('<w:p>');
    expect(close).toBe('</w:p>');
    expect(pPr).toBe(
      '<w:pPr><w:pStyle w:val="Body"/>' +
        '<w:pPrChange w:id="9" w:author="A" w:date="D"><w:pPr><w:pStyle w:val="Old"/></w:pPr></w:pPrChange>' +
        '</w:pPr>',
    );
    expect(inner).toBe('<w:r><w:t>x</w:t></w:r>');
    // The split is lossless: the four parts reassemble the input exactly.
    expect(open + pPr + inner + close).toBe(RESTYLED);
    // And each part is independently well-formed, which is what callers assume
    // when they re-emit `open + pPr + <something> + close`.
    expect(wellFormed(`<w:p>${pPr}</w:p>`)).toBe(true);
    expect(wellFormed(`<w:p>${inner}</w:p>`)).toBe(true);
  });

  it('still handles the shapes that have no nesting to trip over', () => {
    expect(splitParagraph('<w:p><w:pPr/><w:r><w:t>x</w:t></w:r></w:p>').pPr).toBe('<w:pPr/>');
    expect(splitParagraph('<w:p><w:r><w:t>x</w:t></w:r></w:p>').pPr).toBe('');
    expect(splitParagraph('<w:p><w:pPr><w:pStyle w:val="B"/></w:pPr></w:p>').inner).toBe('');
  });

  it('declines to split an unbalanced pPr rather than cutting at a guess', () => {
    // Truncated input must not yield a `pPr` that is missing its close tag —
    // the caller re-emits it verbatim and would author unopenable XML.
    const truncated = '<w:p><w:pPr><w:pStyle w:val="B"/><w:r><w:t>x</w:t></w:r></w:p>';
    expect(splitParagraph(truncated).pPr).toBe('');
    expect(splitParagraph(truncated).inner).toBe('<w:pPr><w:pStyle w:val="B"/><w:r><w:t>x</w:t></w:r>');
  });
});

describe('the seam never leaves the author (ADR-0032, ADR-0008)', () => {
  // PARA_SEAM is an AUTHORING device. Every other consumer of `op.wordDiffs` is
  // user- or LLM-visible — the AI-export change set, the exported compare
  // report, the change-list row preview — and `goToChange` uses the first mark
  // as a host-search anchor, which can never match document text while it holds
  // a Private Use Area code point. Nothing asserted this before.
  const ORIGINAL = [
    '9.1 The Vendor shall give notice of any Claim within 30 days.',
    'the notice shall specify the nature of the Claim in reasonable detail.',
  ].join('\n');
  const REVISED =
    '9.1 The Vendor shall give notice of any Claim within 30 days; and the notice shall specify the nature of the Claim in reasonable detail.';

  it('leaves no U+E000 in any op.wordDiffs, while still authoring the mark', async () => {
    const result = await runComparePipeline({
      originalParas: extractTextRecords(ORIGINAL),
      revisedParas: extractTextRecords(REVISED),
      config: CONFIG,
      author: AUTHOR,
      date: DATE,
    });

    // The merge really did happen — otherwise this asserts nothing.
    expect(result.authoredBodyXml).toContain('<w:pPr><w:rPr><w:del ');

    for (const op of result.ops) {
      for (const wd of op.wordDiffs ?? []) {
        expect(wd.text).not.toContain(PARA_SEAM);
      }
    }
    expect(result.authoredBodyXml).not.toContain(PARA_SEAM);
    expect(JSON.stringify(result.ops)).not.toContain(PARA_SEAM);
  });

  it('spells the struck paragraph mark as a newline, so both sides still reconstruct', async () => {
    const result = await runComparePipeline({
      originalParas: extractTextRecords(ORIGINAL),
      revisedParas: extractTextRecords(REVISED),
      config: CONFIG,
      author: AUTHOR,
      date: DATE,
    });
    const merged = result.ops.find((o) => (o.origRange?.end ?? 0) > (o.origRange?.start ?? 0));
    expect(merged).toBeDefined();
    const diffs = merged!.wordDiffs ?? [];
    expect(diffs.some((d) => d.op === 'del' && d.text.includes('\n'))).toBe(true);

    // The paragraph break is still THERE, spelled plainly — not silently
    // dropped, which would be a suppression.
    const rebuild = (side: 'ins' | 'del'): string =>
      diffs.filter((d) => d.op === 'eq' || d.op === side).map((d) => d.text).join('');
    expect(rebuild('del')).toContain('\n');
    expect(rebuild('ins')).not.toContain('\n');
  });
  it("does not rewrite the DOCUMENT'S own U+E000 in an ordinary one-to-one op", () => {
    // The sentinel is a Private Use Area code point, and `normaliseText` does not
    // strip those: broken-ToUnicode PDF extraction and verbatim pasted text both
    // yield them. `spellOutSeams` originally rewrote every U+E000 in every op, so
    // a document that legitimately contains one had it silently turned into a
    // paragraph break in the AI-export change set, the exported report and the
    // change-list preview. `detectMergesAndSplits` declines to re-pair such a
    // record (tested above), which is exactly what makes "claims more than one
    // record" a sound guard: a one-to-one op can only be carrying the
    // document's own character.
    const ORIGINAL = `The reference is \uE000 and shall not change.`;
    const REVISED = `The reference is \uE000 and must not change.`;

    return runComparePipeline({
      originalParas: extractTextRecords(ORIGINAL),
      revisedParas: extractTextRecords(REVISED),
      config: CONFIG,
      author: AUTHOR,
      date: DATE,
    }).then((result) => {
      const ops = result.ops.filter((o) => o.wordDiffs !== undefined);
      expect(ops).toHaveLength(1);
      expect(ops[0]!.origRange!.start).toBe(ops[0]!.origRange!.end);

      const all = ops[0]!.wordDiffs!.map((w) => w.text).join('');
      expect(all, 'the document character was rewritten').toContain('\uE000');
      expect(all, 'a paragraph break was invented').not.toContain('\n');

      // ADR-0008 still holds on both sides, sentinel and all.
      const side = (k: 'ins' | 'del'): string =>
        ops[0]!.wordDiffs!.filter((d) => d.op === 'eq' || d.op === k).map((d) => d.text).join('');
      expect(side('del')).toBe(ORIGINAL);
      expect(side('ins')).toBe(REVISED);
    });
  });

});
