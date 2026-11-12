/**
 * Feature 08 — Compare C1: pure OOXML authoring (author + revision-markup).
 * Hand-built DiffOp / ParaRecord fixtures (no preprocess/diff/classify imports).
 * Acceptance: T11 — correct elements, w:author/w:date/unique w:id, move w:name.
 */
import { describe, it, expect } from 'vitest';
import { XMLValidator } from 'fast-xml-parser';
import { authorBody, makeAuthorContext } from '@features/compare/output/author.js';
import { wrapDel, toDelText } from '@features/compare/output/revision-markup.js';
import { PARA_SEAM } from '@features/compare/diff/merge-split.js';
import type { DiffOp, ParaRecord } from '@features/compare/model.js';

const AUTHOR = 'Mutatis Compare';
const DATE = '2026-06-08T10:00:00Z';

function para(index: number, text: string, opts: Partial<ParaRecord> = {}): ParaRecord {
  return {
    index,
    partId: 'document',
    rawXml: `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`,
    normText: text,
    normTextAggressive: text.toLowerCase(),
    numberingPrefix: '',
    hash: `h${index}`,
    styleId: null,
    isTableCell: false,
    isSignatureBlock: false,
    isHeaderFooter: false,
    isFootnote: false,
    ...opts,
  };
}

function op(partial: Partial<DiffOp> & Pick<DiffOp, 'id' | 'type'>): DiffOp {
  return {
    origRange: null,
    revRange: null,
    partId: 'document',
    classification: 'substantiveText',
    classes: ['substantiveText'],
    author: AUTHOR,
    date: DATE,
    ...partial,
  };
}

describe('authorBody — insert', () => {
  it('wraps inserted runs in <w:ins> with correct id/author/date', () => {
    const revised = [para(0, 'New clause text')];
    const ops: DiffOp[] = [op({ id: 'op1', type: 'insert', revRange: { start: 0, end: 0 } })];
    const body = authorBody(ops, [], revised, makeAuthorContext(AUTHOR, DATE, 1));

    expect(body).toContain('<w:body>');
    expect(body).toContain('<w:ins w:id="1"');
    expect(body).toContain(`w:author="${AUTHOR}"`);
    expect(body).toContain(`w:date="${DATE}"`);
    expect(body).toContain('New clause text');
    // The <w:ins> wraps the run, inside the paragraph.
    expect(body).toMatch(/<w:p><w:ins[^>]*><w:r>/);
  });
});

describe('authorBody — equal emits the paragraph verbatim', () => {
  it('does NOT double a manually-numbered prefix already present in rawXml', () => {
    // The real extractor keeps rawXml whole: a manually-typed "(1)  " sits in the
    // run text, while numberingPrefix carries the normalised "(1) " stripped only
    // from the comparison text (§5.0). Re-injecting it would double the prefix
    // (regression: "(1) (1)  ASHWORTH"). Equal output must emit rawXml as-is.
    const revised = [
      para(0, 'ASHWORTH CAPITAL PTE. LTD.', {
        numberingPrefix: '(1) ',
        rawXml: '<w:p><w:r><w:t xml:space="preserve">(1)  ASHWORTH CAPITAL PTE. LTD.</w:t></w:r></w:p>',
      }),
    ];
    const ops: DiffOp[] = [op({ id: 'eq1', type: 'equal', revRange: { start: 0, end: 0 }, classification: 'formatting' })];
    const body = authorBody(ops, [], revised, makeAuthorContext(AUTHOR, DATE, 1));
    expect(body).toContain('(1)  ASHWORTH CAPITAL PTE. LTD.');
    expect(body).not.toContain('(1) (1)');
    // No revision markup on an equal paragraph.
    expect(body).not.toContain('<w:ins');
    expect(body).not.toContain('<w:del');
  });
});

describe('authorBody — delete', () => {
  it('wraps deleted runs in <w:del> and converts <w:t> to <w:delText>', () => {
    const original = [para(0, 'Removed clause')];
    const ops: DiffOp[] = [op({ id: 'del1', type: 'delete', origRange: { start: 0, end: 0 } })];
    const body = authorBody(ops, original, [], makeAuthorContext(AUTHOR, DATE, 1));
    expect(body).toContain('<w:del w:id="1"');
    expect(body).toContain('<w:delText>Removed clause</w:delText>');
    expect(body).not.toContain('<w:t>Removed clause</w:t>');
  });
});

describe('authorBody — replace with wordDiffs', () => {
  it('interleaves run-level <w:ins>/<w:del> from the WordDiff list', () => {
    const revised = [para(0, 'The price is £50,000 today')];
    const original = [para(0, 'The price is £40,000 today')];
    const ops: DiffOp[] = [
      op({
        id: 'rep1',
        type: 'replace',
        origRange: { start: 0, end: 0 },
        revRange: { start: 0, end: 0 },
        wordDiffs: [
          { op: 'eq', text: 'The price is ' },
          { op: 'del', text: '£40,000', atomic: 'amount' },
          { op: 'ins', text: '£50,000', atomic: 'amount' },
          { op: 'eq', text: ' today' },
        ],
        classification: 'numericalValue',
      }),
    ];
    const body = authorBody(ops, original, revised, makeAuthorContext(AUTHOR, DATE, 1));

    // eq tokens plain, del wrapped (delText), ins wrapped.
    expect(body).toContain('The price is ');
    expect(body).toContain('<w:del w:id="1"');
    expect(body).toContain('<w:delText>£40,000</w:delText>');
    expect(body).toContain('<w:ins w:id="2"');
    expect(body).toContain('£50,000');
    // amount is atomic — never split into digits.
    expect(body).not.toContain('<w:delText>4</w:delText>');
  });

  it('emits the WordDiff list in LIST ORDER — it never re-normalises the sides', () => {
    // The mark order the reader sees (§5.5, ADR-0043) is decided at Stage 6.5 of
    // the pipeline and carried here untouched. This pins the author's half of
    // that contract: hand it an ins-first list and `<w:ins>` comes out first;
    // hand it the del-first list diff-match-patch used to produce and `<w:del>`
    // comes out first. If `emitReplace` ever re-sorted a changed region, the
    // pipeline's decision would be silently overruled.
    //
    // NB this test does NOT fail if Stage 6.5's wiring is deleted — it never
    // runs the pipeline. That guard is
    // `tests/unit/compare/pipeline.mark-order.test.ts`, which drives
    // `runComparePipeline` end to end; both are needed and neither substitutes
    // for the other.
    const original = [para(0, 'an exclusive licence')];
    const revised = [para(0, 'a non-exclusive licence')];
    const build = (wordDiffs: DiffOp['wordDiffs']): string =>
      authorBody(
        [op({
          id: 'order',
          type: 'replace',
          origRange: { start: 0, end: 0 },
          revRange: { start: 0, end: 0 },
          ...(wordDiffs ? { wordDiffs } : {}),
        })],
        original,
        revised,
        makeAuthorContext(AUTHOR, DATE, 1),
      );

    const insFirst = build([
      { op: 'ins', text: 'a non-exclusive' },
      { op: 'del', text: 'an exclusive' },
      { op: 'eq', text: ' licence' },
    ]);
    expect(insFirst.indexOf('<w:ins ')).toBeLessThan(insFirst.indexOf('<w:del '));

    const delFirst = build([
      { op: 'del', text: 'an exclusive' },
      { op: 'ins', text: 'a non-exclusive' },
      { op: 'eq', text: ' licence' },
    ]);
    expect(delFirst.indexOf('<w:del ')).toBeLessThan(delFirst.indexOf('<w:ins '));
  });

  it('preserves xml:space on tokens with leading/trailing whitespace', () => {
    const revised = [para(0, 'a b')];
    const ops: DiffOp[] = [
      op({
        id: 'rep2',
        type: 'replace',
        revRange: { start: 0, end: 0 },
        origRange: { start: 0, end: 0 },
        wordDiffs: [
          { op: 'eq', text: 'a' },
          { op: 'ins', text: ' inserted ' },
          { op: 'eq', text: 'b' },
        ],
      }),
    ];
    const body = authorBody(ops, [para(0, 'ab')], revised, makeAuthorContext(AUTHOR, DATE, 1));
    expect(body).toContain('xml:space="preserve"');
  });
});

describe('authorBody — manual numbering prefix survives an edit (ADR-0028)', () => {
  // A replace REBUILDS the paragraph from wordDiffs, and those come from
  // normText — which had the manual prefix stripped so a renumbering cascade
  // would not read as a change. Without re-injection the clause loses the very
  // number the reader navigates and quotes by, and only on the clauses that
  // changed.
  const numbered = (index: number, prefix: string, body: string): ParaRecord =>
    para(index, `${prefix}${body}`, { numberingPrefix: prefix, normText: body });

  it('re-emits the clause number, unmarked, on an edited paragraph', () => {
    const original = [numbered(0, '9.1 ', 'The Vendor shall indemnify the Purchaser.')];
    const revised = [numbered(0, '9.1 ', 'The Vendor may indemnify the Purchaser.')];
    const ops: DiffOp[] = [
      op({
        id: 'rep-num',
        type: 'replace',
        origRange: { start: 0, end: 0 },
        revRange: { start: 0, end: 0 },
        wordDiffs: [
          { op: 'eq', text: 'The Vendor ' },
          { op: 'del', text: 'shall' },
          { op: 'ins', text: 'may' },
          { op: 'eq', text: ' indemnify the Purchaser.' },
        ],
        classification: 'obligationLanguage',
      }),
    ];
    const body = authorBody(ops, original, revised, makeAuthorContext(AUTHOR, DATE, 1));

    // The number is present exactly once, as a clean run — never struck,
    // inserted, or duplicated.
    expect(body).toContain('<w:r><w:t xml:space="preserve">9.1 </w:t></w:r>');
    expect(body.match(/9\.1 /g)).toHaveLength(1);
    expect(body).not.toContain('<w:delText>9.1 </w:delText>');
    // It leads the paragraph, ahead of the first word-diff run.
    expect(body.indexOf('9.1 ')).toBeLessThan(body.indexOf('The Vendor '));
  });

  it('does not double the number on an unchanged paragraph (regression: 2fd4eb5)', () => {
    // equal emits rawXml wholesale, which already carries the prefix.
    const recs = [numbered(0, '(1) ', 'Each Party acts severally.')];
    const ops: DiffOp[] = [
      op({ id: 'eq-num', type: 'equal', origRange: { start: 0, end: 0 }, revRange: { start: 0, end: 0 } }),
    ];
    const body = authorBody(ops, recs, recs, makeAuthorContext(AUTHOR, DATE, 1));
    expect(body.match(/\(1\) /g)).toHaveLength(1);
  });

  it('emits nothing extra for an auto-numbered paragraph (no literal prefix)', () => {
    const revised = [para(0, 'Auto-numbered clause text')];
    const ops: DiffOp[] = [
      op({
        id: 'rep-auto',
        type: 'replace',
        origRange: { start: 0, end: 0 },
        revRange: { start: 0, end: 0 },
        wordDiffs: [{ op: 'eq', text: 'Auto-numbered clause ' }, { op: 'ins', text: 'text' }],
      }),
    ];
    const body = authorBody(ops, [para(0, 'Auto-numbered clause ')], revised,
      makeAuthorContext(AUTHOR, DATE, 1));
    expect(body).not.toContain('xml:space="preserve"></w:t>');
    expect(body).toMatch(/<w:p>(<w:pPr>.*?<\/w:pPr>)?<w:r>/);
  });
});

describe('authorBody — move pair shares a w:name', () => {
  it('emits moveFrom + moveTo ranges with a shared w:name', () => {
    const original = [para(0, 'Indemnity clause verbatim')];
    const revised = [para(0, 'Indemnity clause verbatim')];
    const ops: DiffOp[] = [
      op({
        id: 'mvA',
        type: 'move',
        origRange: { start: 0, end: 0 },
        moveTargetId: 'mvB',
        classification: 'structural',
      }),
      op({
        id: 'mvB',
        type: 'move',
        revRange: { start: 0, end: 0 },
        moveTargetId: 'mvA',
        classification: 'structural',
      }),
    ];
    const body = authorBody(ops, original, revised, makeAuthorContext(AUTHOR, DATE, 1));

    expect(body).toContain('<w:moveFromRangeStart');
    expect(body).toContain('<w:moveFromRangeEnd');
    expect(body).toContain('<w:moveToRangeStart');
    expect(body).toContain('<w:moveToRangeEnd');
    expect(body).toContain('<w:moveFrom ');
    expect(body).toContain('<w:moveTo ');
    // Shared name = move-{lower id} = move-mvA for both endpoints.
    const names = [...body.matchAll(/w:name="([^"]+)"/g)].map(m => m[1]);
    expect(names.length).toBeGreaterThanOrEqual(2);
    expect(new Set(names).size).toBe(1);
    expect(names[0]).toBe('move-mvA');
  });
});

describe('authorBody — monotonic unique ids', () => {
  it('assigns a unique w:id to every revision element across the body', () => {
    const revised = [para(0, 'inserted one'), para(1, 'inserted two')];
    const original = [para(0, 'deleted one')];
    const ops: DiffOp[] = [
      op({ id: 'i1', type: 'insert', revRange: { start: 0, end: 0 } }),
      op({ id: 'd1', type: 'delete', origRange: { start: 0, end: 0 } }),
      op({ id: 'i2', type: 'insert', revRange: { start: 1, end: 1 } }),
    ];
    const body = authorBody(ops, original, revised, makeAuthorContext(AUTHOR, DATE, 1));
    const ids = [...body.matchAll(/w:id="(\d+)"/g)].map(m => Number(m[1]));
    // Four, not three: since ADR-0035 a whole-paragraph change also revises the
    // paragraph MARK, which takes an id of its own — but only where the mark is
    // revisable. Here that is i1 alone (runs + mark = 2); d1's paragraph is the
    // last of `original` and i2's the last of `revised`, so each emits runs only.
    expect(ids.length).toBe(4);
    expect(new Set(ids).size).toBe(ids.length); // all unique
    // Monotonically ASSIGNED — consecutive from the context's start value, none
    // wasted. Not ascending in document order: a paragraph mark's revision
    // serialises inside `w:pPr`, ahead of the runs whose id was minted first.
    expect([...ids].sort((a, b) => a - b)).toEqual([1, 2, 3, 4]);
  });
});

describe('revision-markup primitives', () => {
  it('wrapDel converts <w:t> to <w:delText> and adds attributes', () => {
    const out = wrapDel('<w:r><w:t>gone</w:t></w:r>', 7, AUTHOR, DATE);
    expect(out).toBe(`<w:del w:id="7" w:author="${AUTHOR}" w:date="${DATE}"><w:r><w:delText>gone</w:delText></w:r></w:del>`);
  });

  it('toDelText preserves xml:space attribute on the text element', () => {
    const out = toDelText('<w:r><w:t xml:space="preserve"> spaced </w:t></w:r>');
    expect(out).toContain('<w:delText xml:space="preserve"> spaced </w:delText>');
  });

  it('escapes ampersands in the author attribute', () => {
    const out = wrapDel('<w:r><w:t>x</w:t></w:r>', 1, 'A & B LLP', DATE);
    expect(out).toContain('w:author="A &amp; B LLP"');
  });
});

describe('authorBody — paragraph merge / split (ADR-0032)', () => {
  // These paths (`emitMergeSplit`, `splitAtSeams`, `prefixRun`) had no direct
  // coverage: the whole 30-case readability corpus exercises the MERGE branch
  // through exactly one case and the SPLIT branch not at all.
  const P1 = 'The Vendor shall give notice of any Claim within 30 days';
  const P2 = 'the notice shall specify the nature of the Claim';

  const ctx = (): ReturnType<typeof makeAuthorContext> => makeAuthorContext(AUTHOR, DATE, 1);

  it('authors a merge as two w:p, with the lost paragraph mark marked deleted', () => {
    const original = [para(0, `${P1}.`), para(1, P2)];
    const revised = [para(0, `${P1}; and ${P2}`)];
    const xml = authorBody(
      [
        op({
          id: 'o1',
          type: 'replace',
          origRange: { start: 0, end: 1 },
          revRange: { start: 0, end: 0 },
          wordDiffs: [
            { op: 'eq', text: P1 },
            { op: 'del', text: `.${PARA_SEAM}` },
            { op: 'ins', text: '; and ' },
            { op: 'eq', text: P2 },
          ],
        }),
      ],
      original,
      revised,
      ctx(),
    );

    // Two paragraphs until the change is ACCEPTED — a merge that cannot be
    // rejected is not a redline of the merge.
    expect(xml.match(/<w:p[\s>]/g)).toHaveLength(2);
    // The paragraph mark of the FIRST one carries the deletion.
    expect(xml).toContain('<w:pPr><w:rPr><w:del ');
    // The sentinel is an authoring device and must never reach the output.
    expect(xml).not.toContain(PARA_SEAM);
    // Both sides' words survive; nothing is suppressed (ADR-0008).
    expect(xml).toContain(P2);
    expect(xml).toContain('; and ');
  });

  it('authors a split as two w:p, with the new paragraph mark marked inserted', () => {
    const original = [para(0, `${P1}; and ${P2}`)];
    const revised = [para(0, `${P1}.`), para(1, P2)];
    const xml = authorBody(
      [
        op({
          id: 'o1',
          type: 'replace',
          origRange: { start: 0, end: 0 },
          revRange: { start: 0, end: 1 },
          wordDiffs: [
            { op: 'eq', text: P1 },
            { op: 'del', text: '; and ' },
            { op: 'ins', text: `.${PARA_SEAM}` },
            { op: 'eq', text: P2 },
          ],
        }),
      ],
      original,
      revised,
      ctx(),
    );

    expect(xml.match(/<w:p[\s>]/g)).toHaveLength(2);
    expect(xml).toContain('<w:pPr><w:rPr><w:ins ');
    expect(xml).not.toContain(PARA_SEAM);
    expect(xml).toContain(P1);
    expect(xml).toContain(P2);
  });

  it('rides a later paragraph\'s numbering prefix on the revision of the mark before it', () => {
    // On a merge that number really does disappear when the change is accepted,
    // so showing it unmarked would be a fabrication (ADR-0028).
    const original = [
      para(0, `${P1}.`, { numberingPrefix: '9.1' }),
      para(1, P2, { numberingPrefix: '9.2' }),
    ];
    const revised = [para(0, `${P1}; and ${P2}`, { numberingPrefix: '9.1' })];
    const xml = authorBody(
      [
        op({
          id: 'o1',
          type: 'replace',
          origRange: { start: 0, end: 1 },
          revRange: { start: 0, end: 0 },
          wordDiffs: [
            { op: 'eq', text: P1 },
            { op: 'del', text: `.${PARA_SEAM}` },
            { op: 'ins', text: '; and ' },
            { op: 'eq', text: P2 },
          ],
        }),
      ],
      original,
      revised,
      ctx(),
    );
    // 9.1 survives unmarked; 9.2 is struck along with the mark that precedes it.
    const paras = xml.split('</w:p>');
    expect(paras[0]).toContain('<w:r><w:t xml:space="preserve">9.1</w:t></w:r>');
    expect(paras[1]).toContain('<w:delText xml:space="preserve">9.2</w:delText>');
  });
});

describe('authorBody — a whole-paragraph change takes its paragraph MARK with it (ADR-0035)', () => {
  /**
   * The paragraph mark's own revision lives in `<w:pPr><w:rPr>`, not in a run.
   * Returns one flag per emitted `<w:p>`: is that paragraph's MARK marked with
   * a `w:<kind>` revision?
   */
  function markedParagraphs(xml: string, kind: 'del' | 'ins' = 'del'): boolean[] {
    const probe = new RegExp(`<w:pPr>(?:(?!</w:pPr>)[\\s\\S])*?<w:rPr>\\s*<w:${kind}\\b`);
    return xml
      .split('</w:p>')
      .filter((chunk) => chunk.includes('<w:p'))
      .map((chunk) => probe.test(chunk));
  }

  it('strikes the paragraph mark, so accepting removes the paragraph', () => {
    // Two paragraphs so the struck one is not the last of its part.
    const original = [para(0, 'Removed clause'), para(1, 'Surviving clause')];
    const ops: DiffOp[] = [op({ id: 'd1', type: 'delete', origRange: { start: 0, end: 0 } })];
    const body = authorBody(ops, original, [], makeAuthorContext(AUTHOR, DATE, 1));
    // The runs are struck …
    expect(body).toContain('<w:delText>Removed clause</w:delText>');
    // … AND the mark, which is what stops an empty paragraph being left behind.
    expect(markedParagraphs(body)).toEqual([true]);
  });

  it('leaves a sectPr-bearing mark alone — striking it would re-paginate the document', () => {
    // The last paragraph of a section carries that section's page setup. Losing
    // it to a clause deletion is a worse defect than an empty paragraph.
    const original = [
      para(0, 'Removed clause', {
        rawXml:
          '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr>' +
          '<w:r><w:t>Removed clause</w:t></w:r></w:p>',
      }),
      para(1, 'Surviving clause'),
    ];
    const ops: DiffOp[] = [op({ id: 'd1', type: 'delete', origRange: { start: 0, end: 0 } })];
    const body = authorBody(ops, original, [], makeAuthorContext(AUTHOR, DATE, 1));
    expect(body).toContain('<w:delText>Removed clause</w:delText>');
    expect(markedParagraphs(body)).toEqual([false]);
    expect(body).toContain('<w:sectPr>');
  });

  it('leaves the LAST paragraph of a part alone — there is nothing to merge into', () => {
    const original = [para(0, 'Only clause')];
    const ops: DiffOp[] = [op({ id: 'd1', type: 'delete', origRange: { start: 0, end: 0 } })];
    const body = authorBody(ops, original, [], makeAuthorContext(AUTHOR, DATE, 1));
    expect(body).toContain('<w:delText>Only clause</w:delText>');
    expect(markedParagraphs(body)).toEqual([false]);
  });

  it('treats each part separately — a table cell keeps its final paragraph', () => {
    // `document` has a later paragraph, so its deleted one IS markable; the cell
    // paragraph is the last of `tableCell:r1c1`, so it is not. One op, two
    // answers — the guard is per-part, not per-document.
    const original = [
      para(0, 'Cell clause', { partId: 'tableCell:r1c1', isTableCell: true }),
      para(1, 'Body clause', { partId: 'document' }),
      para(2, 'Body tail', { partId: 'document' }),
    ];
    const ops: DiffOp[] = [
      op({ id: 'd1', type: 'delete', origRange: { start: 0, end: 0 }, partId: 'tableCell:r1c1' }),
      op({ id: 'd2', type: 'delete', origRange: { start: 1, end: 1 } }),
    ];
    const body = authorBody(ops, original, [], makeAuthorContext(AUTHOR, DATE, 1));
    expect(markedParagraphs(body)).toEqual([false, true]);
  });

  it('keeps every w:id unique once the mark takes one of its own', () => {
    const original = [para(0, 'Removed clause'), para(1, 'Surviving clause')];
    const ops: DiffOp[] = [op({ id: 'd1', type: 'delete', origRange: { start: 0, end: 0 } })];
    const body = authorBody(ops, original, [], makeAuthorContext(AUTHOR, DATE, 1));
    const ids = [...body.matchAll(/w:id="(\d+)"/g)].map((m) => Number(m[1]));
    expect(ids.length).toBe(2); // one for the runs, one for the mark
    expect(new Set(ids).size).toBe(ids.length);
  });

  it('emits well-formed XML with the mark revision inside a single w:pPr/w:rPr', () => {
    const original = [
      para(0, 'Removed clause', {
        // Already carries paragraph properties AND a paragraph-mark rPr, the
        // shape that `CT_PPr`'s "one rPr, ordered first" rule makes delicate.
        rawXml:
          '<w:p><w:pPr><w:pStyle w:val="Body"/><w:rPr><w:b/></w:rPr></w:pPr>' +
          '<w:r><w:t>Removed clause</w:t></w:r></w:p>',
      }),
      para(1, 'Surviving clause'),
    ];
    const ops: DiffOp[] = [op({ id: 'd1', type: 'delete', origRange: { start: 0, end: 0 } })];
    const body = authorBody(ops, original, [], makeAuthorContext(AUTHOR, DATE, 1));
    expect(XMLValidator.validate(`<w:root xmlns:w="urn:w">${body}</w:root>`)).toBe(true);
    // Exactly one rPr in the pPr, with the revision first inside it (CT_ParaRPr).
    expect(body).toContain('<w:rPr><w:del w:id=');
    expect((body.match(/<w:pPr>/g) ?? []).length).toBe(1);
    expect(body).toContain('<w:b/>');
  });

  // ── the mirror direction: reject-all must not leave blanks either ──────────

  it('marks an INSERTED paragraph mark, so rejecting removes the paragraph', () => {
    // Symmetry: a reviewer who rejects the whole redline must get their own
    // document back. Wrapping only the runs would strip the text and leave an
    // empty paragraph for every clause the other side added.
    const revised = [para(0, 'Added clause'), para(1, 'Surviving clause')];
    const ops: DiffOp[] = [op({ id: 'i1', type: 'insert', revRange: { start: 0, end: 0 } })];
    const body = authorBody(ops, [], revised, makeAuthorContext(AUTHOR, DATE, 1));
    expect(body).toContain('<w:t>Added clause</w:t>');
    expect(markedParagraphs(body, 'ins')).toEqual([true]);
  });

  it('leaves the LAST paragraph of a part unmarked on insert too', () => {
    // A body — and a table cell — must end with a paragraph, so this mark has
    // to survive a reject regardless of who added the text.
    const revised = [para(0, 'Added clause')];
    const ops: DiffOp[] = [op({ id: 'i1', type: 'insert', revRange: { start: 0, end: 0 } })];
    const body = authorBody(ops, [], revised, makeAuthorContext(AUTHOR, DATE, 1));
    expect(body).toContain('<w:t>Added clause</w:t>');
    expect(markedParagraphs(body, 'ins')).toEqual([false]);
  });

  it('emits well-formed XML for an inserted mark, ids still unique', () => {
    const revised = [
      para(0, 'Added clause', {
        rawXml:
          '<w:p><w:pPr><w:pStyle w:val="Body"/><w:rPr><w:i/></w:rPr></w:pPr>' +
          '<w:r><w:t>Added clause</w:t></w:r></w:p>',
      }),
      para(1, 'Surviving clause'),
    ];
    const ops: DiffOp[] = [op({ id: 'i1', type: 'insert', revRange: { start: 0, end: 0 } })];
    const body = authorBody(ops, [], revised, makeAuthorContext(AUTHOR, DATE, 1));
    expect(XMLValidator.validate(`<w:root xmlns:w="urn:w">${body}</w:root>`)).toBe(true);
    expect(body).toContain('<w:rPr><w:ins w:id=');
    const ids = [...body.matchAll(/w:id="(\d+)"/g)].map((m) => Number(m[1]));
    expect(ids.length).toBe(2);
    expect(new Set(ids).size).toBe(ids.length);
    expect(body).toContain('<w:i/>');
  });
});
