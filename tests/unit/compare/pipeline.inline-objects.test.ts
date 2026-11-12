/**
 * ADR-0045 — what is not text survives a changed paragraph, and a change to it
 * is reported.
 *
 * Three silent losses, found by probe on 2026-10-17 through `runPipeline`:
 *
 *  1. An OMML equation was invisible to paragraph extraction (`<m:t>` was never
 *     read), so changing ONLY the formula gave `summary.total` 0, ops
 *     `['equal', …]`, and the revised formula in `authoredBodyXml` with no
 *     tracked markup — ADR-0008 suppression.
 *  2. A changed paragraph is rebuilt from its word diff, which is text, so its
 *     footnote reference, hyperlink and complex-field structure went with the
 *     rebuild (a REF field became frozen text).
 *  3. A text box anchored in a changed paragraph was poured into the body text
 *     (`"the Price.Boxed note"`) and the box disappeared.
 *
 * Each describe below pins one of them at the single-pair seam every add-in
 * user takes, plus the shapes the fix had to get right on the way: an edit
 * INSIDE a link or field, a text box whose words change, the insertion-first
 * order (ADR-0043) when an object meets plain text, and the paragraphs that
 * must NOT change — an edited paragraph with no objects is authored byte for
 * byte as before.
 *
 * Every test also checks the redline both ways: accept every revision and it
 * reads as the revised document, reject every one and it reads as the
 * original. That is the never-suppress property (ADR-0008) stated as an
 * assertion, and it is the check `inline-objects.ts` runs on itself.
 */
import { describe, it, expect } from 'vitest';
import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import { runPipeline } from '@features/compare/pipeline.js';
import { flattenTrackedChanges, extractParaRecords } from '@features/compare/preprocess/index.js';
import { authorObjectAwareRuns } from '@features/compare/output/inline-objects.js';
import { paragraphNode, paragraphText } from '@features/compare/preprocess/inline-units.js';
import { normaliseText, stripNumberingPrefix } from '@features/compare/preprocess/normalise.js';
import type { ParaRecord, WordDiff } from '@features/compare/model.js';
import {
  inlineParts, para, run, math, footnoteRef, hyperlink, refField, bookmarked, textBox, alternateTextBox,
  endnoteRef, inlinePicture, embeddedObject,
} from '../../fixtures/inline-objects-docx.js';
import { applyFilters, previewText } from '@features/compare/components/ChangeList.js';
import { buildCompareReport } from '@features/compare/export/report-model.js';

const CONFIG = getCompareConfig(SG_PROFILE);
const AUTHOR = 'Mutatis Compare';
const DATE = '2026-10-17T00:00:00Z';

async function compare(original: string[], revised: string[]) {
  return runPipeline({
    originalParts: inlineParts(...original),
    revisedParts: inlineParts(...revised),
    config: CONFIG,
    author: AUTHOR,
    date: DATE,
  });
}

type Result = Awaited<ReturnType<typeof compare>>;

/**
 * A paragraph's text as the DOCUMENT spells it: no boundary spaces around an
 * equation or text box (those are `normText`'s, and they can stand in for a
 * real space the author lost — `"r=3%per annum"` passed a `normText` check).
 */
function strictText(rawXml: string): string {
  return stripNumberingPrefix(normaliseText(paragraphText(paragraphNode(rawXml)!, false))).body;
}

/** Containers a whole-object deletion empties: once accepted, only the shell is left. */
const SHELLS = /<(w:hyperlink|w:fldSimple|m:oMath|w:sdtContent)(?:\s[^>]*)?(?:\/>|>(?:<w:sdtPr>.*?<\/w:sdtPr>)?<\/\1>)/g;

/**
 * Objects a text comparison cannot see: counted, so a lost or doubled one
 * shows. An emptied shell is not an object (see SHELLS), and is dropped first —
 * repeatedly, since dropping an empty equation can empty the content control
 * around it.
 */
function census(raw: string, bookmarks: boolean): Record<string, number> {
  let xml = raw;
  for (let prev = ''; prev !== xml; ) {
    prev = xml;
    xml = xml.replace(SHELLS, '');
  }
  const n = (re: RegExp) => (xml.match(re) ?? []).length;
  const open = (tag: string) => n(new RegExp(`<${tag}[\\s/>]`, 'g'));
  return {
    footnote: n(/<w:footnoteReference[\s/>]/g),
    endnote: n(/<w:endnoteReference[\s/>]/g),
    object: open('w:object'),
    fieldBegin: n(/w:fldCharType="begin"/g),
    instr: n(/<w:instrText[\s>]/g),
    hyperlink: open('w:hyperlink'),
    fldSimple: open('w:fldSimple'),
    math: open('m:oMath'),
    drawing: open('w:drawing'),
    textBox: open('w:txbxContent'),
    sdt: open('w:sdtContent'),
    ...(bookmarks ? { bookmark: n(/<w:bookmarkStart[\s/>]/g) } : {}),
  };
}

/**
 * Accept-all reads as the revised document and reject-all as the original —
 * paragraph by paragraph in the document's own spelling, and object by object.
 * Bookmarks are compared on the accepted view only: they are re-emitted from
 * the revised side and never revised (ADR-0045 §2).
 */
function expectRoundTrip(result: Result): void {
  for (const [policy, source] of [
    ['acceptAll', result.revisedParas],
    ['rejectAll', result.originalParas],
  ] as const) {
    const flat = flattenTrackedChanges(result.authoredBodyXml, policy);
    const authored = extractParaRecords(flat, 'document', CONFIG).map((p) => strictText(p.rawXml));
    const expected = source.map((p: ParaRecord) => strictText(p.rawXml));
    const nonEmpty = (xs: string[]) => xs.filter((t) => t !== '');
    expect(nonEmpty(authored), policy).toEqual(nonEmpty(expected));
    const sourceXml = source.map((p: ParaRecord) => p.rawXml).join('');
    expect(census(flat, policy === 'acceptAll'), policy).toEqual(census(sourceXml, policy === 'acceptAll'));
  }
}

/**
 * The first authored `<w:p>` — the pair under test always comes first. Matched
 * by depth, because a text box nests whole paragraphs inside its host.
 */
function firstPara(result: Result): string {
  const xml = result.authoredBodyXml;
  const start = xml.indexOf('<w:p');
  const tag = /<w:p(?:\s[^>]*)?>|<\/w:p>/g;
  tag.lastIndex = start;
  let depth = 0;
  for (let m = tag.exec(xml); m !== null; m = tag.exec(xml)) {
    depth += m[0].startsWith('</') ? -1 : 1;
    if (depth === 0) return xml.slice(start, m.index + m[0].length);
  }
  throw new Error('unbalanced paragraph');
}

const count = (xml: string, needle: string): number => xml.split(needle).length - 1;

// ---------------------------------------------------------------------------
// 1. Equations
// ---------------------------------------------------------------------------

describe('ADR-0045 (1): an OMML equation is compared and its change is tracked', () => {
  it('an equation-only change inside a sentence is reported (it was summary.total 0)', async () => {
    const r = await compare(
      [para(run('The interest rate is ') + math('r=2%') + run(' per annum.'))],
      [para(run('The interest rate is ') + math('r=3%') + run(' per annum.'))],
    );
    expect(r.summary.total).toBe(1);
    expect(r.ops.map((o) => o.type)).toEqual(['replace', 'equal']);
    // The equation's text is part of what is compared now.
    expect(r.revisedParas[0]!.normText).toContain('r=3%');
    const p = firstPara(r);
    // Both formulas present, each as a tracked revision INSIDE its own
    // `<m:oMath>` (where Word records math revisions), insertion first.
    expect(p).toMatch(/<m:oMath><w:ins [^>]*><m:r><m:t>r=3%<\/m:t><\/m:r><\/w:ins><\/m:oMath>/);
    expect(p).toMatch(/<m:oMath><w:del [^>]*><m:r><m:t>r=2%<\/m:t><\/m:r><\/w:del><\/m:oMath>/);
    expect(p.indexOf('r=3%')).toBeLessThan(p.indexOf('r=2%'));
    // The revised formula is never on the page untracked.
    expect(p).not.toMatch(/<\/w:ins>[^]*?<m:t>r=3%/);
    expectRoundTrip(r);
  });

  it('a paragraph that is only an equation reports the change too', async () => {
    const r = await compare([para(math('E=mc2'))], [para(math('E=mc3'))]);
    expect(r.summary.total).toBe(1);
    expect(r.authoredBodyXml).toMatch(/<w:ins [^>]*><m:r><m:t>E=mc3<\/m:t>/);
    expect(r.authoredBodyXml).toMatch(/<w:del [^>]*><m:r><m:t>E=mc2<\/m:t>/);
    expectRoundTrip(r);
  });

  it('an UNCHANGED equation in an edited paragraph is kept verbatim and unmarked', async () => {
    const r = await compare(
      [para(run('Where ') + math('x=a+b') + run(', the Seller shall pay.'))],
      [para(run('Where ') + math('x=a+b') + run(', the Seller must pay.'))],
    );
    const p = firstPara(r);
    expect(count(p, math('x=a+b'))).toBe(1);
    expect(p).not.toMatch(/<m:oMath><w:(ins|del)/);
    expectRoundTrip(r);
  });

  it('an unchanged paragraph carrying an equation is still emitted verbatim', async () => {
    const same = para(run('Where ') + math('x=a+b') + run(' applies.'));
    const r = await compare([same], [same]);
    expect(r.summary.total).toBe(0);
    expect(r.authoredBodyXml.startsWith(`<w:body>${same}`)).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// 2. Footnote references, hyperlinks, complex fields, bookmarks
// ---------------------------------------------------------------------------

describe('ADR-0045 (2): a changed paragraph keeps its footnote, link and field structure', () => {
  it('keeps a footnote reference, hyperlink and REF field unchanged around the edit', async () => {
    const body = (verb: string) =>
      para(
        run(`The Seller ${verb} pay the Price`) + footnoteRef() + run(' to the account at ') +
          hyperlink('the payment portal') + run(' under Clause ') + refField('4.2') + run('.'),
      );
    const r = await compare([body('shall')], [body('must')]);
    expect(r.summary.total).toBe(1);
    const p = firstPara(r);
    // Each object once, verbatim, outside any revision.
    expect(count(p, footnoteRef())).toBe(1);
    expect(count(p, hyperlink('the payment portal'))).toBe(1);
    expect(count(p, refField('4.2'))).toBe(1);
    // The edit itself is the only tracked change.
    expect(p).toMatch(/<w:ins [^>]*><w:r><w:t>must<\/w:t><\/w:r><\/w:ins><w:del [^>]*><w:r><w:delText>shall<\/w:delText>/);
    expect(count(p, '<w:ins ')).toBe(1);
    expect(count(p, '<w:del ')).toBe(1);
    // And the structure sits where it did: footnote after "Price", field after "Clause".
    expect(p.indexOf('Price')).toBeLessThan(p.indexOf('w:footnoteReference'));
    expect(p.indexOf(' under Clause ')).toBeLessThan(p.indexOf('w:fldCharType="begin"'));
    expectRoundTrip(r);
  });

  it('an edit INSIDE a hyperlink is tracked inside the kept link, in its style', async () => {
    const r = await compare(
      [para(run('Notices go to ') + hyperlink('the Seller portal') + run(' only.'))],
      [para(run('Notices go to ') + hyperlink('the Buyer portal') + run(' only.'))],
    );
    const p = firstPara(r);
    expect(count(p, '<w:hyperlink ')).toBe(1);
    const link = /<w:hyperlink [^>]*>(.*?)<\/w:hyperlink>/.exec(p)![1]!;
    expect(link).toMatch(/<w:ins [^>]*><w:r><w:rPr><w:rStyle w:val="Hyperlink"\/><\/w:rPr><w:t>Buyer<\/w:t>/);
    expect(link).toMatch(/<w:del [^>]*><w:r><w:rPr><w:rStyle w:val="Hyperlink"\/><\/w:rPr><w:delText>Seller<\/w:delText>/);
    // `<w:ins>` may not wrap a hyperlink (EG_PContent) — never authored that way.
    expect(p).not.toMatch(/<w:(ins|del) [^>]*><w:hyperlink/);
    expectRoundTrip(r);
  });

  it('a changed REF result is tracked inside the kept field', async () => {
    const r = await compare(
      [para(run('Subject to ') + refField('Clause 11.2') + run(', the Buyer may terminate.'))],
      [para(run('Subject to ') + refField('Clause 12.2') + run(', the Buyer may terminate.'))],
    );
    const p = firstPara(r);
    expect(count(p, 'w:fldCharType="begin"')).toBe(1);
    expect(count(p, 'w:fldCharType="separate"')).toBe(1);
    expect(count(p, 'w:fldCharType="end"')).toBe(1);
    expect(count(p, '<w:instrText')).toBe(1);
    const sep = p.indexOf('w:fldCharType="separate"');
    const end = p.indexOf('w:fldCharType="end"');
    const result = p.slice(sep, end);
    expect(result).toMatch(/<w:ins [^>]*><w:r><w:t>12\.2<\/w:t>/);
    expect(result).toMatch(/<w:del [^>]*><w:r><w:delText>11\.2<\/w:delText>/);
    expectRoundTrip(r);
  });

  it('a footnote reference removed alongside an edit is shown struck', async () => {
    const r = await compare(
      [para(run('The Seller shall pay.') + footnoteRef() + run(' Then it may go.'))],
      [para(run('The Seller must pay. Then it may go.'))],
    );
    expect(firstPara(r)).toMatch(/<w:del [^>]*><w:r><w:rPr><w:rStyle w:val="FootnoteReference"\/><\/w:rPr><w:footnoteReference/);
    expectRoundTrip(r);
  });

  it('keeps a bookmark another clause can cross-refer to', async () => {
    const r = await compare(
      [para(bookmarked(run('The Seller shall deliver the Goods.')))],
      [para(bookmarked(run('The Seller must deliver the Goods.')))],
    );
    const p = firstPara(r);
    expect(count(p, '<w:bookmarkStart w:id="0" w:name="_Ref10001"/>')).toBe(1);
    expect(count(p, '<w:bookmarkEnd w:id="0"/>')).toBe(1);
    expect(p.indexOf('bookmarkStart')).toBeLessThan(p.indexOf('The Seller'));
    expect(p.indexOf('Goods.')).toBeLessThan(p.indexOf('bookmarkEnd'));
    expectRoundTrip(r);
  });

  it('a link that became plain text reads insertion first (ADR-0043)', async () => {
    const r = await compare(
      [para(run('Visit ') + hyperlink('the portal') + run(' now and pay.'))],
      [para(run('Visit the portal now and pay later.'))],
    );
    const p = firstPara(r);
    expect(p.indexOf('<w:t>the portal</w:t>')).toBeGreaterThan(-1);
    expect(p.indexOf('<w:t>the portal</w:t>')).toBeLessThan(p.indexOf('<w:delText xml:space="preserve">the portal'));
    // w:ids stay in document order (§5.5).
    const ids = [...p.matchAll(/<w:(?:ins|del) w:id="(\d+)"/g)].map((m) => Number(m[1]));
    expect(ids).toEqual([...ids].sort((a, b) => a - b));
    expectRoundTrip(r);
  });

  it('a region the walk revises itself still reads insertion first (ADR-0043)', async () => {
    // The word diff only strikes "promptly"; the walk adds the link-to-text
    // change beside it, and that insertion must still lead the region.
    const r = await compare(
      [para(run('Pay promptly ') + hyperlink('the Seller') + run(' on Completion.'))],
      [para(run('Pay the Seller on Completion.'))],
    );
    const p = firstPara(r);
    expect(p.indexOf('<w:ins ')).toBeGreaterThan(-1);
    expect(p.indexOf('<w:ins ')).toBeLessThan(p.indexOf('<w:del '));
    const ids = [...p.matchAll(/<w:(?:ins|del) w:id="(\d+)"/g)].map((m) => Number(m[1]));
    expect(ids).toEqual([...ids].sort((a, b) => a - b));
    expectRoundTrip(r);
  });

  it('… with a bookmark between the deletion and the insertion too', async () => {
    const bm = (inner: string) => `<w:bookmarkStart w:id="7" w:name="_Ref10003"/>${inner}<w:bookmarkEnd w:id="7"/>`;
    const r = await compare(
      [para(run('Pay promptly ') + hyperlink('the Seller') + run(' on Completion.'))],
      [para(run('Pay ') + bm(run('the Seller')) + run(' on Completion.'))],
    );
    const p = firstPara(r);
    expect(p.indexOf('<w:ins ')).toBeLessThan(p.indexOf('<w:del '));
    // The bookmark still opens on the inserted words in the accepted view.
    expect(flattenTrackedChanges(p, 'acceptAll')).toMatch(/<w:bookmarkStart [^>]*\/><w:r><w:t[^>]*>the Seller/);
    expectRoundTrip(r);
  });

  it('RECORDS ITS LIMIT: a kept footnote between the deletion and the insertion keeps both where they were', async () => {
    // Insertion first would carry "promptly" past the footnote, and rejecting
    // every change would then put the footnote before it (ADR-0045 § Accepted limits 13).
    const r = await compare(
      [para(run('Pay promptly ') + footnoteRef() + hyperlink('the Seller') + run(' on Completion.'))],
      [para(run('Pay ') + footnoteRef() + run('the Seller on Completion.'))],
    );
    const p = firstPara(r);
    const rejected = flattenTrackedChanges(p, 'rejectAll');
    expect(rejected.indexOf('promptly')).toBeLessThan(rejected.indexOf('footnoteReference'));
    const accepted = flattenTrackedChanges(p, 'acceptAll');
    expect(accepted.indexOf('footnoteReference')).toBeLessThan(accepted.indexOf('the Seller'));
    expect(p.indexOf('<w:del ')).toBeLessThan(p.indexOf('<w:ins ')); // the recorded exception to ADR-0043
    expectRoundTrip(r);
  });

  it('a whole inserted or deleted paragraph revises INSIDE its links and fields', async () => {
    const extra = para(run('Notices go to ') + hyperlink('the portal') + run(' under Clause ') + refField('3') + math('n=1'));
    const keep = para(run('The Buyer shall pay.'));
    const ins = await compare([keep], [keep, extra]);
    expect(ins.authoredBodyXml).not.toMatch(/<w:ins [^>]*><w:hyperlink/);
    expect(ins.authoredBodyXml).toMatch(/<w:hyperlink [^>]*><w:ins /);
    expect(ins.authoredBodyXml).toMatch(/<m:oMath><w:ins /);
    expectRoundTrip(ins);

    const del = await compare([keep, extra], [keep]);
    expect(del.authoredBodyXml).not.toMatch(/<w:del [^>]*><w:hyperlink/);
    // A deleted field's instruction is deleted too (ECMA-376 §17.16.13).
    expect(del.authoredBodyXml).toMatch(/<w:delInstrText xml:space="preserve"> REF [^<]* <\/w:delInstrText>/);
    expect(del.authoredBodyXml).not.toMatch(/<\/?w:instrText/);
    expectRoundTrip(del);
  });
});

// ---------------------------------------------------------------------------
// 3. Text boxes
// ---------------------------------------------------------------------------

describe('ADR-0045 (3): a text box in a changed paragraph keeps its box', () => {
  it('an unchanged text box is kept verbatim, not poured into the sentence', async () => {
    const r = await compare(
      [para(run('The Seller shall pay the Price.') + textBox('Boxed note'))],
      [para(run('The Seller must pay the Price.') + textBox('Boxed note'))],
    );
    expect(r.summary.total).toBe(1);
    const p = firstPara(r);
    expect(count(p, textBox('Boxed note'))).toBe(1);
    expect(p).not.toContain('Price.Boxed');
    // Its words read as their own words in the compared text, too.
    expect(r.revisedParas[0]!.normText).toBe('The Seller must pay the Price. Boxed note');
    expectRoundTrip(r);
  });

  it('a text box written with a VML fallback is read once, and kept whole', async () => {
    const r = await compare(
      [para(run('The Seller shall pay.') + alternateTextBox('Boxed note'))],
      [para(run('The Seller must pay.') + alternateTextBox('Boxed note'))],
    );
    // Before: "… pay. Boxed note Boxed note" — the choice and the fallback both read.
    expect(r.revisedParas[0]!.normText).toBe('The Seller must pay. Boxed note');
    expect(count(firstPara(r), alternateTextBox('Boxed note'))).toBe(1);
    expectRoundTrip(r);
  });

  it('a change to the words in a text box is reported and keeps both boxes', async () => {
    const r = await compare(
      [para(run('The Seller shall pay.') + textBox('Payment is due monthly'))],
      [para(run('The Seller shall pay.') + textBox('Payment is due quarterly'))],
    );
    expect(r.summary.total).toBe(1);
    const p = firstPara(r);
    expect(count(p, '<w:txbxContent>')).toBe(2);
    // The revised box as an insertion and the original as a deletion, each with
    // its own words marked inside it so a reject restores them.
    expect(p).toMatch(/<w:ins [^>]*><w:r><w:drawing>.*<w:txbxContent><w:p><w:ins [^>]*><w:r><w:t>Payment is due quarterly/);
    expect(p).toMatch(/<w:del [^>]*><w:r><w:drawing>.*<w:txbxContent><w:p><w:del [^>]*><w:r><w:delText>Payment is due monthly/);
    expect(p.indexOf('quarterly')).toBeLessThan(p.indexOf('monthly'));
    expectRoundTrip(r);
  });
});

// ---------------------------------------------------------------------------
// Shapes the first draft got wrong (spec-validator, 2026-10-17) — each passed
// the draft's own check and the draft's tests
// ---------------------------------------------------------------------------

describe('ADR-0045: content nested in a container, spacing, and doubled shells', () => {
  it('keeps the real space after an unchanged equation or text box', async () => {
    const r = await compare(
      [para(run('Let ') + math('x') + run(' be the Seller amount.')), para(run('Pay.') + textBox('Note') + run(' Then the Seller goes.'))],
      [para(run('Let ') + math('x') + run(' be the Buyer amount.')), para(run('Pay.') + textBox('Note') + run(' Then the Buyer goes.'))],
    );
    const flat = flattenTrackedChanges(r.authoredBodyXml, 'acceptAll');
    const texts = extractParaRecords(flat, 'document', CONFIG).map((p) => strictText(p.rawXml));
    expect(texts[0]).toBe('Let x be the Buyer amount.'); // was "Let xbe …"
    expect(texts[1]).toBe('Pay.Note Then the Buyer goes.'); // was "Pay.NoteThen …"
    expectRoundTrip(r);
  });

  it('a footnote INSIDE an edited hyperlink survives (the link is replaced whole)', async () => {
    const linkWithNote = (t: string) =>
      `<w:hyperlink r:id="rId9">${run(t)}${footnoteRef()}</w:hyperlink>`;
    const r = await compare(
      [para(run('Notices go to ') + linkWithNote('the Seller portal') + run('.'))],
      [para(run('Notices go to ') + linkWithNote('the Buyer portal') + run('.'))],
    );
    expect(r.summary.total).toBe(1);
    expectRoundTrip(r); // footnote count 1 in both views
  });

  it('a footnote only the ORIGINAL link carried is shown struck', async () => {
    const r = await compare(
      [para(run('Notices go to ') + `<w:hyperlink r:id="rId9">${run('the Seller portal')}${footnoteRef()}</w:hyperlink>` + run('.'))],
      [para(run('Notices go to ') + hyperlink('the Buyer portal') + run('.'))],
    );
    expect(r.summary.total).toBe(1);
    expectRoundTrip(r); // footnote 0 accepted, 1 rejected
  });

  it('a footnote inside an edited complex HYPERLINK field result survives', async () => {
    const field = (t: string) =>
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> HYPERLINK "https://example.invalid/" </w:instrText></w:r>' +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${run(t)}${footnoteRef()}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const r = await compare(
      [para(run('See ') + field('the Seller portal') + run('.'))],
      [para(run('See ') + field('the Buyer portal') + run('.'))],
    );
    expect(r.summary.total).toBe(1);
    expectRoundTrip(r);
  });

  it('a PAGEREF inside an edited TOC-style link survives', async () => {
    const entry = (t: string) =>
      `<w:hyperlink w:anchor="_Toc10001">${run(t)}<w:r><w:tab/></w:r>` +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF _Toc10001 \\h </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>3</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink>';
    const r = await compare([para(entry('Payment of the Price'))], [para(entry('Payment of the Purchase Price'))]);
    expect(r.summary.total).toBe(1);
    expect(flattenTrackedChanges(r.authoredBodyXml, 'acceptAll')).toContain('PAGEREF');
    expectRoundTrip(r);
  });

  it('an equation inside an edited content control survives', async () => {
    const control = (t: string) =>
      `<w:sdt><w:sdtPr><w:alias w:val="Rate"/></w:sdtPr><w:sdtContent>${run(t)}${math('n=4')}</w:sdtContent></w:sdt>`;
    const r = await compare([para(control('The rate is fixed at '))], [para(control('The rate is set at '))]);
    expect(r.summary.total).toBe(1);
    expectRoundTrip(r); // math count 1 in both views
  });

  it('an edited content control of plain text is diffed inside its kept content', async () => {
    const control = (t: string) =>
      `<w:sdt><w:sdtPr><w:alias w:val="Party"/></w:sdtPr><w:sdtContent>${run(t)}</w:sdtContent></w:sdt>`;
    const r = await compare(
      [para(run('This deed is made by ') + control('Seller Holdings Limited') + run('.'))],
      [para(run('This deed is made by ') + control('Seller Group Limited') + run('.'))],
    );
    const p = firstPara(r);
    expect(count(p, '<w:sdtContent>')).toBe(1);
    expect(/<w:sdtContent>(.*)<\/w:sdtContent>/.exec(p)![1]).toMatch(/<w:ins [^>]*><w:r><w:t>Group<\/w:t>/);
    expectRoundTrip(r);
  });

  it('a link or field is never emitted twice when a plain edit falls inside its stretch', async () => {
    const link = await compare(
      [para(run('Pay the ') + hyperlink('Seller portal') + run(' fee.'))],
      [para(run('Pay ') + hyperlink('Buyer portal') + run(' fee.'))],
    );
    expect(census(flattenTrackedChanges(link.authoredBodyXml, 'acceptAll'), false).hyperlink).toBe(1);
    expectRoundTrip(link);

    const field = await compare(
      [para(run('Subject to the ') + refField('Seller Clause 4')  + run(' terms.'))],
      [para(run('Subject to ') + refField('Buyer Clause 4') + run(' terms.'))],
    );
    expect(census(flattenTrackedChanges(field.authoredBodyXml, 'acceptAll'), false).fieldBegin).toBe(1);
    expectRoundTrip(field);
  });

  it('a nested field is replaced whole, never rebuilt from its result text', async () => {
    const nested = (t: string) =>
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> IF 1 = 1 "</w:instrText></w:r>' +
      refField('4.2') + '<w:r><w:instrText xml:space="preserve">" </w:instrText></w:r>' +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${run(t)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const r = await compare(
      [para(run('See ') + nested('Clause 4.2 shall apply') + run('.'))],
      [para(run('See ') + nested('Clause 4.2 may apply') + run('.'))],
    );
    expect(r.summary.total).toBe(1);
    expectRoundTrip(r);
  });

  it('w:ids open in document order, text boxes included', async () => {
    const r = await compare(
      [para(run('The Seller shall pay.') + textBox('Payment is due monthly'))],
      [para(run('The Seller shall pay.') + textBox('Payment is due quarterly'))],
    );
    const ids = [...firstPara(r).matchAll(/<w:(?:ins|del) w:id="(\d+)"/g)].map((m) => Number(m[1]));
    expect(ids).toEqual([1, 2, 3, 4]);
  });

  it('an inserted or deleted paragraph with a simple field or text box revises inside them', async () => {
    const keep = para(run('The Buyer shall pay.'));
    const extra = para(run('Page ') + '<w:fldSimple w:instr=" PAGE "><w:r><w:t>2</w:t></w:r></w:fldSimple>' + textBox('Inner words'));
    for (const r of [await compare([keep], [keep, extra]), await compare([keep, extra], [keep])]) {
      expect(r.authoredBodyXml).not.toMatch(/<w:(ins|del) [^>]*><w:fldSimple/);
      expect(r.authoredBodyXml).toMatch(/<w:fldSimple [^>]*><w:(ins|del) /);
      expect(r.authoredBodyXml).toMatch(/<w:txbxContent><w:p><w:(ins|del) /);
      expectRoundTrip(r);
    }
  });
});

// ---------------------------------------------------------------------------
// Review round two: inputs that tripped the
// self-check and sent the paragraph back to the lossy rebuild
// ---------------------------------------------------------------------------

describe('ADR-0045: text-equal objects that differ inside, and edges', () => {
  const linkWithNote = (t: string) => `<w:hyperlink r:id="rId9">${run(t)}${footnoteRef()}</w:hyperlink>`;

  it('a footnote ADDED inside a link whose words are unchanged is reported, link kept', async () => {
    const r = await compare(
      [para(run('The Seller shall use ') + hyperlink('the portal'))],
      [para(run('The Buyer shall use ') + linkWithNote('the portal'))],
    );
    const accepted = flattenTrackedChanges(r.authoredBodyXml, 'acceptAll');
    expect(census(accepted, false).footnote).toBe(1);
    expect(census(accepted, false).hyperlink).toBe(1);
    expectRoundTrip(r);
  });

  it('a footnote REMOVED inside a link is shown struck, and a later equation survives', async () => {
    const r = await compare(
      [para(run('The Seller shall use ') + linkWithNote('the portal') + run(' where ') + math('k=1'))],
      [para(run('The Buyer shall use ') + hyperlink('the portal') + run(' where ') + math('k=1'))],
    );
    expect(census(flattenTrackedChanges(r.authoredBodyXml, 'acceptAll'), false).math).toBe(1);
    expectRoundTrip(r); // footnote 1 on reject, 0 on accept
  });

  it('a link opening with a real space right after an equation keeps both', async () => {
    const r = await compare(
      [para(run('Let ') + math('x') + hyperlink(' the Seller portal'))],
      [para(run('Let ') + math('x') + hyperlink(' the Buyer portal'))],
    );
    const p = firstPara(r);
    expect(count(p, '<m:oMath>')).toBe(1);
    expect(count(p, '<w:hyperlink ')).toBe(1);
    expectRoundTrip(r);
  });

  it('a footnote or bookmark just before a link stays outside it when the diff opens with a deletion', async () => {
    const r = await compare(
      [para(run('Pay ') + footnoteRef() + '<w:bookmarkStart w:id="1" w:name="_Ref10002"/><w:bookmarkEnd w:id="1"/>' + hyperlink('Seller portal'))],
      [para(run('Pay ') + footnoteRef() + '<w:bookmarkStart w:id="1" w:name="_Ref10002"/><w:bookmarkEnd w:id="1"/>' + hyperlink('portal'))],
    );
    const link = /<w:hyperlink [^>]*>(.*?)<\/w:hyperlink>/.exec(firstPara(r))![1]!;
    expect(link).not.toContain('footnoteReference');
    expect(link).not.toContain('bookmark');
    expectRoundTrip(r);
  });

  it('a space moving across a link edge leaves the unchanged link unmarked', async () => {
    const r = await compare(
      [para(hyperlink('the portal ') + run('now, Seller.'))],
      [para(hyperlink('the portal') + run(' now, Buyer.'))],
    );
    const p = firstPara(r);
    expect(count(p, '<w:hyperlink ')).toBe(1);
    expect(p).not.toMatch(/<w:hyperlink [^>]*><w:(ins|del)/);
    expectRoundTrip(r);
  });

  // Found by the seeded sweep (pipeline.inline-objects.sweep.test.ts), 10,000
  // pairs: each of these sent its paragraph back to the lossy rebuild, which
  // drops every field in it.
  it('an equation emptied by a revision reads as nothing, not as a word break', async () => {
    // Rejecting leaves the revised equation's empty shell between "notice" and
    // "days", two words of the original that were never apart.
    const r = await compare(
      [para(refField('pay the notice') + refField('days Buyer') + refField('Price') + math('ab'))],
      [para(refField('pay the notice') + refField('days Buyer') + refField('shall') + math('cd'))],
    );
    expect(count(firstPara(r), 'w:fldCharType="begin"')).toBe(5); // one kept, two each inserted and struck — none rebuilt
    expectRoundTrip(r);
  });

  // a quoted defined term at a link's end makes the word diff
  // isolate a space added between two adjacent objects. The space collapses
  // into the first object's; the second, unchanged, carried its own copy into
  // both views when emitted verbatim.
  it('a space added between two adjacent links, after a quoted term, keeps both links', async () => {
    const r = await compare(
      [para(run('The Seller must use ') + hyperlink('the "Portal"') + hyperlink('Terms') + run(' now.'))],
      [para(run('The Buyer must use ') + hyperlink('the "Portal" ') + hyperlink(' Terms') + run(' now.'))],
    );
    expect(count(flattenTrackedChanges(firstPara(r), 'acceptAll'), '<w:hyperlink ')).toBe(2);
    expect(count(flattenTrackedChanges(firstPara(r), 'rejectAll'), '<w:hyperlink ')).toBe(2);
    expectRoundTrip(r);
  });

  it('… and between two adjacent REF fields', async () => {
    const r = await compare(
      [para(run('The Seller: see ') + refField('Clause "A"') + refField('Schedule') + run(' now.'))],
      [para(run('The Buyer: see ') + refField('Clause "A" ') + refField(' Schedule') + run(' now.'))],
    );
    expect(count(firstPara(r), 'w:fldCharType="begin"')).toBeGreaterThanOrEqual(2);
    expectRoundTrip(r);
  });

  it('a space the diff lines up across two different links is split, not shown twice', async () => {
    // "the " is equal text, but its space is inside the first link in the
    // original and opens the second in the revised version.
    const r = await compare(
      [para(run('The Seller: ') + hyperlink('Price portal (the "Portal")') + hyperlink('days within shall') + run('now'))],
      [para(run('The Buyer: ') + hyperlink('Price portal the') + hyperlink(' days within shall') + run('now'))],
    );
    const accepted = flattenTrackedChanges(firstPara(r), 'acceptAll');
    expect(census(accepted, false).hyperlink).toBe(2);
    expectRoundTrip(r);
  });

  it('an unchanged control’s edge space lined up across a deletion stays where the original had it', async () => {
    // The revised control's trailing space lines up with the original's space
    // AFTER two deleted words; kept verbatim, the control carried it in front of them.
    const sdt = (t: string) => `<w:sdt><w:sdtPr><w:alias w:val="Field"/></w:sdtPr><w:sdtContent>${run(t)}</w:sdtContent></w:sdt>`;
    const r = await compare(
      [para(sdt('days. (the "Portal")') + run('"Goods" Seller ') + run('Seller\'s shall notice shall ') + hyperlink('"Goods"'))],
      [para(sdt('days. (the "Portal") ') + run('Seller\'s shall Price shall ') + hyperlink('"Goods"'))],
    );
    expect(count(firstPara(r), '<w:sdtContent>')).toBeGreaterThanOrEqual(1);
    expectRoundTrip(r);
  });

  it('a link’s space trimmed at the paragraph start on one side is not carried mid-sentence', async () => {
    // The revised link opens the paragraph with a space the extractor trims;
    // kept verbatim after the original's deleted words, it would read as one.
    const picture = '<w:r><w:drawing><wp:inline><a:graphic><a:graphicData/></a:graphic></wp:inline></w:drawing></w:r>';
    const r = await compare(
      [para(run(' S$100') + picture + hyperlink('Clause '))],
      [para(hyperlink(' Clause '))],
    );
    expect(census(flattenTrackedChanges(firstPara(r), 'acceptAll'), false).hyperlink).toBe(1);
    expectRoundTrip(r);
  });

  it('a control’s new space, lined up with a text box’s boundary across a deletion, is struck from the original’s view', async () => {
    // The revised control gained a trailing space; the original reads a space
    // there only as the boundary of a box that a deleted control precedes.
    const sdt = (t: string) => `<w:sdt><w:sdtPr><w:alias w:val="Field"/></w:sdtPr><w:sdtContent>${run(t)}</w:sdtContent></w:sdt>`;
    const r = await compare(
      [para(run('The Seller sells ') + sdt('and;') + sdt('the Clause') + textBox('Price days.') + run(' now'))],
      [para(run('The Buyer sells ') + sdt('and; ') + textBox('Price days.') + run(' now'))],
    );
    expect(count(firstPara(r), '<w:txbxContent>')).toBe(1);
    expectRoundTrip(r);
  });

  it('a link of only a space, moved, is marked where it was and where it is', async () => {
    // Its space collapses or is trimmed away, so it has no character of its
    // own to line up by; taken from the revised side, it carried the space
    // to a place the original never had one.
    const r = await compare(
      [para(run('Buyer (the ') + hyperlink(' ') + run('Seller’s'))],
      [para(hyperlink(' ') + run('the'))],
    );
    expect(census(flattenTrackedChanges(firstPara(r), 'acceptAll'), false).hyperlink).toBe(1);
    expectRoundTrip(r);
  });

  it('three links, one of them only a space, merged into one keep a link', async () => {
    const r = await compare(
      [para(run('The Seller must use ') + hyperlink('the Portal') + hyperlink(' ') + hyperlink('Terms') + run(' now.'))],
      [para(run('The Buyer must use ') + hyperlink('the Portal Terms') + run(' now.'))],
    );
    expect(census(flattenTrackedChanges(firstPara(r), 'acceptAll'), false).hyperlink).toBe(1); // emptied shells aside
    expectRoundTrip(r);
  });

  it('a word deleted beside an unchanged equation leaves the equation unmarked', async () => {
    // The space between "Price" and the formula is only the equation's boundary
    // in the original; credited to the equation, deleting the word marked the
    // unchanged equation struck and re-inserted.
    for (const [original, revised] of [
      [run('Pay Price') + math('k=1') + run('now'), run('Pay') + math('k=1') + run('now')],
      [run('Pay the') + math('k=1') + run('Price'), run('Pay') + math('k=1')],
    ]) {
      const r = await compare([para(original!)], [para(revised!)]);
      const p = firstPara(r);
      expect(count(p, '<m:oMath>')).toBe(1);
      expect(p).not.toMatch(/<m:oMath>(?:(?!<\/m:oMath>).)*<w:(ins|del) /);
      expectRoundTrip(r);
    }
  });

  it('an unchanged content control’s edge space, gone in the revised side, is struck beside it', async () => {
    // The original's real space after "Buyer" lines up with nothing but the
    // text box's boundary on the revised side. The control is kept verbatim
    // from the revised side, which has no such space: it must be struck, or
    // rejecting fuses "Buyer" with the deleted "within".
    const sdt = (t: string) => `<w:sdt><w:sdtPr><w:alias w:val="Field"/></w:sdtPr><w:sdtContent>${run(t)}</w:sdtContent></w:sdt>`;
    const r = await compare(
      [para(sdt('within Seller Buyer ') + bookmarked(run('within ')) + textBox('Seller Seller Completion within '))],
      [para(sdt('within Seller Buyer') + textBox('Seller Seller Completion within '))],
    );
    const p = firstPara(r);
    expect(count(p, '<w:txbxContent>')).toBe(1);
    expect(p).not.toMatch(/<w:txbxContent>(?:(?!<\/w:txbxContent>).)*<w:(ins|del) /);
    expectRoundTrip(r);
  });

  it('… and so is one whose revised space moved into an edited link', async () => {
    const sdt = (t: string) => `<w:sdt><w:sdtPr><w:alias w:val="Field"/></w:sdtPr><w:sdtContent>${run(t)}</w:sdtContent></w:sdt>`;
    const picture = '<w:r><w:drawing><wp:inline><a:graphic><a:graphicData/></a:graphic></wp:inline></w:drawing></w:r>';
    const r = await compare(
      [para(sdt('ten ') + hyperlink('Buyer Buyer may ') + picture + hyperlink(' ten notice may'))],
      [para(sdt('ten') + picture + hyperlink(' ten notice Buyer'))],
    );
    expect(count(firstPara(r), '<w:sdtContent>')).toBe(1);
    expectRoundTrip(r);
  });

  it('a content control replaced whole hands its shared space to the next, which is never shelled twice', async () => {
    // Dropping the first pair from the inside diff re-homes a space the two
    // controls shared; the second pair, already checked, then opened its shell
    // twice. The check runs again until nothing more breaks, which here
    // replaces both pairs whole.
    const sdt = (t: string) => `<w:sdt><w:sdtPr><w:alias w:val="Field"/></w:sdtPr><w:sdtContent>${run(t)}</w:sdtContent></w:sdt>`;
    const r = await compare(
      [para(run('Price ten may ten ') + run('pay') + sdt('Price Completion Completion') + sdt(' ten within days '))],
      [para(run('Price ten may of ') + run('portal') + sdt('Price Completion Completion') + sdt('ten within days '))],
    );
    const p = firstPara(r);
    expect(count(p, '<w:sdtContent>')).toBe(4); // each control inserted once and struck once
    expect(p).not.toContain('<w:sdtContent><w:r>');
    expectRoundTrip(r);
  });
});

// ---------------------------------------------------------------------------
// What must not move, and the limits that are recorded rather than fixed
// ---------------------------------------------------------------------------

describe('ADR-0045: what stays exactly as it was', () => {
  // The source rebuild rebuilds every edited paragraph from its own runs; a paragraph of
  // plain, unformatted runs comes out byte for byte as the text rebuild wrote it.
  it('an edited paragraph of plain runs is authored byte for byte as before', async () => {
    const r = await compare(
      [para(run('The Seller shall pay the Price.'))],
      [para(run('The Seller must pay the Price.'))],
    );
    const rev = (id: number) => `w:id="${id}" w:author="${AUTHOR}" w:date="${DATE}"`;
    expect(firstPara(r)).toBe(
      '<w:p><w:r><w:t xml:space="preserve">The Seller </w:t></w:r>' +
        `<w:ins ${rev(1)}><w:r><w:t>must</w:t></w:r></w:ins>` +
        `<w:del ${rev(2)}><w:r><w:delText>shall</w:delText></w:r></w:del>` +
        '<w:r><w:t xml:space="preserve"> pay the Price.</w:t></w:r></w:p>',
    );
  });

  it('falls back to the text rebuild when the word diff does not spell both sides', () => {
    const recs = extractParaRecords(
      `<w:body>${para(run('Pay ') + hyperlink('here'))}${para(run('Pay ') + hyperlink('there'))}</w:body>`,
      'document',
      CONFIG,
    );
    const broken: WordDiff[] = [{ op: 'eq', text: 'Pay nowhere' }];
    let id = 1;
    const ids = { author: AUTHOR, date: DATE, nextId: () => id++ };
    expect(authorObjectAwareRuns(broken, recs[0]!, recs[1]!, ids, () => '', (x) => x)).toBeNull();
    expect(id).toBe(1); // nothing minted
  });

  it('FALLS BACK: a field crossing a paragraph mark is authored by the text rebuild', async () => {
    // The unit model does not cover a field that opens in one paragraph and
    // closes in another; the paragraph gets the pre-ADR rebuild, field and all
    // lost, and is still a correct redline of its text.
    const open = (t: string) =>
      para(run(t) + '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' + run(' Entry one'));
    const close = para(run('Entry two') + '<w:r><w:fldChar w:fldCharType="end"/></w:r>');
    const r = await compare([open('Contents shall follow.'), close], [open('Contents must follow.'), close]);
    expect(firstPara(r)).not.toContain('w:fldChar');
    expect(r.summary.total).toBe(1);
  });

  it('FALLS BACK: an object inside the stripped numbering prefix', async () => {
    // "4.2 " is stripped into `numberingPrefix` and restored as a plain run;
    // a field producing it would be emitted twice, so the text rebuild is used.
    const r = await compare(
      [para(refField('4.2') + run(' The Seller shall pay.'))],
      [para(refField('4.2') + run(' The Seller must pay.'))],
    );
    expect(r.revisedParas[0]!.numberingPrefix).toBe('4.2 ');
    expect(firstPara(r)).not.toContain('w:fldChar');
  });

  it('an equation straight after a manual number is not inside it', async () => {
    // The space the number is stripped with is only the equation's boundary.
    const r = await compare(
      [para(run('4.2') + math('k=1') + run(' The Seller shall pay.'))],
      [para(run('4.2') + math('k=1') + run(' The Buyer shall pay.'))],
    );
    expect(r.revisedParas[0]!.numberingPrefix).toBe('4.2 ');
    const p = firstPara(r);
    expect(count(p, '<m:oMath>')).toBe(1);
    expect(p).not.toMatch(/<m:oMath>(?:(?!<\/m:oMath>).)*<w:(ins|del) /);
    expectRoundTrip(r);
  });

  it('a footnote reference that moved within an edited paragraph is struck where it was and inserted where it is', async () => {
    // Zero-width objects have no text to line up by, so one revised reference
    // stands for both sides only where a single point is its place in both.
    // A moved one has two places: kept unmarked, rejecting every change left it
    // where it now is (limit 9(a) until a later fix).
    const r = await compare(
      [para(run('The Seller shall pay') + footnoteRef() + run(' the Price.'))],
      [para(run('The Seller must pay the Price.') + footnoteRef())],
    );
    expect(r.summary.total).toBe(1);
    const p = firstPara(r);
    expect(count(p, 'w:footnoteReference')).toBe(2);
    const rejected = flattenTrackedChanges(p, 'rejectAll');
    expect(rejected.indexOf('footnoteReference')).toBeLessThan(rejected.indexOf('the Price'));
    const accepted = flattenTrackedChanges(p, 'acceptAll');
    expect(accepted.indexOf('footnoteReference')).toBeGreaterThan(accepted.indexOf('the Price'));
    expectRoundTrip(r);
  });

  it('a moved footnote struck beside a kept picture keeps the original’s order between them', async () => {
    // Both sit at the paragraph's end in the original, footnote first; the
    // picture is kept and the footnote struck there — in that order.
    const picture = '<w:r><w:drawing><wp:inline><a:graphic><a:graphicData/></a:graphic></wp:inline></w:drawing></w:r>';
    const r = await compare(
      [para(math('k=1') + run('Buyer Price days.') + run('pay the Seller') + footnoteRef() + picture)],
      [para(math('k=1') + footnoteRef() + run('pay the Seller') + picture)],
    );
    const rejected = flattenTrackedChanges(firstPara(r), 'rejectAll');
    expect(rejected.indexOf('footnoteReference')).toBeLessThan(rejected.indexOf('<w:drawing>'));
    expect(rejected.indexOf('footnoteReference')).toBeGreaterThan(rejected.indexOf('pay the Seller'));
    expectRoundTrip(r);
  });

  it('a footnote whose neighbouring words were replaced on both sides keeps its place in both views', async () => {
    // Its original place falls inside the struck words, its revised place
    // inside the inserted ones: no single point is right for both.
    for (const [original, revised] of [
      [run('The Seller') + footnoteRef() + run(' shall pay the Price.'), run('The Buyer') + footnoteRef() + run(' must pay the Price.')],
      [run('payable within ten Business Days') + footnoteRef() + run(' of Completion.'), run('payable within five days') + footnoteRef() + run(' after Completion.')],
    ]) {
      const r = await compare([para(original!)], [para(revised!)]);
      const p = firstPara(r);
      const before = (xml: string, word: string) => xml.indexOf('footnoteReference') > xml.indexOf(word);
      expect(before(flattenTrackedChanges(p, 'rejectAll'), original!.includes('Seller') ? 'Seller' : 'Business Days')).toBe(true);
      expect(before(flattenTrackedChanges(p, 'acceptAll'), original!.includes('Seller') ? 'Buyer' : 'five days')).toBe(true);
      expectRoundTrip(r);
    }
  });

});

// ---------------------------------------------------------------------------
// 5. A change to a zero-text object alone (ADR-0045 § Amendment — limit 1)
// ---------------------------------------------------------------------------

/**
 * A footnote reference, an endnote reference, a picture: no text, so a
 * paragraph whose ONLY change is one of them added, removed or moved pairs as
 * `equal`, and before that fix it produced no op at all (`summary.total` 0) —
 * suppression by omission (ADR-0008). Requirement: the removal of footnotes or pictures must be reported.
 */
describe('ADR-0045 amendment: a footnote or picture added, removed or moved alone is reported', () => {
  const SENTENCE = 'The Seller shall pay the Price.';
  const only = (r: Result) => r.ops.filter((o) => o.type !== 'equal');

  it('a removed footnote reference is reported, struck, and both views read true', async () => {
    const r = await compare([para(run(SENTENCE) + footnoteRef())], [para(run(SENTENCE))]);
    expect(r.summary.total).toBe(1);
    expect(r.ops.map((o) => o.type)).toEqual(['property', 'equal']);
    const op = only(r)[0]!;
    expect(op.propertyChange?.objects).toBe(true);
    const p = firstPara(r);
    expect(p).toMatch(/<w:del [^>]*><w:r><w:rPr><w:rStyle w:val="FootnoteReference"\/><\/w:rPr><w:footnoteReference w:id="1"\/><\/w:r><\/w:del>/);
    // The words are kept, not struck and re-inserted.
    expect(p).not.toContain('w:delText');
    expect(count(p, '<w:ins ')).toBe(0);
    expectRoundTrip(r);
  });

  it('an added footnote reference is reported and inserted', async () => {
    const r = await compare([para(run(SENTENCE))], [para(run(SENTENCE) + footnoteRef())]);
    expect(r.summary.total).toBe(1);
    const p = firstPara(r);
    expect(p).toMatch(/<w:ins [^>]*><w:r><w:rPr><w:rStyle w:val="FootnoteReference"\/><\/w:rPr><w:footnoteReference/);
    expect(count(p, '<w:del ')).toBe(0);
    expectRoundTrip(r);
  });

  it('an endnote reference, and a picture, removed or added alone are reported', async () => {
    for (const object of [endnoteRef(), inlinePicture()]) {
      for (const [original, revised] of [
        [para(run('The Seller ') + object + run('shall pay the Price.')), para(run('The Seller shall pay the Price.'))],
        [para(run('The Seller shall pay the Price.')), para(run('The Seller ') + object + run('shall pay the Price.'))],
      ]) {
        const r = await compare([original!], [revised!]);
        expect(r.summary.total).toBe(1);
        expect(firstPara(r)).toMatch(/<w:(ins|del) [^>]*><w:r>(<w:rPr>.*?<\/w:rPr>)?<w:(endnoteReference|drawing)/);
        expectRoundTrip(r);
      }
    }
  });

  it('a paragraph holding only a picture, emptied, is reported', async () => {
    const r = await compare([para(inlinePicture()), para(run(SENTENCE))], [para(''), para(run(SENTENCE))]);
    expect(r.summary.total).toBe(1);
    expect(firstPara(r)).toMatch(/<w:del [^>]*><w:r><w:drawing>/);
    expectRoundTrip(r);
  });

  it('a footnote reference moved within unchanged words is reported: inserted where it is, struck where it was', async () => {
    const r = await compare(
      [para(run('The Seller') + footnoteRef() + run(' shall pay the Price.'))],
      [para(run('The Seller shall pay the Price.') + footnoteRef())],
    );
    expect(r.summary.total).toBe(1);
    const p = firstPara(r);
    expect(count(p, '<w:footnoteReference')).toBe(2);
    const rejected = flattenTrackedChanges(p, 'rejectAll');
    expect(rejected.indexOf('footnoteReference')).toBeLessThan(rejected.indexOf(' shall pay'));
    const accepted = flattenTrackedChanges(p, 'acceptAll');
    expect(accepted.indexOf('footnoteReference')).toBeGreaterThan(accepted.indexOf('the Price.'));
    expectRoundTrip(r);
  });

  it('one footnote of three removed strikes that one and keeps the other two unmarked', async () => {
    // Before the alignment, any difference in the sequence of zero-width
    // objects marked EVERY one of them: three struck and two re-inserted.
    const r = await compare(
      [para(run('The Seller') + footnoteRef('1') + run(' shall pay') + footnoteRef('2') + run(' the Price') + footnoteRef('3') + run('.'))],
      [para(run('The Seller') + footnoteRef('1') + run(' shall pay the Price') + footnoteRef('3') + run('.'))],
    );
    expect(r.summary.total).toBe(1);
    const p = firstPara(r);
    expect(count(p, '<w:footnoteReference')).toBe(3);
    expect(count(p, '<w:del ')).toBe(1);
    expect(count(p, '<w:ins ')).toBe(0);
    expect(p).toMatch(/<w:del [^>]*><w:r>(<w:rPr>.*?<\/w:rPr>)?<w:footnoteReference w:id="2"\/>/);
    expectRoundTrip(r);
  });

  it('the words keep their own run formatting: only the object is marked', async () => {
    // The replace author rebuilds every word from the diff (ADR-0045 § Accepted
    // limits 5); a clause whose words did not change must not lose its bold.
    const bold = '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Completion Date</w:t></w:r>';
    const r = await compare(
      [para(run('On the ') + bold + footnoteRef() + run(' the Seller shall pay.'))],
      [para(run('On the ') + bold + run(' the Seller shall pay.'))],
    );
    const p = firstPara(r);
    expect(p).toContain(bold);
    expect(count(p, '<w:del ')).toBe(1);
    expectRoundTrip(r);
  });

  it('a struck object after a kept one at the same place stays after it, and the words keep their formatting', async () => {
    // Picture then footnote, both after "Pay."; the footnote goes. Struck before
    // the kept picture, rejecting would swap them — the self-check would catch
    // that and send the pair to the replace author, which rebuilds the words
    // (and drops their bold), so the bold is what shows which path ran.
    const bold = '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Pay.</w:t></w:r>';
    const r = await compare(
      [para(bold + inlinePicture() + footnoteRef()), para(run(SENTENCE))],
      [para(bold + inlinePicture()), para(run(SENTENCE))],
    );
    const p = firstPara(r);
    expect(p).toContain(bold);
    const rejected = flattenTrackedChanges(p, 'rejectAll');
    expect(rejected.indexOf('<w:drawing>')).toBeLessThan(rejected.indexOf('footnoteReference'));
    expectRoundTrip(r);
  });

  it('a footnote removed from inside a link is reported; the link is marked whole', async () => {
    const link = (inner: string) =>
      `<w:hyperlink r:id="rId9"><w:r><w:t xml:space="preserve">the Seller</w:t></w:r>${inner}</w:hyperlink>`;
    const r = await compare(
      [para(run('Pay ') + link(footnoteRef()) + run(' on Completion.'))],
      [para(run('Pay ') + link('') + run(' on Completion.'))],
    );
    expect(r.summary.total).toBe(1);
    expect(firstPara(r)).toMatch(/<w:del [^>]*>(?:(?!<\/w:del>).)*<w:footnoteReference/);
    expectRoundTrip(r);
  });

  it('a footnote swapped for an endnote in the same place reads insertion first (ADR-0043)', async () => {
    const r = await compare([para(run(SENTENCE) + footnoteRef())], [para(run(SENTENCE) + endnoteRef())]);
    expect(r.summary.total).toBe(1);
    const p = firstPara(r);
    expect(p.indexOf('<w:endnoteReference')).toBeLessThan(p.indexOf('<w:footnoteReference'));
    expect(p).toMatch(/<w:ins [^>]*><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:endnoteReference/);
    expectRoundTrip(r);
  });

  it('…and still does just after a bookmark (found by the sweep)', async () => {
    // A struck object at a bookmark's offset used to be emitted before the
    // bookmark, and the inserted one after it: deletion first.
    const r = await compare(
      [para(bookmarked(run('Pay.')) + footnoteRef())],
      [para(bookmarked(run('Pay.')) + inlinePicture())],
    );
    const p = firstPara(r);
    expect(p.indexOf('<w:drawing')).toBeLessThan(p.indexOf('<w:footnoteReference'));
    expect(p.indexOf('<w:bookmarkEnd')).toBeLessThan(p.indexOf('<w:drawing'));
    expectRoundTrip(r);
  });

  it('a picture removed beside a text box is struck at the box’s edge, not inside it', async () => {
    const r = await compare(
      [para(run('See the plan ') + inlinePicture() + textBox(' Plan A ') + run(' attached.'))],
      [para(run('See the plan ') + textBox(' Plan A ') + run(' attached.'))],
    );
    expect(r.summary.total).toBe(1);
    const p = firstPara(r);
    expect(p.indexOf('<w:drawing><wp:inline>')).toBeLessThan(p.indexOf('<w:txbxContent>'));
    expect(count(p, '<w:txbxContent>')).toBe(1); // the box itself is not replaced
    expectRoundTrip(r);
  });

  it('a moved footnote reads as struck and inserted in the change-list row', async () => {
    const r = await compare(
      [para(run('The Seller') + footnoteRef() + run(' shall pay the Price.'))],
      [para(run('The Seller shall pay the Price.') + footnoteRef())],
    );
    expect(previewText(only(r)[0]!)).toBe('-[footnote reference] +[footnote reference]');
  });

  it('an embedded object removed from inside a link is marked, not left in the revised paragraph unmarked', async () => {
    // the replace author's census did not count `w:object`, so it
    // kept the two links as an unchanged pair and the redline carried no mark.
    const link = (inner: string) =>
      `<w:hyperlink r:id="rId9"><w:r><w:t xml:space="preserve">the Seller</w:t></w:r>${inner}</w:hyperlink>`;
    const r = await compare(
      [para(run('Pay ') + link(embeddedObject()) + run(' now.'))],
      [para(run('Pay ') + link('') + run(' now.'))],
    );
    expect(r.summary.total).toBe(1);
    const p = firstPara(r);
    expect(count(flattenTrackedChanges(p, 'rejectAll'), '<w:object')).toBe(1);
    expect(count(flattenTrackedChanges(p, 'acceptAll'), '<w:object')).toBe(0);
    // The link is marked whole; the words around it are not touched.
    expect(p).not.toMatch(/<w:delText[^>]*>[^<]*(Pay|now)/);
    expectRoundTrip(r);
  });

  it('an edited clause that also lost an embedded object from inside a link keeps it in the rejected view', async () => {
    // A `replace` op: the census decides whether the link is an unchanged pair,
    // and before `w:object` was counted it was — the object vanished from both views.
    const link = (inner: string) =>
      `<w:hyperlink r:id="rId9"><w:r><w:t xml:space="preserve">the Seller</w:t></w:r>${inner}</w:hyperlink>`;
    const r = await compare(
      [para(run('Pay ') + link(embeddedObject()) + run(' within ten days.'))],
      [para(run('Pay ') + link('') + run(' within five days.'))],
    );
    expect(r.ops[0]!.type).toBe('replace');
    expect(count(flattenTrackedChanges(firstPara(r), 'rejectAll'), '<w:object')).toBe(1);
    expectRoundTrip(r);
  });

  it('a footnote moved INSIDE a link is marked: a census cannot see a move', async () => {
    // The splice declines (the footnote is not a top-level object); the replace
    // author keeps the link verbatim because its text and census match, so the
    // places check rejects that second try and the paragraph is marked by the
    // third — whole, over-marked but never hidden.
    const link = (inner: string) => `<w:hyperlink r:id="rId9">${inner}</w:hyperlink>`;
    const t = (x: string) => `<w:r><w:t xml:space="preserve">${x}</w:t></w:r>`;
    const r = await compare(
      [para(run('Pay ') + link(t('the') + footnoteRef() + t(' Seller')) + run(' now.'))],
      [para(run('Pay ') + link(t('the Seller') + footnoteRef()) + run(' now.'))],
    );
    expect(r.summary.total).toBe(1);
    const p = firstPara(r);
    const rejected = flattenTrackedChanges(p, 'rejectAll');
    expect(rejected.indexOf('footnoteReference')).toBeLessThan(rejected.indexOf(' Seller'));
    const accepted = flattenTrackedChanges(p, 'acceptAll');
    expect(accepted.indexOf('footnoteReference')).toBeGreaterThan(accepted.indexOf('the Seller'));
    expectRoundTrip(r);
  });

  it('a footnote moved by one character at the start of a text box, or into the box from just before it, is a move', async () => {
    // with a space before the box, its leading boundary collapsed
    // into that space, and both pairs below gave identical places.
    const box = (inner: string) =>
      '<w:r><w:drawing><wp:anchor><a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent>' +
      `<w:p>${inner}</w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
    for (const [o, v] of [
      [run('See ') + box(footnoteRef() + run('Plan A')), run('See ') + box(run('P') + footnoteRef() + run('lan A'))],
      [run('See ') + footnoteRef() + box(run('Plan A')), run('See ') + box(footnoteRef() + run('Plan A'))],
    ]) {
      const r = await compare([para(o! + run(' attached.'))], [para(v! + run(' attached.'))]);
      expect(r.ops.some((op) => op.propertyChange?.objects === true), o).toBe(true);
      expectRoundTrip(r);
    }
  });

  it('a footnote moved from just before an equation to just inside it is a move', async () => {
    // The space before the equation collapses with its leading boundary, so the
    // two places have the same count of characters; only `within` tells them
    // apart. A math run may carry a footnote reference (EG_RunInnerContent).
    const r = await compare(
      [para(run('Rate x ') + footnoteRef() + math('r=2%') + run(' a year.'))],
      [para(run('Rate x ') + '<m:oMath><m:r><w:footnoteReference w:id="1"/></m:r><m:r><m:t>r=2%</m:t></m:r></m:oMath>' + run(' a year.'))],
    );
    expect(r.summary.total).toBe(1);
    expect(r.ops.some((op) => op.propertyChange?.objects === true)).toBe(true);
    expectRoundTrip(r);
  });

  it('a footnote or picture moved across a manual clause number is a move', async () => {
    // places were clamped to the first character after the
    // number, so "before the number" and "after it" read as one place.
    for (const object of [footnoteRef(), inlinePicture()]) {
      const r = await compare(
        [para(object + run('1. The Seller shall pay.'))],
        [para(run('1. ') + object + run('The Seller shall pay.'))],
      );
      expect(r.summary.total).toBe(1);
      const p = firstPara(r);
      const tag = object.includes('footnote') ? 'footnoteReference' : '<w:drawing';
      expect(p).not.toContain('w:delText'); // spliced: the words are kept, not struck and re-inserted
      expect(flattenTrackedChanges(p, 'rejectAll').indexOf(tag)).toBeLessThan(flattenTrackedChanges(p, 'rejectAll').indexOf('1. '));
      expect(flattenTrackedChanges(p, 'acceptAll').indexOf(tag)).toBeGreaterThan(flattenTrackedChanges(p, 'acceptAll').indexOf('1. '));
      expectRoundTrip(r);
    }
  });

  it('a renumbered clause keeps its footnote where it was: no change, before the number or after it', async () => {
    for (const [o, v] of [
      [footnoteRef() + run('9. The Seller shall pay.'), footnoteRef() + run('10. The Seller shall pay.')],
      [run('9. ') + footnoteRef() + run('The Seller shall pay.'), run('10. ') + footnoteRef() + run('The Seller shall pay.')],
    ]) {
      const r = await compare([para(o!)], [para(v!)]);
      expect(r.ops.some((op) => op.propertyChange?.objects === true)).toBe(false);
    }
  });

  it('a picture-only paragraph emptied to a self-closed <w:p/> stays ONE paragraph', async () => {
    // `<w:p/>` has no open and close tags, so every try failed and
    // the fallback wrapped a whole paragraph in <w:ins> at body level.
    const r = await compare(
      [para(inlinePicture()), para(run(SENTENCE))],
      ['<w:p w:rsidR="00A1B2C3" w:rsidRDefault="00A1B2C3"/>', para(run(SENTENCE))],
    );
    expect(r.summary.total).toBe(1);
    expect(r.authoredBodyXml).not.toMatch(/<w:(ins|del) [^>]*><w:p[\s>/]/);
    const paragraphs = (xml: string) => (xml.match(/<w:p[\s>/]/g) ?? []).length;
    expect(paragraphs(r.authoredBodyXml)).toBe(3);
    expect(paragraphs(flattenTrackedChanges(r.authoredBodyXml, 'acceptAll'))).toBe(3);
    expect(firstPara(r)).toMatch(/<w:del [^>]*><w:r><w:drawing>/);
    expectRoundTrip(r);
  });

  it('RECORDS ITS LIMIT: a symbol character (w:sym) added alone is not reported', async () => {
    // A Wingdings tick box is a run with no text either, and not one of the
    // sorts compared (ADR-0045 § Amendment 1, limit 2).
    const tick = '<w:r><w:sym w:font="Wingdings" w:char="F0FC"/></w:r>';
    const r = await compare([para(run(SENTENCE))], [para(tick + run(SENTENCE))]);
    expect(r.summary.total).toBe(0);
  });

  it('is classified substantive: neither display filter hides it', async () => {
    const r = await compare([para(run(SENTENCE) + footnoteRef())], [para(run(SENTENCE))]);
    const op = only(r)[0]!;
    expect(op.classification).toBe('substantiveText');
    expect(op.classes).not.toContain('crossRefRenumber');
    expect(op.propertyChange?.cascade).toBeUndefined();
    expect(r.summary.byCategory.substantiveText).toBe(1);
    expect(applyFilters(r.ops, { hideFormattingOnly: true, collapseRenumbering: true })).toContainEqual(op);
  });

  it('with a style change too, both are carried and substantive leads', async () => {
    const styled = (inner: string) => `<w:p><w:pPr><w:pStyle w:val="BodyAlt"/></w:pPr>${inner}</w:p>`;
    const r = await compare([para(run(SENTENCE) + footnoteRef())], [styled(run(SENTENCE))]);
    const op = only(r)[0]!;
    expect(op.propertyChange).toMatchObject({ styleId: true, objects: true });
    expect(op.classes).toEqual(['substantiveText', 'formatting']);
    expectRoundTrip(r);
  });

  it('the change-list row and the report name the object', async () => {
    const r = await compare([para(run(SENTENCE) + footnoteRef())], [para(run(SENTENCE))]);
    const op = only(r)[0]!;
    expect(previewText(op)).toBe('-[footnote reference]');
    const report = buildCompareReport({
      mode: 'word', ops: r.ops, summary: r.summary,
      originalParas: r.originalParas, revisedParas: r.revisedParas,
      generatedAt: '2026-10-18', disclosures: [],
    });
    const row = report.rows.find((x) => x.type === 'property')!;
    expect(row.original).toBe(`${SENTENCE}[footnote reference]`);
    expect(row.revised).toBe(SENTENCE);
  });

  it('the report prints the added object on the revised side', async () => {
    const r = await compare([para(run('The Seller ') + run('shall pay.'))], [para(run('The Seller ') + inlinePicture() + run('shall pay.'))]);
    const report = buildCompareReport({
      mode: 'word', ops: r.ops, summary: r.summary,
      originalParas: r.originalParas, revisedParas: r.revisedParas,
      generatedAt: '2026-10-18', disclosures: [],
    });
    const row = report.rows.find((x) => x.type === 'property')!;
    expect(row.original).toBe('The Seller shall pay.');
    expect(row.revised).toBe('The Seller [picture]shall pay.');
  });

  it('a footnote renumbered (its id changed) in the same place is NOT a change', async () => {
    // A footnote inserted earlier in the document renumbers every reference
    // after it; the id is document-local bookkeeping (cf. ADR-0042's numId).
    const r = await compare(
      [para(run(SENTENCE) + footnoteRef('1') + inlinePicture('rId20'))],
      [para(run(SENTENCE) + footnoteRef('2') + inlinePicture('rId31'))],
    );
    expect(r.summary.total).toBe(0);
    expect(r.ops.map((o) => o.type)).toEqual(['equal', 'equal']);
  });

  it('an unchanged paragraph with footnotes and pictures is authored verbatim', async () => {
    const same = para(run('The Seller') + footnoteRef() + run(' shall pay ') + inlinePicture() + run('the Price.') + endnoteRef());
    const r = await compare([same], [same]);
    expect(r.summary.total).toBe(0);
    expect(r.authoredBodyXml).toContain(same);
  });

  it('RECORDS ITS LIMIT: a picture swapped for another in the same place is not reported', async () => {
    // Same sort, same place; which image it shows lives in a media part behind
    // a document-local relationship id (ADR-0045 § Accepted limits 1).
    const r = await compare([para(run(SENTENCE) + inlinePicture('rId20'))], [para(run(SENTENCE) + inlinePicture('rId21'))]);
    expect(r.summary.total).toBe(0);
  });

  it('when the paragraph cannot be cut into units, its runs are struck and re-inserted in ONE paragraph', async () => {
    // A field that opens here and closes in the next paragraph has no unit
    // model (`paragraphUnits` → null). Kept verbatim, the redline would be
    // silent; rebuilt from text, the footnote would be lost.
    const open = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';
    const close = para('<w:r><w:fldChar w:fldCharType="end"/></w:r>' + run('End.'));
    const r = await compare(
      [para(open + run(SENTENCE) + footnoteRef()), close],
      [para(open + run(SENTENCE)), close],
    );
    expect(r.summary.total).toBe(1);
    const xml = r.authoredBodyXml;
    // The footnote reference sits inside the struck runs' deletion.
    expect(xml).toMatch(/<w:del [^>]*>(?:(?!<\/w:del>).)*<w:footnoteReference/);
    expect(xml).toContain('w:delText');
    // Insertion first, and the paragraph mark itself not revised.
    expect(firstPara(r).indexOf('<w:ins ')).toBeLessThan(firstPara(r).indexOf('<w:del '));
    expect(firstPara(r)).not.toMatch(/<w:pPr><w:rPr><w:(ins|del) /);
    expectRoundTrip(r);
  });

  it('…and a section break on that paragraph is authored once, not twice', async () => {
    // a deleted paragraph plus an inserted one kept BOTH marks,
    // because a mark carrying w:sectPr is never revised (ADR-0035).
    const open = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';
    const sect = (inner: string) => `<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr>${inner}</w:p>`;
    const close = para('<w:r><w:fldChar w:fldCharType="end"/></w:r>' + run('End.'));
    const r = await compare(
      [sect(open + run(SENTENCE) + footnoteRef()), close],
      [sect(open + run(SENTENCE)), close],
    );
    expect(r.summary.total).toBe(1);
    expect(count(r.authoredBodyXml, '<w:sectPr')).toBe(1);
    expect(count(flattenTrackedChanges(r.authoredBodyXml, 'acceptAll'), '<w:sectPr')).toBe(1);
    expectRoundTrip(r);
  });
});
