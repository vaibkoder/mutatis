/**
 * Compare — non-text property changes (ADR-0042).
 *
 * The defect: a clause that was outdented has the same WORDS on both sides, so
 * it pairs as `equal` and produces no op. The user accepts the redline and the
 * summary reports nothing — suppression by omission, which ADR-0008 forbids.
 *
 * Everything here runs the real pipeline end to end from OOXML, because the
 * defect lived in the seam between four things that each worked: the extractor
 * never read the style chain, the structural diff never compared what it found,
 * the classifier only labels ops that already exist, and the summary counts ops.
 *
 * The prose is drafting written for this suite. No client material.
 */
import { describe, it, expect } from 'vitest';

import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import { preprocessSide, buildNumberingContext } from '@features/compare/preprocess/index.js';
import { runComparePipeline } from '@features/compare/pipeline.js';
import { authorBody, makeAuthorContext } from '@features/compare/output/index.js';
import { comparePropertyPair } from '@features/compare/diff/index.js';
import { previewText } from '@features/compare/components/ChangeList.js';
import type { DiffOp, ParaRecord } from '@features/compare/model.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const config = getCompareConfig(SG_PROFILE);

/**
 * A styles part where every clause level carries its numbering, and nothing
 * else does — which is how a firm precedent is actually built, and the reason
 * `document.xml` in such a file contains zero `<w:numPr>`.
 */
const STYLES = `<w:styles ${W}>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:pPr/></w:style>
  <w:style w:type="paragraph" w:styleId="Clause1"><w:pPr>
    <w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="Clause2"><w:pPr>
    <w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="Plain"><w:pPr/></w:style>
  <w:style w:type="paragraph" w:styleId="PlainAlt"><w:pPr/></w:style>
</w:styles>`;

const NUMBERING = `<w:numbering ${W}>
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%2)"/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="4"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`;

/** One `<w:p>` with a style and a single run. */
function para(styleId: string, text: string): string {
  return (
    `<w:p><w:pPr><w:pStyle w:val="${styleId}"/></w:pPr>` +
    `<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
  );
}

function body(...paras: string[]): string {
  return `<w:document ${W}><w:body>${paras.join('')}</w:body></w:document>`;
}

/** Records for one side, exactly as `runPipeline` builds them for Word-vs-Word. */
function side(documentXml: string, policy: 'acceptAll' | 'rejectAll'): ParaRecord[] {
  return preprocessSide(
    documentXml,
    'document',
    policy,
    config,
    buildNumberingContext(STYLES, NUMBERING),
  ).paras;
}

async function compare(origXml: string, revXml: string) {
  const original = side(origXml, config.flattenOriginal);
  const revised = side(revXml, config.flattenRevised);
  const result = await runComparePipeline({
    originalParas: original,
    revisedParas: revised,
    config,
    author: 'Mutatis',
    date: '2026-09-12T00:00:00Z',
  });
  return { original, revised, result };
}

const properties = (ops: DiffOp[]): DiffOp[] => ops.filter((o) => o.type === 'property');

// The same three clauses, edited only in the ways each test names.
const C1 = 'The Vendor shall procure that the Company complies with this Clause.';
const C2 = 'Settlement of the balance shall be made in cleared funds within ten Business Days.';
const C3 = 'Tax treatment of each such sum shall be determined by the Reporting Accountant.';

describe('non-text property changes (ADR-0042)', () => {
  it('never reads zero-text objects off a table record — the table path is its own (ADR-0027)', () => {
    const [plain] = side(body('<w:p><w:r><w:t>Cell text.</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>'), 'acceptAll');
    const [bare] = side(body('<w:p><w:r><w:t>Cell text.</w:t></w:r></w:p>'), 'acceptAll');
    expect(comparePropertyPair(plain!, bare!)?.objects).toBe(true);
    const table = { rows: [] } as unknown as NonNullable<ParaRecord['table']>;
    expect(comparePropertyPair({ ...plain!, table }, { ...bare!, table })).toBeNull();
  });

  it('a renumbered clause that ALSO lost a footnote is never collapsed into the cascade (ADR-0045 § Amendment)', async () => {
    // Its number moved because a clause was inserted above — and its footnote
    // went. A cascade member is hidden from the list by *Collapse renumbering*,
    // so this one must not be a member: it changed in its own right.
    const footnoted = (styleId: string, text: string) =>
      `<w:p><w:pPr><w:pStyle w:val="${styleId}"/></w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r>` +
      '<w:r><w:footnoteReference w:id="1"/></w:r></w:p>';
    const C4 = 'Each notice shall be given in writing to the address in Schedule 2.';
    const C5 = 'No variation shall be effective unless made in writing by both parties.';
    const orig = body(para('Clause1', C1), para('Clause2', C2), footnoted('Clause2', C3), para('Clause2', C4), para('Clause2', C5));
    const rev = body(
      para('Clause1', C1),
      para('Clause2', 'The Purchaser shall countersign each such computation.'),
      para('Clause2', C2),
      para('Clause2', C3),
      para('Clause2', C4),
      para('Clause2', C5),
    );

    const { result } = await compare(orig, rev);
    const lost = properties(result.ops).find((o) => o.propertyChange?.objects === true)!;
    expect(lost.propertyChange).toMatchObject({ numberLabel: true, objects: true });
    expect(lost.propertyChange?.cascade).toBeUndefined();
    expect(lost.collapsedInto).toBeUndefined();
    expect(lost.classification).toBe('substantiveText');
    // The clauses whose numbers alone moved still collapse.
    expect(properties(result.ops).filter((o) => o.collapsedInto !== undefined).length).toBeGreaterThanOrEqual(3);
  });

  it('emits EXACTLY ONE op for a clause that was outdented, and classifies it formatting', async () => {
    const orig = body(para('Clause1', C1), para('Clause2', C2));
    const rev = body(para('Clause1', C1), para('Clause1', C2)); // C2 promoted a level

    const { result } = await compare(orig, rev);
    const props = properties(result.ops);

    expect(props).toHaveLength(1);
    expect(props[0]!.classification).toBe('formatting');
    expect(props[0]!.propertyChange).toEqual({
      styleId: true,
      numberingIdentity: true,
      numberLabel: true,
      // The values that moved, carried so the row preview, the navigation
      // anchor and the exported report can say what changed — a property op has
      // no `wordDiffs`, and all three of those surfaces read `wordDiffs`.
      from: { styleId: 'Clause2', numberLabel: '(a)', ilvl: 1 },
      to: { styleId: 'Clause1', numberLabel: '2.', ilvl: 0 },
    });
    // And it reaches the user: the summary counts ops, and before this it counted none.
    expect(result.summary.total).toBe(1);
    expect(result.summary.byCategory.formatting).toBe(1);
  });

  it('records the rendered labels the reader would see change', async () => {
    const orig = body(para('Clause1', C1), para('Clause2', C2));
    const rev = body(para('Clause1', C1), para('Clause1', C2));

    const { original, revised } = await compare(orig, rev);

    expect(original[1]!.numbering?.label).toBe('(a)');
    expect(revised[1]!.numbering?.label).toBe('2.');
  });

  it('classifies a label that moved with the IDENTITY UNCHANGED as a renumbering cascade', async () => {
    // A clause is inserted above C3, so C3's own level and instance are
    // untouched and only its printed letter moves. §9.4 — and §6.1's collapse
    // keys on this class, which is what stops one insertion producing one op per
    // clause that follows it.
    const orig = body(para('Clause1', C1), para('Clause2', C2), para('Clause2', C3));
    const rev = body(
      para('Clause1', C1),
      para('Clause2', C2),
      para('Clause2', 'The Purchaser shall countersign each such computation.'),
      para('Clause2', C3),
    );

    const { original, revised, result } = await compare(orig, rev);
    const props = properties(result.ops);

    expect(props).toHaveLength(1);
    expect(props[0]!.classification).toBe('crossRefRenumber');
    expect(props[0]!.propertyChange).toMatchObject({
      styleId: false,
      numberingIdentity: false,
      numberLabel: true,
      from: { numberLabel: '(b)' },
      to: { numberLabel: '(c)' },
      // The cascade descriptor §6.1's collapse needs — a property op has no
      // `wordDiffs` for `cascade.ts` to read, so it carries its own evidence.
      cascade: { from: '(b)', to: '(c)', shiftedIlvl: 1, shift: 1 },
    });
    expect(original[2]!.numbering?.label).toBe('(b)');
    expect(revised[3]!.numbering?.label).toBe('(c)');
  });

  it('reports a clause re-levelled by the INDENT BUTTON as formatting, not as a cascade', async () => {
    // Same style on both sides; the level was changed by a direct <w:numPr>
    // carrying only an ilvl. This paragraph CHANGED — it is not a bystander
    // whose letter moved — so it must not be collapsed into a renumbering
    // cascade. The only signal separating the two is the identity comparison.
    const orig = body(para('Clause1', C1), para('Clause1', C2));
    const rev = body(
      para('Clause1', C1),
      `<w:p><w:pPr><w:pStyle w:val="Clause1"/>` +
        `<w:numPr><w:ilvl w:val="1"/></w:numPr></w:pPr>` +
        `<w:r><w:t xml:space="preserve">${C2}</w:t></w:r></w:p>`,
    );

    const { result } = await compare(orig, rev);
    const props = properties(result.ops);

    expect(props).toHaveLength(1);
    expect(props[0]!.classification).toBe('formatting');
    expect(props[0]!.propertyChange).toMatchObject({
      styleId: false,
      numberingIdentity: true,
      numberLabel: true,
    });
  });

  it('reports NOTHING when Word merely reassigned the list instance id', async () => {
    // The false positive this comparison must not have. `<w:numId>` is
    // DOCUMENT-LOCAL bookkeeping and Word reassigns it freely — a Save-As onto
    // another template, a list pasted in, a counterparty re-creating the file.
    // Comparing it directly reported SIX changes on this document, which has
    // none: same body, same styles, same abstract definition, different
    // instance ids.
    const styles = (numId: string) =>
      `<w:styles ${W}><w:style w:type="paragraph" w:styleId="C"><w:pPr>` +
      `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr></w:style></w:styles>`;
    const numbering = (numId: string, fmt = 'decimal') =>
      `<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">` +
      `<w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>` +
      `<w:num w:numId="${numId}"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
    const doc = body(para('C', C1), para('C', C2), para('C', C3));

    const read = (numId: string, fmt: string) =>
      preprocessSide(doc, 'document', 'acceptAll', config,
        buildNumberingContext(styles(numId), numbering(numId, fmt))).paras;

    const original = read('3', 'decimal');
    const revised = read('7', 'decimal');
    const result = await runComparePipeline({
      originalParas: original, revisedParas: revised, config,
      author: 'Mutatis', date: '2026-09-12T00:00:00Z',
    });

    expect(original[0]!.numbering?.label).toBe('1.');
    expect(revised[0]!.numbering?.label).toBe('1.');
    expect(properties(result.ops)).toHaveLength(0);
    expect(result.summary.total).toBe(0);
  });

  it('DOES report a redefined numbering format, as formatting and not as a cascade', async () => {
    // The other side of the same coin. Nothing here moved and no sibling was
    // inserted — every clause's own number was REDEFINED, so this is a change to
    // each paragraph's numbering, not a bystander cascade. It is also the case
    // that isolates the identity comparison: with it ablated, these classify as
    // `crossRefRenumber` on the strength of the label alone, which would say a
    // document-wide renumbering happened because a sibling moved.
    const styles = (numId: string) =>
      `<w:styles ${W}><w:style w:type="paragraph" w:styleId="C"><w:pPr>` +
      `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr></w:style></w:styles>`;
    const numbering = (numId: string, fmt: string) =>
      `<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">` +
      `<w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>` +
      `<w:num w:numId="${numId}"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
    const doc = body(para('C', C1), para('C', C2), para('C', C3));

    const read = (fmt: string) =>
      preprocessSide(doc, 'document', 'acceptAll', config,
        buildNumberingContext(styles('3'), numbering('3', fmt))).paras;

    const original = read('decimal');
    const revised = read('lowerRoman');
    const result = await runComparePipeline({
      originalParas: original, revisedParas: revised, config,
      author: 'Mutatis', date: '2026-09-12T00:00:00Z',
    });
    const props = properties(result.ops);

    expect(original[0]!.numbering?.label).toBe('1.');
    expect(revised[0]!.numbering?.label).toBe('i.');
    expect(props).toHaveLength(3);
    expect(props.every((op) => op.classification === 'formatting')).toBe(true);
  });

  it('treats a redefined <w:start> as the list changing, not as a sibling moving', async () => {
    // The same defect as the redefined FORMAT, on a different attribute. The
    // first comparison key covered numFmt/lvlText/isLgl only, so `<w:start>`
    // 1 → 5 left it equal and every clause reported as a bystander cascade
    // ("its number moved because a sibling did") shifted +4 — with no sibling
    // anywhere. The key now covers the whole level definition.
    const styles = `<w:styles ${W}><w:style w:type="paragraph" w:styleId="C"><w:pPr>` +
      `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style></w:styles>`;
    const numbering = (start: number) =>
      `<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">` +
      `<w:start w:val="${start}"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>` +
      `</w:abstractNum><w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
    const doc = body(para('C', C1), para('C', C2), para('C', C3));

    const read = (start: number) =>
      preprocessSide(doc, 'document', 'acceptAll', config,
        buildNumberingContext(styles, numbering(start))).paras;

    const original = read(1);
    const revised = read(5);
    const result = await runComparePipeline({
      originalParas: original, revisedParas: revised, config,
      author: 'Mutatis', date: '2026-09-12T00:00:00Z',
    });
    const props = properties(result.ops);

    expect(original[0]!.numbering?.label).toBe('1.');
    expect(revised[0]!.numbering?.label).toBe('5.');
    expect(props).toHaveLength(3);
    expect(props.every((op) => op.classification === 'formatting')).toBe(true);
    // And so it is NOT collapsed away as a cascade of bystanders.
    expect(result.summary.renumberingCascades).toHaveLength(0);
  });

  it('names the LEVEL when neither side resolves a label, instead of "— → —"', async () => {
    // A `numId` the numbering part does not define: the style chain still
    // resolves a level, so the re-levelling is real and detectable, but no label
    // renders on either side. Without a fallback the row reads "— → —" and the
    // exported row reads "(unnumbered)" twice — the blank row and the
    // identical-sides row, in a narrower window.
    const styles = `<w:styles ${W}>` +
      `<w:style w:type="paragraph" w:styleId="A"><w:pPr><w:numPr>` +
      `<w:ilvl w:val="1"/><w:numId w:val="99"/></w:numPr></w:pPr></w:style>` +
      `<w:style w:type="paragraph" w:styleId="B"><w:pPr><w:numPr>` +
      `<w:ilvl w:val="2"/><w:numId w:val="99"/></w:numPr></w:pPr></w:style></w:styles>`;
    const numbering = `<w:numbering ${W}></w:numbering>`; // numId 99 is defined nowhere

    const read = (styleId: string) =>
      preprocessSide(body(para(styleId, C1)), 'document', 'acceptAll', config,
        buildNumberingContext(styles, numbering)).paras;

    const original = read('A');
    const revised = read('B');
    const result = await runComparePipeline({
      originalParas: original, revisedParas: revised, config,
      author: 'Mutatis', date: '2026-09-12T00:00:00Z',
    });
    const props = properties(result.ops);

    expect(original[0]!.numbering?.label).toBeNull();
    expect(revised[0]!.numbering?.label).toBeNull();
    expect(props).toHaveLength(1);
    expect(props[0]!.propertyChange!.numberingIdentity).toBe(true);
    expect(props[0]!.propertyChange!.from.ilvl).toBe(1);
    expect(props[0]!.propertyChange!.to.ilvl).toBe(2);
    expect(previewText(props[0]!)).toContain('level 2 → level 3');
  });

  it('emits a formatting op for a styleId change that touches no numbering at all', async () => {
    const orig = body(para('Plain', C1));
    const rev = body(para('PlainAlt', C1));

    const { result } = await compare(orig, rev);
    const props = properties(result.ops);

    expect(props).toHaveLength(1);
    expect(props[0]!.classification).toBe('formatting');
    expect(props[0]!.propertyChange).toMatchObject({
      styleId: true,
      numberingIdentity: false,
      numberLabel: false,
      from: { styleId: 'Plain' },
      to: { styleId: 'PlainAlt' },
    });
  });

  it('leaves the AUTHORED REDLINE byte-identical — this is visibility, not new markup', async () => {
    const orig = body(para('Clause1', C1), para('Clause2', C2));
    const rev = body(para('Clause1', C1), para('Clause1', C2));

    const { original, revised, result } = await compare(orig, rev);
    expect(properties(result.ops)).toHaveLength(1);

    // Author the same ops with the property op put back to `equal` — what the
    // pipeline produced before ADR-0042. Byte equality is the ADR's core claim.
    const asEqual = result.ops.map((op) =>
      op.type === 'property' ? { ...op, type: 'equal' as const } : op,
    );
    const xml = authorBody(
      asEqual,
      original,
      revised,
      makeAuthorContext('Mutatis', '2026-09-12T00:00:00Z', 1, { captureClasses: true }),
    );
    expect(xml).toBe(result.authoredBodyXml);
    // No revision markup was invented for a change that has no text delta.
    expect(result.authoredBodyXml).not.toContain('<w:ins ');
    expect(result.authoredBodyXml).not.toContain('<w:del ');
  });

  it('does NOT fold numbering into the hash — the pair still anchors on its text', async () => {
    const orig = body(para('Clause1', C1), para('Clause2', C2));
    const rev = body(para('Clause1', C1), para('Clause1', C2));

    const { original, revised } = await compare(orig, rev);

    expect(original[1]!.numbering).toMatchObject({ numId: '3', ilvl: 1 });
    expect(revised[1]!.numbering).toMatchObject({ numId: '3', ilvl: 0 });
    // The comparison key carries the LEVEL and the level definition's shape, and
    // never the document-local instance id.
    expect(original[1]!.numbering!.key).not.toContain('3');
    expect(original[1]!.numbering!.key).not.toBe(revised[1]!.numbering!.key);
    // Different level, same text ⇒ same hash. If this ever fails, the structural
    // diff will stop pairing re-levelled clauses and start deleting and
    // re-inserting them, which is the defect the whole readability plan closed.
    expect(original[1]!.hash).toBe(revised[1]!.hash);
  });

  it('forces a signature-block style change to substantive (§9.9 carve-out)', async () => {
    const signed = 'SIGNED by the Director for and on behalf of the Vendor';
    const orig = body(para('Plain', signed));
    const rev = body(para('PlainAlt', signed));

    const { result } = await compare(orig, rev);
    const props = properties(result.ops);

    expect(props).toHaveLength(1);
    expect(props[0]!.classification).toBe('substantiveText');
    // §9.9 — "except where they alter a signature/structural paragraph". A
    // signature block is never reduced to a filterable formatting row.
    expect(props[0]!.classes).not.toContain('formatting');
  });

  it('emits nothing when a source has no numbering to resolve (the Text and PDF paths)', async () => {
    // Records built WITHOUT a numbering context: `numbering` is absent on both
    // sides, which must read as "no difference", never as "changed to nothing".
    const original = preprocessSide(body(para('Plain', C1)), 'document', 'acceptAll', config).paras;
    const revised = preprocessSide(body(para('Plain', C1)), 'document', 'acceptAll', config).paras;

    const result = await runComparePipeline({
      originalParas: original,
      revisedParas: revised,
      config,
      author: 'Mutatis',
      date: '2026-09-12T00:00:00Z',
    });

    expect(properties(result.ops)).toHaveLength(0);
    expect(result.summary.total).toBe(0);
  });

  it('COLLAPSES a renumbering cascade instead of putting one row per shifted clause', async () => {
    // The §6.1 pain point, and the one this op type could have caused rather
    // than cured: insert one sub-clause at the head of a list of twelve and
    // every later letter moves. Measured before the collapse existed: 12
    // uncollapsed rows and ZERO cascades, because `cascade.ts` derives a member
    // from `wordDiffs` and a property op has none.
    const subs = Array.from({ length: 12 }, (_, i) =>
      para('Clause2', `Sub-clause number ${i + 1} sets out an obligation of the Vendor.`),
    );
    const orig = body(para('Clause1', C1), ...subs);
    const rev = body(
      para('Clause1', C1),
      para('Clause2', 'A newly inserted sub-clause sits at the head of the list.'),
      ...subs,
    );

    const { result } = await compare(orig, rev);
    const props = properties(result.ops);

    expect(props).toHaveLength(12);
    expect(result.summary.renumberingCascades).toHaveLength(1);
    const cascade = result.summary.renumberingCascades[0]!;
    expect(cascade.source).toBe('numbering');
    expect(cascade.shift).toBe(1);
    expect(cascade.affectedCount).toBe(12);
    expect(cascade.tokenKind).toContain('level 2');
    expect(cascade.exampleRefs[0]).toBe('(a) → (b)');
    // Every member is collapsed, so the change list shows ONE row, expandable —
    // and ADR-0008 holds: not one op was dropped.
    expect(props.every((op) => op.collapsedInto === cascade.id)).toBe(true);
    expect(result.summary.total).toBe(13); // 12 collapsed members + the insert
  });

  it('does not collapse a two-member shift — the §6.1 threshold applies here too', async () => {
    const subs = Array.from({ length: 2 }, (_, i) =>
      para('Clause2', `Sub-clause number ${i + 1} sets out an obligation of the Vendor.`),
    );
    const orig = body(para('Clause1', C1), ...subs);
    const rev = body(para('Clause1', C1), para('Clause2', 'An inserted sub-clause.'), ...subs);

    const { result } = await compare(orig, rev);

    expect(properties(result.ops)).toHaveLength(2);
    expect(result.summary.renumberingCascades).toHaveLength(0);
    expect(properties(result.ops).every((op) => op.collapsedInto === undefined)).toBe(true);
  });

  it('does not report a label change when only ONE side resolved a label', () => {
    const base = {
      index: 0, partId: 'document', rawXml: '', normText: C1, normTextAggressive: C1,
      numberingPrefix: '', hash: 'h', styleId: 'Clause1', isTableCell: false,
      isSignatureBlock: false, isHeaderFooter: false, isFootnote: false,
    } satisfies Omit<ParaRecord, 'numbering'>;
    const num = (label: string | null, path: readonly number[] | null) => ({
      numId: '3', ilvl: 1, key: '1|lowerLetter|(%2)|false', label, path,
      // `(%2)` substitutes level 1 and nothing else, so that is all it renders.
      renders: path === null ? null : [1],
    });

    // Same identity; one side's level definition did not resolve to a label.
    // That is a gap in the numbering part, not a change in the document.
    const orig: ParaRecord = { ...base, numbering: num('(a)', [1, 1]) };
    const rev: ParaRecord = { ...base, numbering: num(null, null) };

    expect(comparePropertyPair(orig, rev)).toBeNull();
  });
});

/**
 * The cascade that a paragraph's OWN counter cannot see.
 *
 * `%1.%2` — "1.1", "1.2" — is the dominant SG/E&W clause style, and under it a
 * sub-clause renumbers when its PARENT does while its own counter never moves.
 * The first implementation derived the cascade shift by subtracting own-level
 * ordinals, so every one of those sub-clauses measured a shift of ZERO, §6.1's
 * collapse dropped zero-shift members, and one inserted clause put one
 * uncollapsed row in the change list for every sub-clause in the document —
 * measured at six rows and zero cascades, with the inserted clause's own row
 * alone in a group of one, below the threshold. Neither default filter hides
 * these, so the reader is handed the flood §6.1 exists to prevent.
 *
 * The suite above cannot catch it: its level-1 `lvlText` is `(%2)`, which shows
 * the own counter only, so a top-level insert leaves the sub-clause labels
 * alone and there is nothing to collapse. This block is that gap.
 */
/**
 * RECORDED LIMIT — `w:pPrChange` is not reverted by `rejectAll` (ADR-0042
 * § Accepted limits — cited by its subject, not its ordinal, because the list
 * has been renumbered once already and the citation went stale).
 *
 * `preprocess/flatten.ts` handles `w:ins`/`w:del`/`w:moveFrom`/`w:moveTo` and
 * `quarantine.ts` drops the same four plus comment markers. `w:pPrChange` — the
 * Word-native record of a tracked PARAGRAPH-PROPERTY change — is in neither, so
 * under `rejectAll` a paragraph's TEXT reverts and its `w:pStyle`/`w:numPr` do
 * not. The user asked to reject the counterparty's changes and the re-levelling
 * stays.
 *
 * This was harmless until this plan: `styleId` was captured and read by nothing,
 * and numbering was not resolved at all, so nothing could observe the surviving
 * property. ADR-0042 makes it observable, which is why the limit is recorded
 * here rather than left to be discovered. The DEFAULT (`acceptAll` on both
 * sides) is unaffected — `rejectAll` is a panel choice.
 *
 * Spec §5.0 names only the four revision wrappers, so this is a limit and not a
 * violation. Change this test only together with ADR-0042.
 */
describe('a tracked paragraph-property change under rejectAll', () => {
  const STYLES_PC = `<w:styles ${W}>
    <w:style w:type="paragraph" w:styleId="C1"><w:pPr><w:numPr>
      <w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style>
    <w:style w:type="paragraph" w:styleId="C2"><w:pPr><w:numPr>
      <w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style>
  </w:styles>`;
  const NUM_PC = `<w:numbering ${W}><w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%2)"/></w:lvl>
    </w:abstractNum><w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;

  const TEXT = 'The Vendor shall deliver the Completion Documents.';
  const plain = (styleId: string) =>
    body(`<w:p><w:pPr><w:pStyle w:val="${styleId}"/></w:pPr>` +
      `<w:r><w:t xml:space="preserve">${TEXT}</w:t></w:r></w:p>`);
  /** The same paragraph outdented WITH TRACKING ON: `w:pPrChange` holds the old pPr. */
  const tracked =
    body(`<w:p><w:pPr><w:pStyle w:val="C1"/>` +
      `<w:pPrChange w:id="1" w:author="Counterparty" w:date="2026-01-01T00:00:00Z">` +
      `<w:pPr><w:pStyle w:val="C2"/></w:pPr></w:pPrChange></w:pPr>` +
      `<w:r><w:t xml:space="preserve">${TEXT}</w:t></w:r></w:p>`);

  async function run(revisedPolicy: 'acceptAll' | 'rejectAll') {
    const cfg = { ...config, flattenOriginal: 'acceptAll' as const, flattenRevised: revisedPolicy };
    const read = (xml: string, policy: 'acceptAll' | 'rejectAll'): ParaRecord[] =>
      preprocessSide(xml, 'document', policy, cfg, buildNumberingContext(STYLES_PC, NUM_PC)).paras;
    const revised = read(tracked, revisedPolicy);
    const result = await runComparePipeline({
      originalParas: read(plain('C2'), 'acceptAll'),
      revisedParas: revised,
      config: cfg,
      author: 'Mutatis',
      date: '2026-09-12T00:00:00Z',
    });
    return { revised, result };
  }

  it('is reported under acceptAll, which is the default and is correct', async () => {
    const { revised, result } = await run('acceptAll');
    expect(revised[0]!.styleId).toBe('C1');
    expect(revised[0]!.numbering?.label).toBe('1.');
    expect(properties(result.ops)).toHaveLength(1);
  });

  it('SURVIVES rejectAll — the text reverts, the paragraph properties do not', async () => {
    const { revised, result } = await run('rejectAll');

    // Rejecting should have restored `C2` / `(a)`. It does not: `w:pPrChange` is
    // not one of the wrappers `flatten.ts` unwinds, so the revised side still
    // reads the ACCEPTED level while its text reads the rejected wording.
    expect(revised[0]!.styleId).toBe('C1');
    expect(revised[0]!.numbering?.label).toBe('1.');
    expect(properties(result.ops)).toHaveLength(1);
    expect(result.summary.total).toBe(1);
  });
});

describe('a cascade carried by an ANCESTOR level (lvlText "%1.%2")', () => {
  const STYLES_ML = `<w:styles ${W}>
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:pPr/></w:style>
    <w:style w:type="paragraph" w:styleId="L1"><w:pPr>
      <w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style>
    <w:style w:type="paragraph" w:styleId="L2"><w:pPr>
      <w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style>
  </w:styles>`;

  const NUMBERING_ML = `<w:numbering ${W}>
    <w:abstractNum w:abstractNumId="0">
      <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
      <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl>
    </w:abstractNum>
    <w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>
  </w:numbering>`;

  const sideML = (xml: string, policy: 'acceptAll' | 'rejectAll'): ParaRecord[] =>
    preprocessSide(xml, 'document', policy, config, buildNumberingContext(STYLES_ML, NUMBERING_ML))
      .paras;

  async function compareML(origXml: string, revXml: string) {
    const original = sideML(origXml, config.flattenOriginal);
    const revised = sideML(revXml, config.flattenRevised);
    const result = await runComparePipeline({
      originalParas: original,
      revisedParas: revised,
      config,
      author: 'Mutatis',
      date: '2026-09-12T00:00:00Z',
    });
    return { original, revised, result };
  }

  const SUBS = Array.from({ length: 6 }, (_, i) =>
    para('L2', `Sub-clause number ${i + 1} binds the Vendor under this Agreement.`),
  );
  const HEAD = para('L1', 'The Vendor covenants with the Purchaser as follows.');
  const ORIG = body(HEAD, ...SUBS);
  const REV = body(para('L1', 'A newly inserted top-level clause precedes the covenant.'), HEAD, ...SUBS);

  const DEFS = (l0fmt: string, l0start: string) => `<w:numbering ${W}>
    <w:abstractNum w:abstractNumId="0">
      <w:lvl w:ilvl="0"><w:start w:val="${l0start}"/><w:numFmt w:val="${l0fmt}"/><w:lvlText w:val="%1."/></w:lvl>
      <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl>
    </w:abstractNum>
    <w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>
  </w:numbering>`;

  /** The same body read through two different numbering definitions. */
  async function compareDefs(origNum: string, revNum: string) {
    const read = (numXml: string, policy: 'acceptAll' | 'rejectAll'): ParaRecord[] =>
      preprocessSide(ORIG, 'document', policy, config, buildNumberingContext(STYLES_ML, numXml))
        .paras;
    const original = read(origNum, config.flattenOriginal);
    const revised = read(revNum, config.flattenRevised);
    const result = await runComparePipeline({
      originalParas: original,
      revisedParas: revised,
      config,
      author: 'Mutatis',
      date: '2026-09-12T00:00:00Z',
    });
    return { original, revised, result };
  }

  // ── The ANCESTOR half of the two redefinition tests above ─────────────────
  // Both of those use a single-level list, so the level redefined is the only
  // level any label is drawn from. That is why this defect survived five review
  // passes: under `%1.%2` a sub-clause's number is rendered from its ANCESTORS'
  // formats too, so redefining level 0 moves every sub-clause's label while the
  // sub-clause's own definition is untouched. A key covering the own level alone
  // left those keys EQUAL — and a redefinition moves no counter either, so the
  // shift was null and there was no cascade to collapse into. Measured then:
  // twelve rows, each classified `crossRefRenumber`, i.e. each asserting "its
  // number moved because a sibling did", with no sibling anywhere.

  it('reports an ANCESTOR format redefinition as formatting on every level it reaches', async () => {
    const { original, revised, result } = await compareDefs(
      DEFS('decimal', '1'),
      DEFS('upperRoman', '1'),
    );
    const props = properties(result.ops);

    // Level 0 renders from its own format; level 1 renders from level 0's too.
    expect(original.map((r) => r.numbering?.label).slice(0, 3)).toEqual(['1.', '1.1', '1.2']);
    expect(revised.map((r) => r.numbering?.label).slice(0, 3)).toEqual(['I.', 'I.1', 'I.2']);

    expect(props).toHaveLength(7);
    // Every one of them: this paragraph's numbering WAS changed. Not one of them
    // is a bystander, so not one of them may say a sibling moved.
    expect(props.every((op) => op.classification === 'formatting')).toBe(true);
    expect(result.summary.byCategory.crossRefRenumber).toBe(0);
    expect(result.summary.renumberingCascades).toHaveLength(0);
    // ADR-0008: categorised, never hidden. The total counts all seven, and the
    // default *Hide formatting-only* filter (§8.2, ON) is what keeps them out of
    // the navigation list — a user-visible switch, not a suppression.
    expect(result.summary.total).toBe(7);
    expect(result.summary.byCategory.formatting).toBe(7);
  });

  it('reports an ANCESTOR <w:start> change as formatting, not as a +4 sibling shift', async () => {
    const { original, revised, result } = await compareDefs(DEFS('decimal', '1'), DEFS('decimal', '5'));
    const props = properties(result.ops);

    expect(original.map((r) => r.numbering?.label).slice(0, 3)).toEqual(['1.', '1.1', '1.2']);
    expect(revised.map((r) => r.numbering?.label).slice(0, 3)).toEqual(['5.', '5.1', '5.2']);

    expect(props).toHaveLength(7);
    expect(props.every((op) => op.classification === 'formatting')).toBe(true);
    // The wrong answer this replaces: six sub-clauses collapsed into
    // "Clause numbering (level 1) shifted (+4)" — a tidy row saying something
    // that did not happen.
    expect(result.summary.renumberingCascades).toHaveLength(0);
  });

  // ── The key must be no WIDER than the label either ────────────────────────
  // Having been too narrow three times, the key was then made too wide: every
  // field of every level 0…ilvl. A key wider than the label reports paragraphs
  // that render identically end to end — measured at 240 reported changes on a
  // 40x5 document where 40 paragraphs changed. It now carries only what
  // `renderListNumbers` actually reads for THIS paragraph.

  it('is SILENT on an ancestor lvlText change, which no descendant substitutes', async () => {
    // Re-wording the top-level prefix: `%1.` → `Clause %1`. `applyLvlText`
    // substitutes `%1`/`%2` from the OWN level's template only, so `1.1` stays
    // `1.1` — the sub-clauses render byte-identically.
    const before = DEFS('decimal', '1');
    const after = before.replace('w:lvlText w:val="%1."', 'w:lvlText w:val="Clause %1"');
    const { original, revised, result } = await compareDefs(before, after);
    const props = properties(result.ops);

    expect(original[0]!.numbering?.label).toBe('1.');
    expect(revised[0]!.numbering?.label).toBe('Clause 1');
    expect(original[1]!.numbering?.label).toBe('1.1');
    expect(revised[1]!.numbering?.label).toBe('1.1'); // unchanged

    // Exactly the one paragraph whose own template changed.
    expect(props).toHaveLength(1);
    expect(props[0]!.propertyChange?.from.numberLabel).toBe('1.');
    expect(props[0]!.propertyChange?.to.numberLabel).toBe('Clause 1');

    // …and it is the LIST that changed, not a sibling that moved. Without this
    // pair of assertions the test passes with `own.lvlText` dropped from the key
    // entirely — see the test below, which is the guard for that.
    expect(props[0]!.propertyChange?.numberingIdentity).toBe(true);
    expect(props[0]!.classification).toBe('formatting');
  });

  it('calls a RE-WORDED own template a changed list, not a sibling moving', async () => {
    // `own.lvlText` is the FIRST field of the comparison key, and it was the one
    // field of it that no test held: deleting it left the whole suite green —
    // 2376 passed, 0 failed — while turning this case from `formatting` into a
    // false `crossRefRenumber`.
    //
    // The case is the commonest style edit a firm makes: re-word the sub-clause
    // template and keep the levels it substitutes. `%1.%2` → `Clause %1.%2`
    // moves every sub-clause's label and changes nothing about any sibling, so
    // §9.4's answer is *"the numbering of these paragraphs was changed"*. Read
    // without `lvlText`, the key sees the same levels, the same formats, the
    // same starts — identical — and the only visible difference is the label,
    // which is the signature of a bystander. The op then says *"its number moved
    // because a sibling did"* with no sibling anywhere, and lands in
    // `crossRefRenumber`, which *Hide formatting-only* does NOT hide.
    const before = DEFS('decimal', '1');
    const after = before.replace('w:lvlText w:val="%1.%2"', 'w:lvlText w:val="Clause %1.%2"');
    const { original, revised, result } = await compareDefs(before, after);
    const props = properties(result.ops);

    // The substituted-level SET is identical on both sides — that is what makes
    // this the case the rest of the key cannot see.
    expect(original[1]!.numbering?.renders).toEqual(revised[1]!.numbering?.renders);
    expect(original[1]!.numbering?.label).toBe('1.1');
    expect(revised[1]!.numbering?.label).toBe('Clause 1.1');

    expect(props.length).toBeGreaterThan(0);
    for (const op of props) {
      expect(op.propertyChange?.numberingIdentity).toBe(true);
      expect(op.classification).toBe('formatting');
    }
    expect(result.summary.byCategory.crossRefRenumber).toBe(0);
    expect(result.summary.renumberingCascades).toHaveLength(0);
  });

  it('is SILENT on an ancestor `isLgl` that renders nothing differently', async () => {
    // `<w:isLgl/>` forces every substituted token to `decimal`. On a list whose
    // formats are already decimal it changes no label anywhere — and it is read
    // from the OWN level, so it cannot reach a descendant at all.
    const before = DEFS('decimal', '1');
    const after = before.replace(
      '<w:lvlText w:val="%1."/>',
      '<w:lvlText w:val="%1."/><w:isLgl/>',
    );
    const { original, revised, result } = await compareDefs(before, after);

    expect(revised.map((r) => r.numbering?.label)).toEqual(
      original.map((r) => r.numbering?.label),
    );
    expect(properties(result.ops)).toHaveLength(0);
    expect(result.summary.total).toBe(0);
  });

  it('is SILENT on an ancestor format its OWN `isLgl` has already overruled', async () => {
    // The key's `isLgl` fold, which nothing held: revert
    // `own.isLgl ? 'decimal' : …` in `numberingKey` and the whole suite stays
    // green at 2391 while this list reports one spurious `formatting` op per
    // paragraph, scaling 1:1 with its length. That is the pass-6 failure — 240
    // changes where 40 paragraphs changed — wearing a different hat, and it
    // contradicts ADR-0042's *"two fields as BEHAVIOUR, not as written"*.
    //
    // `<w:isLgl/>` on the paragraph's OWN level forces every substituted token
    // to `decimal`, so level 0's `numFmt` is overruled before it can reach the
    // label: `lowerRoman` → `decimal` up there changes nothing down here. The
    // twin fold in `renderedLevels` is guarded by the `numFmt="none"` test
    // further down; this is the one in the key.
    //
    // The body carries NO level-0 paragraph on purpose. Add one and its own
    // label moves (`i.` → `1.`), so the op count differs with or without the
    // fold and the guard passes for a reason that has nothing to do with it.
    // Silence against noise is the only shape that isolates this.
    const lgl = (l0fmt: string) =>
      DEFS(l0fmt, '1').replace(
        '<w:lvlText w:val="%1.%2"/>',
        '<w:lvlText w:val="%1.%2"/><w:isLgl/>',
      );
    const subsOnly = body(...SUBS.slice(0, 2));
    const read = (numXml: string, policy: 'acceptAll' | 'rejectAll'): ParaRecord[] =>
      preprocessSide(subsOnly, 'document', policy, config,
        buildNumberingContext(STYLES_ML, numXml)).paras;
    const original = read(lgl('lowerRoman'), config.flattenOriginal);
    const revised = read(lgl('decimal'), config.flattenRevised);
    const result = await runComparePipeline({
      originalParas: original, revisedParas: revised, config,
      author: 'Mutatis', date: '2026-09-12T00:00:00Z',
    });

    expect(original.map((r) => r.numbering?.label)).toEqual(['1.1', '1.2']);
    expect(revised.map((r) => r.numbering?.label)).toEqual(['1.1', '1.2']);
    expect(properties(result.ops)).toHaveLength(0);
    expect(result.summary.total).toBe(0);
  });

  it('REPORTS the own level\'s format changing under `isLgl`, which the fold erased', async () => {
    // A late review found this, the worst defect in this area after
    // an earlier review's: `render-list-numbers.ts` applies `isLgl` to the paragraph's OWN
    // token as well as to its ancestors, Word does not (measured against the
    // reference rendering: `decimalZero` + `isLgl` over an `upperRoman` ancestor
    // gives `1.01` — ancestor converted, own format kept), and this key MIRRORED
    // the renderer's fold. So both detectors went blind at once and an edit to
    // the own level\'s `numFmt` produced NO OP AT ALL:
    //
    //     with `<w:isLgl/>`      labels 1.1 1.2 -> 1.1 1.2    total 0
    //     the same edit without  labels 1.01 …  -> 1.1 …      total 3
    //
    // Word renders the first pair `1.01` against `1.1`, so a reader sees every
    // sub-clause change and the summary said nothing — ADR-0008 suppression, on
    // a shape that is common in real-world
    // `numbering.xml` parts. The plan and ADR-0042 both said the renderer defect
    // *"cannot affect this plan, because both sides render identically"*: true
    // of the LABELS and false of the document, which is the distinction the key
    // exists to cover (ADR-0042 § *Where the key still exceeds the label*).
    //
    // The labels are deliberately IDENTICAL on both sides here. Under the
    // shipped renderer they always will be — that is the whole defect — so a
    // guard that asserted a label move would pass with the fix reverted.
    const lgl = (ownFmt: string) =>
      DEFS('decimal', '1')
        .replace('<w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/>',
          `<w:numFmt w:val="${ownFmt}"/><w:lvlText w:val="%1.%2"/><w:isLgl/>`);
    const subsOnly = body(...SUBS.slice(0, 2));
    const read = (numXml: string, policy: 'acceptAll' | 'rejectAll'): ParaRecord[] =>
      preprocessSide(subsOnly, 'document', policy, config,
        buildNumberingContext(STYLES_ML, numXml)).paras;
    const original = read(lgl('decimalZero'), config.flattenOriginal);
    const revised = read(lgl('decimal'), config.flattenRevised);
    const result = await runComparePipeline({
      originalParas: original, revisedParas: revised, config,
      author: 'Mutatis', date: '2026-09-12T00:00:00Z',
    });

    expect(original.map((r) => r.numbering?.label)).toEqual(['1.1', '1.2']);
    expect(revised.map((r) => r.numbering?.label)).toEqual(['1.1', '1.2']);
    const props = properties(result.ops);
    expect(props).toHaveLength(2);
    for (const op of props) {
      expect(op.propertyChange?.numberingIdentity).toBe(true);
      expect(op.propertyChange?.numberLabel).toBe(false);
      expect(op.classification).toBe('formatting');
    }
    expect(result.summary.total).toBe(2);
  });

  it('is SILENT when `isLgl` is added to a level whose own format is already decimal', async () => {
    // The wrong-fire direction of the test above, and the reason the new field
    // is pushed unconditionally rather than only when `<w:isLgl/>` is present.
    // Pushed conditionally, the field made the PRESENCE of the element visible
    // and invented a change here — where the own format is `decimal` on both
    // sides, so neither Word nor `renderListNumbers` renders anything
    // differently. `<w:isLgl/>` on an all-decimal list is the original false
    // positive this function's header records; it must stay silent.
    const withLgl = DEFS('decimal', '1').replace(
      '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/>',
      '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/><w:isLgl/>');
    const subsOnly = body(...SUBS.slice(0, 2));
    const read = (numXml: string, policy: 'acceptAll' | 'rejectAll'): ParaRecord[] =>
      preprocessSide(subsOnly, 'document', policy, config,
        buildNumberingContext(STYLES_ML, numXml)).paras;
    const original = read(DEFS('decimal', '1'), config.flattenOriginal);
    const revised = read(withLgl, config.flattenRevised);
    const result = await runComparePipeline({
      originalParas: original, revisedParas: revised, config,
      author: 'Mutatis', date: '2026-09-12T00:00:00Z',
    });

    expect(revised.map((r) => r.numbering?.label))
      .toEqual(original.map((r) => r.numbering?.label));
    expect(properties(result.ops)).toHaveLength(0);
    expect(result.summary.total).toBe(0);
  });

  it('is SILENT on `lvlRestart` at level 0, which has nothing shallower to restart it', async () => {
    // ECMA-376 §17.9.10: a level restarts when a SHALLOWER level advances.
    // Level 0 has none, so the renderer never consults its `lvlRestart`.
    const before = DEFS('decimal', '1');
    const after = before.replace(
      '<w:lvlText w:val="%1."/>',
      '<w:lvlText w:val="%1."/><w:lvlRestart w:val="0"/>',
    );
    const { original, revised, result } = await compareDefs(before, after);

    expect(revised.map((r) => r.numbering?.label)).toEqual(
      original.map((r) => r.numbering?.label),
    );
    expect(properties(result.ops)).toHaveLength(0);
  });

  it('is SILENT on an ancestor format the sub-level template never substitutes', async () => {
    // A common shape: sub-clauses labelled `(a)`, `(b)` — an own
    // level only template. Redefining level 0 moves the top-level labels and
    // leaves every `(a)` exactly where it was.
    const lettered = (l0fmt: string) =>
      DEFS(l0fmt, '1').replace('w:lvlText w:val="%1.%2"', 'w:lvlText w:val="(%2)"');
    const { original, revised, result } = await compareDefs(
      lettered('decimal'),
      lettered('upperRoman'),
    );
    const props = properties(result.ops);

    expect(original[1]!.numbering?.label).toBe('(1)');
    expect(revised[1]!.numbering?.label).toBe('(1)'); // `(%2)` renders the own counter only
    // One op: the single level-0 paragraph. Not the six sub-clauses beneath it.
    expect(props).toHaveLength(1);
    expect(props[0]!.propertyChange?.from.numberLabel).toBe('1.');
    expect(props[0]!.propertyChange?.to.numberLabel).toBe('I.');
  });

  // ── The counter machine is a SEQUENCE, not a per-paragraph function ───────
  // The key's contract is "what the renderer reads". That was implemented as
  // what the renderer reads for THIS paragraph, which is not the same thing:
  // `renderListNumbers` walks every paragraph in document order and SKIPS one
  // whose level does not resolve, so it neither advances that level nor resets
  // the deeper ones. Deleting a level definition therefore moves the counters of
  // the paragraphs below it while touching none of their own fields.

  it('does not call a DELETED ancestor level a sibling moving', async () => {
    // Three levels; the sub-points' template is `(%3)`, so level 1 is outside
    // everything about their own token. Remove level 1's definition and the
    // level-1 paragraphs stop resolving, so level 2 never resets and the second
    // block's points run on: `(a)`,`(b)` → `(c)`,`(d)`.
    //
    // Measured before the fix: those two reported `crossRefRenumber` — §9.4's
    // *"this paragraph did not change; its number moved because a sibling did"* —
    // with no sibling anywhere and the list itself redefined. Exactly the false
    // statement the identity comparison exists to prevent.
    const L0 = '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>';
    const L1 = '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl>';
    const L2 = '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%3)"/></w:lvl>';
    const scheme = (lvls: string) =>
      `<w:numbering ${W}><w:abstractNum w:abstractNumId="0">${lvls}</w:abstractNum>` +
      '<w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num></w:numbering>';
    const styles = `<w:styles ${W}>` +
      [0, 1, 2]
        .map(
          (k) =>
            `<w:style w:type="paragraph" w:styleId="D${k}"><w:pPr><w:numPr>` +
            `<w:ilvl w:val="${k}"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style>`,
        )
        .join('') +
      '</w:styles>';
    const doc = body(
      para('D0', 'The Vendor covenants with the Purchaser as follows.'),
      para('D1', 'First sub-clause of the covenant.'),
      para('D2', 'A point arising under the first sub-clause.'),
      para('D1', 'Second sub-clause of the covenant.'),
      para('D2', 'A point arising under the second sub-clause.'),
    );

    const read = (lvls: string, policy: 'acceptAll' | 'rejectAll'): ParaRecord[] =>
      preprocessSide(doc, 'document', policy, config, buildNumberingContext(styles, scheme(lvls)))
        .paras;
    const original = read(L0 + L1 + L2, config.flattenOriginal);
    const revised = read(L0 + L2, config.flattenRevised); // level 1 deleted
    const result = await runComparePipeline({
      originalParas: original, revisedParas: revised, config,
      author: 'Mutatis', date: '2026-09-12T00:00:00Z',
    });
    const props = properties(result.ops);

    // The sub-point after the (now unresolvable) second sub-clause runs on.
    expect(original[4]!.numbering?.label).toBe('(a)');
    expect(revised[4]!.numbering?.label).toBe('(b)');

    // Every one of them: this list was redefined. Not one is a bystander.
    expect(props.every((op) => op.classification === 'formatting')).toBe(true);
    expect(result.summary.byCategory.crossRefRenumber).toBe(0);
    expect(result.summary.renumberingCascades).toHaveLength(0);
  });

  it('RECORDS its limit: a level deleted DEEPER than the paragraph reads as a sibling', async () => {
    // The other side of the same coin, and the boundary of what a per-paragraph
    // key can answer.
    //
    // `numberingKey` covers exactly what the renderer reads and writes WHEN
    // RENDERING THIS PARAGRAPH: levels 0…`ilvl`, whose formats it substitutes and
    // whose counters it touches. Levels deeper than `ilvl` are touched by OTHER
    // paragraphs' render calls, and a key is not a function of those.
    //
    // The renderer initialises a shallower level's counter when it renders a
    // DEEPER paragraph (`render-list-numbers.ts`, the `k <= para.ilvl` loop), and
    // skips that paragraph entirely when its own level does not resolve. So
    // deleting level 2's definition moves a level-1 paragraph's counter —
    // `1.2` → `1.1` — while nothing in that paragraph's own key moves.
    //
    // What is GUARANTEED, and is what keeps this a limit rather than a defect:
    // `numberLabel` is compared from the two RENDERED labels, never from the key,
    // so the move is always reported. The key decides only which of §9.4's two
    // sentences it gets, and the failure direction is over-report — into
    // `crossRefRenumber`, which `hideFormattingOnly` does not touch — never
    // suppression (ADR-0008). ADR-0042 § Accepted limits.
    const L0 = '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>';
    const L1 = '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl>';
    const L2 = '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%3)"/></w:lvl>';
    const scheme = (lvls: string) =>
      `<w:numbering ${W}><w:abstractNum w:abstractNumId="0">${lvls}</w:abstractNum>` +
      '<w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num></w:numbering>';
    const styles = `<w:styles ${W}>` +
      [0, 1, 2]
        .map(
          (k) =>
            `<w:style w:type="paragraph" w:styleId="D${k}"><w:pPr><w:numPr>` +
            `<w:ilvl w:val="${k}"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style>`,
        )
        .join('') +
      '</w:styles>';
    // The level-2 point comes BEFORE the level-1 sub-clause, so it is what
    // initialises level 1's counter. That ordering is the trigger.
    const doc = body(
      para('D0', 'The Vendor covenants with the Purchaser as follows.'),
      para('D2', 'A point standing under the covenant itself.'),
      para('D1', 'The sub-clause that follows that point.'),
    );
    const read = (lvls: string, policy: 'acceptAll' | 'rejectAll'): ParaRecord[] =>
      preprocessSide(doc, 'document', policy, config, buildNumberingContext(styles, scheme(lvls)))
        .paras;
    const original = read(L0 + L1 + L2, config.flattenOriginal);
    const revised = read(L0 + L1, config.flattenRevised); // level 2 deleted
    const result = await runComparePipeline({
      originalParas: original, revisedParas: revised, config,
      author: 'Mutatis', date: '2026-09-12T00:00:00Z',
    });
    const props = properties(result.ops);

    // The label really does move, and both sides render it.
    expect(original[2]!.numbering?.label).toBe('1.2');
    expect(revised[2]!.numbering?.label).toBe('1.1');

    // NOT HIDDEN — the change is reported, which is the invariant that matters.
    const moved = props.find((op) => op.propertyChange?.from.numberLabel === '1.2');
    expect(moved).toBeDefined();
    expect(moved?.propertyChange?.numberLabel).toBe(true);
    expect(moved?.propertyChange?.to.numberLabel).toBe('1.1');

    // The recorded limit: its own key did not move, so §9.4 calls it a bystander
    // of a sibling's renumbering. No sibling renumbered; the list was redefined
    // under it. Change this expectation only together with ADR-0042.
    expect(moved?.propertyChange?.numberingIdentity).toBe(false);
    expect(moved?.classification).toBe('crossRefRenumber');

    // The level-2 paragraph itself is not in doubt — its own definition is gone.
    const relevelled = props.find((op) => op.propertyChange?.from.numberLabel === '(a)');
    expect(relevelled?.propertyChange?.numberingIdentity).toBe(true);
    expect(relevelled?.classification).toBe('formatting');
  });

  // ── `lvlRestart` is stored as its BEHAVIOUR, not as written ───────────────
  // `render-list-numbers.ts` resolves an absent `<w:lvlRestart>` to the level's
  // own index and then asks whether the advancing level is shallower than it. So
  // at level k, every written value at or above k — and an absent element — mean
  // the same thing, and only 1…k-1 differ. Storing the raw value reported a
  // semantic no-op as twenty changes with both label lists identical.

  it('is SILENT on a `lvlRestart` that says what absence already said', async () => {
    const withRestart = (v: string | null) =>
      DEFS('decimal', '1').replace(
        '<w:lvlText w:val="%1.%2"/>',
        `<w:lvlText w:val="%1.%2"/>${v === null ? '' : `<w:lvlRestart w:val="${v}"/>`}`,
      );

    // Absent → 1 at level 1: identical behaviour.
    const a = await compareDefs(withRestart(null), withRestart('1'));
    expect(a.revised.map((r) => r.numbering?.label)).toEqual(
      a.original.map((r) => r.numbering?.label),
    );
    expect(properties(a.result.ops)).toHaveLength(0);

    // 1 → 4 at level 1: every value at or above the level is the same thing.
    const b = await compareDefs(withRestart('1'), withRestart('4'));
    expect(properties(b.result.ops)).toHaveLength(0);
  });

  it('DOES report `lvlRestart` 0, which is the one value that changes behaviour', async () => {
    // `0` means never auto-restart, so the sub-clause numbers run on across
    // top-level clauses instead of restarting under each.
    const withRestart = (v: string | null) =>
      DEFS('decimal', '1').replace(
        '<w:lvlText w:val="%1.%2"/>',
        `<w:lvlText w:val="%1.%2"/>${v === null ? '' : `<w:lvlRestart w:val="${v}"/>`}`,
      );
    const { result } = await compareDefs(withRestart(null), withRestart('0'));

    expect(properties(result.ops).length).toBeGreaterThan(0);
    expect(properties(result.ops).every((op) => op.classification === 'formatting')).toBe(true);
  });

  it('does not print the same label on both sides of the arrow', async () => {
    // The identity can change while the rendered label stands still — ADR-0042
    // § "Where the key still exceeds the label" says so, and this is what a
    // reader then sees. `<w:lvlRestart w:val="0"/>` on the sub-clause level of
    // a list with ONE top-level clause: nothing runs on, so every label is
    // exactly where it was, and the op is a list redefinition with
    // `numberLabel` false.
    //
    // Both property surfaces build their before/after from `numberLabel`, and
    // both carried a fallback for "no label resolved" and none for "label
    // resolved identically" — so they printed `1.1 → 1.1`, and the exported
    // report printed a row whose two sides were the same string. That is the
    // exact artefact `propertySides`' own doc comment cites as the reason it
    // exists.
    const withRestart0 = DEFS('decimal', '1').replace(
      '<w:lvlText w:val="%1.%2"/>',
      '<w:lvlText w:val="%1.%2"/><w:lvlRestart w:val="0"/>',
    );
    const oneClause = body(HEAD, ...SUBS.slice(0, 2));
    const read = (numXml: string, policy: 'acceptAll' | 'rejectAll'): ParaRecord[] =>
      preprocessSide(oneClause, 'document', policy, config,
        buildNumberingContext(STYLES_ML, numXml)).paras;
    const original = read(DEFS('decimal', '1'), config.flattenOriginal);
    const revised = read(withRestart0, config.flattenRevised);
    const result = await runComparePipeline({
      originalParas: original, revisedParas: revised, config,
      author: 'Mutatis', date: '2026-09-12T00:00:00Z',
    });

    // Every label is provably where it was.
    expect(revised.map((r) => r.numbering?.label)).toEqual(
      original.map((r) => r.numbering?.label),
    );
    const props = properties(result.ops);
    expect(props.length).toBeGreaterThan(0);
    for (const op of props) {
      expect(op.propertyChange?.numberingIdentity).toBe(true);
      expect(op.propertyChange?.numberLabel).toBe(false);
      const preview = previewText(op);
      expect(preview).not.toMatch(/^(.+) → \1$/);
      expect(preview).toContain('list definition changed');
    }
  });

  it('is SILENT on a NEGATIVE `lvlRestart`, which is `0` said differently', async () => {
    // `parseIntVal` uses `parseInt`, so `<w:lvlRestart w:val="-1"/>` parses to
    // -1 rather than being rejected. The renderer asks `advancedIlvl <=
    // restart - 1` and `advancedIlvl` is never negative, so every value at or
    // below zero means the same thing: never auto-restart. Store the raw value
    // and `0` vs `-1` — a document that renders identically — reports a list
    // redefinition on every member of the list.
    const withRestart = (v: string) =>
      DEFS('decimal', '1').replace(
        '<w:lvlText w:val="%1.%2"/>',
        `<w:lvlText w:val="%1.%2"/><w:lvlRestart w:val="${v}"/>`,
      );
    const { original, revised, result } = await compareDefs(withRestart('0'), withRestart('-1'));

    expect(revised.map((r) => r.numbering?.label)).toEqual(
      original.map((r) => r.numbering?.label),
    );
    expect(properties(result.ops)).toHaveLength(0);
    expect(result.summary.total).toBe(0);
  });

  it('invents nothing when the two definitions are identical', async () => {
    const { result } = await compareDefs(DEFS('decimal', '1'), DEFS('decimal', '1'));
    expect(properties(result.ops)).toHaveLength(0);
    expect(result.summary.total).toBe(0);
  });

  it('renders the ancestor-carried labels the reader would actually see move', async () => {
    const { original, revised } = await compareML(ORIG, REV);
    expect(original.map((r) => r.numbering?.label)).toEqual([
      '1.', '1.1', '1.2', '1.3', '1.4', '1.5', '1.6',
    ]);
    expect(revised.map((r) => r.numbering?.label)).toEqual([
      '1.', '2.', '2.1', '2.2', '2.3', '2.4', '2.5', '2.6',
    ]);
  });

  it('measures the shift at the level that MOVED, not at the paragraph own level', async () => {
    const { result } = await compareML(ORIG, REV);
    const cascades = properties(result.ops).map((op) => op.propertyChange?.cascade);

    // Every row — the level-0 clause and all six level-1 sub-clauses — reports
    // the SAME fact: the top-level counter moved by one. Deriving this from the
    // paragraph's own counter gives `shift: 0` on the six sub-clauses, which is
    // both wrong and uncollapsible.
    expect(cascades).toHaveLength(7);
    expect(cascades.every((c) => c?.shiftedIlvl === 0 && c.shift === 1)).toBe(true);
    expect(cascades[1]).toEqual({ from: '1.1', to: '2.1', shiftedIlvl: 0, shift: 1 });
  });

  it('COLLAPSES all seven into one row, because they are one fact', async () => {
    const { result } = await compareML(ORIG, REV);
    const props = properties(result.ops);

    expect(props).toHaveLength(7);
    expect(result.summary.renumberingCascades).toHaveLength(1);
    const cascade = result.summary.renumberingCascades[0]!;
    expect(cascade.source).toBe('numbering');
    expect(cascade.shift).toBe(1);
    expect(cascade.affectedCount).toBe(7);
    // Level 1 one-based — the level whose counter moved, not the sub-clauses' own.
    expect(cascade.tokenKind).toBe('Clause numbering (level 1)');
    // ADR-0008 holds: seven ops still exist, and every one is accounted for.
    expect(props.every((op) => op.collapsedInto === cascade.id)).toBe(true);
    expect(result.summary.total).toBe(8); // 7 collapsed members + the insert
  });
});

/**
 * The cascade HEADER must agree with the labels printed beneath it (§6.1, §8.3).
 *
 * §6.1 collapses a renumbering into one row that names a level and a shift, with
 * up to three example labels under it. Deriving that shift from the first
 * differing level of the counter PATH is right only when the level that moved is
 * one the template prints. Under `(%2)` — the own-level-only shape, and the shape
 * this suite's own shared fixture uses — splitting a clause list reported:
 *
 *     header: "Clause numbering (level 1) shifted (+1) — 3 paragraphs"
 *     rows:   (c) → (a)   (d) → (b)   (e) → (c)
 *
 * Three rows that each moved MINUS TWO, under a header claiming plus one. The
 * `+1` is the parent clause's counter, which `(%2)` renders nowhere. The two
 * templates disagree and only `%1.%2` had ever been probed — the same blind spot,
 * in the other direction, as the one that produced this plan's fourth defect.
 *
 * Both templates are asserted here, together, so neither can be fixed by breaking
 * the other.
 */
describe('the cascade header agrees with the labels beneath it', () => {
  const STYLES_CH = `<w:styles ${W}>
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:pPr/></w:style>
    <w:style w:type="paragraph" w:styleId="H"><w:pPr>
      <w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style>
    <w:style w:type="paragraph" w:styleId="S"><w:pPr>
      <w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style>
  </w:styles>`;

  /** Same list, same everything, ONE template token different. */
  const scheme = (subTemplate: string, headFmt = 'decimal', subExtra = '') => `<w:numbering ${W}>
    <w:abstractNum w:abstractNumId="0">
      <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${headFmt}"/><w:lvlText w:val="%1."/></w:lvl>
      <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/>${subExtra}
        <w:lvlText w:val="${subTemplate}"/></w:lvl>
    </w:abstractNum>
    <w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>
  </w:numbering>`;

  const SUBS = [
    'It shall maintain the insurances listed in the Schedule.',
    'It shall not amend the Disclosure Letter without consent.',
    'It shall keep the Records for six years.',
    'It shall notify the Purchaser of any Claim within ten days.',
    'It shall procure the release of each Encumbrance.',
  ].map((t) => para('S', t));
  const HEAD_A = para('H', 'The Vendor covenants with the Purchaser as follows.');
  const HEAD_B = para('H', 'The Purchaser covenants with the Vendor as follows.');

  /** A second top-level clause splits the list after the second sub-clause. */
  const ORIG_CH = body(HEAD_A, ...SUBS);
  const REV_CH = body(HEAD_A, ...SUBS.slice(0, 2), HEAD_B, ...SUBS.slice(2));

  const split = (subTemplate: string) => splitFmt(subTemplate, 'decimal');

  async function splitFmt(subTemplate: string, headFmt: string, subExtra = '') {
    return runOn(ORIG_CH, REV_CH, subTemplate, headFmt, subExtra);
  }

  async function runOn(
    origXml: string, revXml: string,
    subTemplate: string, headFmt: string, subExtra: string,
  ) {
    const read = (xml: string, policy: 'acceptAll' | 'rejectAll'): ParaRecord[] =>
      preprocessSide(xml, 'document', policy, config,
        buildNumberingContext(STYLES_CH, scheme(subTemplate, headFmt, subExtra))).paras;
    const original = read(origXml, config.flattenOriginal);
    const revised = read(revXml, config.flattenRevised);
    const result = await runComparePipeline({
      originalParas: original, revisedParas: revised, config,
      author: 'Mutatis', date: '2026-09-12T00:00:00Z',
    });
    return { original, revised, result };
  }

  it('names the level the label PRINTS when the template shows the own level only', async () => {
    const { original, revised, result } = await split('(%2)');

    // The parent counter moved 1 → 2 and `(%2)` prints none of it.
    expect(original.map((r) => r.numbering?.label ?? '·')).toEqual(
      ['1.', '(a)', '(b)', '(c)', '(d)', '(e)'],
    );
    expect(revised.map((r) => r.numbering?.label ?? '·')).toEqual(
      ['1.', '(a)', '(b)', '2.', '(a)', '(b)', '(c)'],
    );

    const [cascade] = result.summary.renumberingCascades;
    expect(cascade?.exampleRefs).toEqual(['(c) → (a)', '(d) → (b)', '(e) → (c)']);
    // `(c)` is 3 and `(a)` is 1 — the rows moved MINUS TWO, at level 2 (1-based).
    expect(cascade?.tokenKind).toBe('Clause numbering (level 2)');
    expect(cascade?.shift).toBe(-2);
    expect(cascade?.affectedCount).toBe(3);
  });

  it('skips a substituted ancestor whose format prints NOTHING', async () => {
    // `numFmt="none"` is a real Word value — a multilevel list that carries a
    // level for structure and shows no number for it. The template still
    // substitutes `%1`, so filter 1 (is it referenced?) lets it through; it
    // renders `''` for every counter value, so the `+1` at level 0 is invisible
    // and the rows move at level 1 only. Same contradiction as the test above,
    // one step further in.
    const { original, revised, result } = await splitFmt('%1.%2', 'none');

    expect(original.map((r) => r.numbering?.label ?? '·')).toEqual(
      ['.', '.a', '.b', '.c', '.d', '.e'],   // the head keeps its literal "."

    );
    expect(revised.map((r) => r.numbering?.label ?? '·')).toEqual(
      ['.', '.a', '.b', '.', '.a', '.b', '.c'],
    );

    const [cascade] = result.summary.renumberingCascades;
    expect(cascade?.exampleRefs).toEqual(['.c → .a', '.d → .b', '.e → .c']);
    expect(cascade?.tokenKind).toBe('Clause numbering (level 2)');
    expect(cascade?.shift).toBe(-2);
  });

  it('still names the ANCESTOR level when the template prints it', async () => {
    const { original, revised, result } = await split('%1.%2');

    expect(original.map((r) => r.numbering?.label ?? '·')).toEqual(
      ['1.', '1.a', '1.b', '1.c', '1.d', '1.e'],
    );
    expect(revised.map((r) => r.numbering?.label ?? '·')).toEqual(
      ['1.', '1.a', '1.b', '2.', '2.a', '2.b', '2.c'],
    );

    const [cascade] = result.summary.renumberingCascades;
    expect(cascade?.exampleRefs).toEqual(['1.c → 2.a', '1.d → 2.b', '1.e → 2.c']);
    // Here the `1` → `2` IS on the page, so level 1 (1-based) and +1 are what the
    // reader can check. Narrowing the scan to printed levels must not break this.
    expect(cascade?.tokenKind).toBe('Clause numbering (level 1)');
    expect(cascade?.shift).toBe(1);
    expect(cascade?.affectedCount).toBe(3);
  });

  it('counts a level `isLgl` makes printable, though its own format prints none', async () => {
    // `<w:isLgl/>` — "legal numbering", ECMA-376 §17.9 — forces every SUBSTITUTED
    // token to `decimal` regardless of the level's own `numFmt`. So a level that
    // prints nothing on its own line can still print on its descendants', and
    // `renderedLevels` has to fold `isLgl` in exactly as the key does. Both
    // filters read the own level's `isLgl`, not the substituted level's,
    // because that is what the renderer reads.
    //
    // The shape is a top-level clause inserted AHEAD of the list, so the
    // sub-clauses' own counters never move — only their parent's. That is what
    // makes this a guard: with the fold removed the scan sees level 1 alone,
    // finds no difference there, and returns no shift AT ALL, so the three rows
    // land uncollapsed. Under `numFmt="none"` the header's own label really is
    // just `.`, and the counter behind it is still the one the reader sees move
    // in `1.1` → `2.1`.
    const HEAD_C = para('H', 'The Vendor further covenants with the Purchaser.');
    const { original, revised, result } = await runOn(
      body(HEAD_A, ...SUBS.slice(0, 3)),
      body(HEAD_A, HEAD_C, ...SUBS.slice(0, 3)),
      '%1.%2', 'none', '<w:isLgl/>',
    );

    expect(original.map((r) => r.numbering?.label ?? '·')).toEqual(
      ['.', '1.1', '1.2', '1.3'],
    );
    expect(revised.map((r) => r.numbering?.label ?? '·')).toEqual(
      ['.', '.', '2.1', '2.2', '2.3'],
    );
    // The sub-clauses render level 0 — `none` notwithstanding — and the head,
    // whose own level carries no `isLgl`, renders nothing at all.
    expect(revised[0]!.numbering?.renders).toEqual([]);
    expect(revised[2]!.numbering?.renders).toEqual([0, 1]);

    // One row, not three: §6.1's collapse is the whole reason this matters.
    const [cascade] = result.summary.renumberingCascades;
    expect(cascade?.tokenKind).toBe('Clause numbering (level 1)');
    expect(cascade?.shift).toBe(1);
    expect(cascade?.affectedCount).toBe(3);
    expect(cascade?.exampleRefs).toEqual(['1.1 → 2.1', '1.2 → 2.2', '1.3 → 2.3']);
    const moved = properties(result.ops);
    expect(moved).toHaveLength(3);
    for (const op of moved) expect(op.collapsedInto).toBe(cascade?.id);
  });
});

/**
 * The cascade header's unit noun comes from the Config Layer (§7.4, ADR-0002).
 *
 * `derivePropertyMember` builds "Clause numbering (level 2)" from
 * `config.primaryUnit`. It once read a hardcoded `'Clause'`, and then
 * `referenceTokens[0]` — a DETECTION set whose order carries no meaning, which
 * silently renamed the heading to "paragraph numbering" for a profile that
 * listed its nouns another way. Both were fixed without a test, so the fix could
 * be reverted in full and the whole suite stayed green: measured at 206 files,
 * 2372 passed, with the Config read replaced by the literal again.
 *
 * The jurisdiction invariant is the reason this gets a guard
 * rather than a comment: the noun is jurisdiction-owned, and a hardcoded one is
 * exactly the class of defect ADR-0002 exists to prevent — invisible here only
 * because `Clause` happens to be the right answer for both shipped profiles.
 */
describe('the cascade header names the jurisdiction\'s unit, from Config', () => {
  /**
   * `xref.units.primaryBody` is typed as the literal `'Clause'` — the two shipped
   * profiles agree and the type says so. A guard that cannot vary the value
   * cannot fail, so the profile is widened here rather than asserted against
   * itself. This is the only knob that makes the derivation observable.
   */
  const articleProfile = {
    ...SG_PROFILE,
    xref: {
      ...SG_PROFILE.xref,
      units: { ...SG_PROFILE.xref.units, primaryBody: 'Article' },
    },
  } as unknown as typeof SG_PROFILE;

  it('derives the unit from the profile\'s xref.units.primaryBody', () => {
    expect(getCompareConfig(SG_PROFILE).primaryUnit).toBe(SG_PROFILE.xref.units.primaryBody);
    expect(getCompareConfig(articleProfile).primaryUnit).toBe('Article');
  });

  it('lets a profile override Compare\'s noun without moving the xref layer\'s', () => {
    const overridden = {
      ...SG_PROFILE,
      compare: { ...getCompareConfig(SG_PROFILE), primaryUnit: 'Paragraph' },
    };
    expect(getCompareConfig(overridden).primaryUnit).toBe('Paragraph');
    expect(overridden.xref.units.primaryBody).toBe('Clause');
  });

  it('puts that noun in the collapsed cascade row, end to end', async () => {
    const articleConfig = getCompareConfig(articleProfile);
    const STYLES_U = `<w:styles ${W}>
      <w:style w:type="paragraph" w:styleId="H"><w:pPr><w:numPr>
        <w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style>
      <w:style w:type="paragraph" w:styleId="S"><w:pPr><w:numPr>
        <w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style>
    </w:styles>`;
    const NUM_U = `<w:numbering ${W}><w:abstractNum w:abstractNumId="0">
      <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
      <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl>
      </w:abstractNum><w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;

    const subs = [
      'It shall maintain the insurances listed in the Schedule.',
      'It shall not amend the Disclosure Letter without consent.',
      'It shall keep the Records for six years.',
      'It shall notify the Purchaser of any Claim within ten days.',
    ].map((t) => para('S', t));
    const head = para('H', 'The Vendor covenants with the Purchaser as follows.');
    const origXml = body(head, ...subs);
    const revXml = body(para('H', 'A newly inserted division precedes the covenant.'), head, ...subs);

    const read = (xml: string, policy: 'acceptAll' | 'rejectAll'): ParaRecord[] =>
      preprocessSide(xml, 'document', policy, articleConfig,
        buildNumberingContext(STYLES_U, NUM_U)).paras;
    const result = await runComparePipeline({
      originalParas: read(origXml, articleConfig.flattenOriginal),
      revisedParas: read(revXml, articleConfig.flattenRevised),
      config: articleConfig,
      author: 'Mutatis',
      date: '2026-09-12T00:00:00Z',
    });

    const [cascade] = result.summary.renumberingCascades;
    // Five, not four: the covenant HEAD moves `1.` → `2.` at level 0 too, which
    // is the same (level, shift) fact as its four sub-clauses' `1.n` → `2.n`.
    expect(cascade?.affectedCount).toBe(5);
    // Hardcode `'Clause'` back into `derivePropertyMember` and this is the
    // assertion that notices.
    expect(cascade?.tokenKind).toBe('Article numbering (level 1)');
  });
});
/**
 * ADR-0042 § Accepted limits, the "restart numbering" limit — both halves.
 *
 * That limit is the only one of the four whose *reported* half rested entirely
 * on a figure in the ADR with no test behind it, and it is the half a reader is
 * most likely to doubt: the key deliberately excludes `numId`, so "a clause
 * moved to another list" looks like exactly the thing the key cannot see.
 *
 * It sees it, and by the route ADR-0042 says it does — a restarted instance has
 * a different key AND different labels, because `<w:startOverride>` reaches both:
 * `resolveLevel` folds it into the level's `start` (so it is in the key), and
 * `startOverrideFor` restarts the shared counter at it (so it is in the labels).
 * The second half is a later review's: counters belong to the abstract definition, so a
 * fresh instance no longer restarts by itself and the override is what does it. What is genuinely unreported is only the
 * residue: a move between two instances that render identically, where nothing a
 * reader could see has changed. Both are asserted here, together, because a test
 * for the first alone would leave the limit's scope unpinned in the direction
 * that matters (ADR-0008: a detection limit must be a limit, not a hole).
 */
describe('a clause moved onto a RESTARTED list instance', () => {
  const STYLES_OV = `<w:styles ${W}>
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:pPr/></w:style>
    <w:style w:type="paragraph" w:styleId="H"><w:pPr>
      <w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style>
    <w:style w:type="paragraph" w:styleId="S"><w:pPr>
      <w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style>
    <w:style w:type="paragraph" w:styleId="S7"><w:pPr>
      <w:numPr><w:ilvl w:val="1"/><w:numId w:val="7"/></w:numPr></w:pPr></w:style>
  </w:styles>`;

  /** Two instances over ONE abstract definition; `numId` 7 carries the override. */
  const scheme = (override: string) => `<w:numbering ${W}>
    <w:abstractNum w:abstractNumId="0">
      <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
      <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl>
    </w:abstractNum>
    <w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>
    <w:num w:numId="7"><w:abstractNumId w:val="0"/>${override}</w:num>
  </w:numbering>`;

  const restartedAt = (start: string) =>
    `<w:lvlOverride w:ilvl="1"><w:startOverride w:val="${start}"/></w:lvlOverride>`;

  const SUBS_OV = [
    'It shall maintain the insurances listed in the Schedule.',
    'It shall not amend the Disclosure Letter without consent.',
    'It shall keep the Records for six years.',
    'It shall notify the Purchaser of any Claim within ten days.',
    'It shall procure the release of each Encumbrance.',
    'It shall pay the Retention into the Escrow Account.',
  ];

  /** The paragraph's own `<w:numPr>` overrides the style's instance, level intact. */
  const onInstance = (text: string, numId?: string) =>
    `<w:p><w:pPr><w:pStyle w:val="S"/>` +
    (numId === undefined ? '' : `<w:numPr><w:ilvl w:val="1"/><w:numId w:val="${numId}"/></w:numPr>`) +
    `</w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

  /** Same six sub-clauses; in the revised side the last three move to `numId` 7. */
  async function moveLastThree(override: string) {
    const head = para('H', 'The Vendor covenants with the Purchaser as follows.');
    const origXml = body(head, ...SUBS_OV.map((t) => onInstance(t)));
    const revXml = body(
      head,
      ...SUBS_OV.slice(0, 3).map((t) => onInstance(t)),
      ...SUBS_OV.slice(3).map((t) => onInstance(t, '7')),
    );
    const read = (xml: string, policy: 'acceptAll' | 'rejectAll'): ParaRecord[] =>
      preprocessSide(xml, 'document', policy, config,
        buildNumberingContext(STYLES_OV, scheme(override))).paras;
    const original = read(origXml, config.flattenOriginal);
    const revised = read(revXml, config.flattenRevised);
    const result = await runComparePipeline({
      originalParas: original, revisedParas: revised, config,
      author: 'Mutatis', date: '2026-09-12T00:00:00Z',
    });
    return { original, revised, result };
  }

  it('IS reported — the figure ADR-0042 records, reproduced', async () => {
    // Word's "restart numbering at 1", which is what the button writes.
    //
    // **This comment used to say the opposite, and a later review turned
    // it around.** It read: *"not evidence for the `startOverride` fold —
    // `render-list-numbers.ts` keys counters per `numId`, so the fresh instance
    // restarts at 1 with or without the override."* That per-`numId` keying was
    // itself the defect (it invented a restart Word does not do, and hid a real
    // one — see the module's fidelity bounds). Counters now belong to the
    // abstract DEFINITION, so without the override these three would continue
    // `1.4 1.5 1.6` and this test would fail. The restart is now load-bearing
    // here, and `startOverrideFor` is what performs it — the fold in
    // `resolveLevel` only covers a counter that has not been initialised yet.
    //
    // A guard's premise goes stale when the thing under it changes: running the
    // suite after that change would have said this test still passes, and could
    // not have said it now tests something else.
    const { original, revised, result } = await moveLastThree(restartedAt('1'));

    expect(original.map((r) => r.numbering?.label ?? '·')).toEqual(
      ['1.', '1.1', '1.2', '1.3', '1.4', '1.5', '1.6'],
    );
    expect(revised.map((r) => r.numbering?.label ?? '·')).toEqual(
      ['1.', '1.1', '1.2', '1.3', '1.1', '1.2', '1.3'],
    );

    const moved = properties(result.ops);
    expect(moved).toHaveLength(3);
    // The paragraphs themselves did not move; the list under them restarted —
    // so `crossRefRenumber`, which §8.2’s filter does NOT hide, and one collapsed row.
    for (const op of moved) expect(op.classification).toBe('crossRefRenumber');
    const [cascade] = result.summary.renumberingCascades;
    expect(cascade?.tokenKind).toBe('Clause numbering (level 2)');
    expect(cascade?.shift).toBe(-3);
    expect(cascade?.affectedCount).toBe(3);
    expect(cascade?.exampleRefs).toEqual(['1.4 → 1.1', '1.5 → 1.2', '1.6 → 1.3']);
    for (const op of moved) expect(op.collapsedInto).toBe(cascade?.id);
  });

  it('reads the OVERRIDDEN start, not the abstract definition\'s', async () => {
    // The fold itself: `<w:startOverride w:val="9"/>`, a value the abstract
    // definition does not carry and the counter would never reach on its own.
    // Ignore it and the three clauses read `1.1`–`1.3` — the right *kind* of
    // change under the wrong numbers, which is worse than silence because the
    // cascade row tells the reader a shift that did not happen.
    const { original, revised, result } = await moveLastThree(restartedAt('9'));

    expect(original.map((r) => r.numbering?.label ?? '·')).toEqual(
      ['1.', '1.1', '1.2', '1.3', '1.4', '1.5', '1.6'],
    );
    expect(revised.map((r) => r.numbering?.label ?? '·')).toEqual(
      ['1.', '1.1', '1.2', '1.3', '1.9', '1.10', '1.11'],
    );

    // And the classification differs from the test above, which is the fold
    // showing up a second time: `start` is part of the comparison key, so an
    // override to a value the definition does not carry is a list REDEFINITION
    // (`formatting`) rather than a paragraph whose place in the count moved
    // (`crossRefRenumber`). §6.1 collapses only the latter, so there is no
    // cascade row here and there should not be one.
    const moved = properties(result.ops);
    expect(moved).toHaveLength(3);
    for (const op of moved) {
      expect(op.propertyChange?.numberingIdentity).toBe(true);
      expect(op.classification).toBe('formatting');
    }
    expect(result.summary.renumberingCascades).toHaveLength(0);
  });

  it('carries no `cascade` when the paragraph ITSELF also changed', async () => {
    // `PropertyChange.cascade` is documented — in `model.ts`, in §4.2 and in
    // ADR-0042 — as set ONLY for a pure cascade: the paragraph did not change,
    // its number moved because a sibling's did. Relax the condition in
    // `comparePropertyPair` to the label alone and the whole suite stays green,
    // because every other case in it has paths of different lengths and
    // `counterShift` returns null before the contract is tested.
    //
    // This is the case that reaches it: a sub-clause INSERTED above them moves
    // their level-1 counter by +1, so the counter path is the same LENGTH and
    // differs, while their own `w:pStyle` moved too. A cascade row here would
    // say *"3 clauses renumbered because a sibling moved"* about three
    // paragraphs that were themselves restyled — §6.1 collapsing bystanders that
    // are not bystanders.
    //
    // **The shape changed in a later review and the reason is worth keeping.** It used
    // to move the three paragraphs onto a second `w:num` over the same abstract
    // definition, with no override, and read their labels as `1.1 1.2 1.3` —
    // numbers Word never renders, because it continues one sequence across
    // instances. That fabricated restart was what made the label "move" here.
    // With counters keyed by the definition the labels no longer move and this
    // test no longer reached its own subject, so it now moves them with a real
    // insertion. `S7` still supplies the style change: it names a different
    // instance of the SAME definition, which is invisible to the key by design.
    const head = para('H', 'The Vendor covenants with the Purchaser as follows.');
    const onStyle = (styleId: string, text: string) =>
      `<w:p><w:pPr><w:pStyle w:val="${styleId}"/></w:pPr>` +
      `<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
    const origXml = body(head, ...SUBS_OV.map((t) => onStyle('S', t)));
    const revXml = body(
      head,
      ...SUBS_OV.slice(0, 3).map((t) => onStyle('S', t)),
      onStyle('S', 'It shall deliver the Completion Accounts when due.'),
      ...SUBS_OV.slice(3).map((t) => onStyle('S7', t)),
    );
    const read = (xml: string, policy: 'acceptAll' | 'rejectAll'): ParaRecord[] =>
      preprocessSide(xml, 'document', policy, config,
        buildNumberingContext(STYLES_OV, scheme(''))).paras;
    const original = read(origXml, config.flattenOriginal);
    const revised = read(revXml, config.flattenRevised);
    const result = await runComparePipeline({
      originalParas: original, revisedParas: revised, config,
      author: 'Mutatis', date: '2026-09-12T00:00:00Z',
    });

    const moved = properties(result.ops);
    expect(moved).toHaveLength(3);
    for (const op of moved) {
      const pc = op.propertyChange!;
      // All three preconditions the relaxed condition would ignore.
      expect(pc.styleId).toBe(true);
      expect(pc.numberingIdentity).toBe(false);
      expect(pc.numberLabel).toBe(true);
      expect(original[4]!.numbering?.path).toHaveLength(
        revised[4]!.numbering?.path?.length ?? -1,
      );
      expect(pc.cascade).toBeUndefined();
    }
    expect(result.summary.renumberingCascades).toHaveLength(0);
  });

  it('is NOT reported when the two instances render identically — the recorded limit', async () => {
    // The residue the limit actually names, and it is narrower than "a clause
    // moved lists": the move has to leave every label where it was. ONE
    // sub-clause moved to a second instance of the same definition does that —
    // it is the only member of the level-1 sequence either way, so the label is
    // `1.1` on both sides and nothing a reader could see has happened.
    //
    // The key excludes `numId` on purpose (Word reassigns instance ids on a
    // Save-As while every clause keeps its number), so the two sides are
    // identical on every field it carries. A detection limit, not a
    // suppression: nothing is found and then hidden.
    //
    // **A later fix narrowed this residue, and that is the right direction.** The
    // renderer used to key counters per `numId`, so a move to a second instance
    // restarted the numbers and SIX moved that way were "reported" — on labels
    // Word never renders. Counters now belong to the abstract definition, so a
    // move between two instances of one definition changes nothing unless an
    // override or the sequence itself moves, and that is exactly what this test
    // now says. The test above reports because `numId` 7 carries a
    // `<w:startOverride>`, which is a real restart and not an artefact.
    const head = para('H', 'The Vendor covenants with the Purchaser as follows.');
    const read = (xml: string, policy: 'acceptAll' | 'rejectAll'): ParaRecord[] =>
      preprocessSide(xml, 'document', policy, config,
        buildNumberingContext(STYLES_OV, scheme(''))).paras;
    const original = read(body(head, onInstance(SUBS_OV[0]!)), config.flattenOriginal);
    const revised = read(body(head, onInstance(SUBS_OV[0]!, '7')), config.flattenRevised);
    const result = await runComparePipeline({
      originalParas: original, revisedParas: revised, config,
      author: 'Mutatis', date: '2026-09-12T00:00:00Z',
    });

    expect(original.map((r) => r.numbering?.label)).toEqual(['1.', '1.1']);
    expect(revised.map((r) => r.numbering?.label)).toEqual(['1.', '1.1']);
    expect(properties(result.ops)).toHaveLength(0);
    expect(result.summary.total).toBe(0);
  });
});
/**
 * Three levels — `Clause 1.1.1`, which is ordinary SG/E&W drafting and which
 * nothing in this suite probed.
 *
 * Every numbering defect that slipped past review was a case the tests
 * could not express: a single-level list, then an own-level-only template, then
 * a two-level one. Two levels is enough to make an ancestor's counter reach a
 * descendant's label; it is not enough to put a level BETWEEN the one that moved
 * and the one reading it, which is where a scan that stops at the first
 * difference, or a key that carries the wrong span, would show.
 */
describe('a three-level list, which two levels cannot stand in for', () => {
  const STYLES_3 = `<w:styles ${W}>
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:pPr/></w:style>
    <w:style w:type="paragraph" w:styleId="L0"><w:pPr>
      <w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style>
    <w:style w:type="paragraph" w:styleId="L1"><w:pPr>
      <w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style>
    <w:style w:type="paragraph" w:styleId="L2"><w:pPr>
      <w:numPr><w:ilvl w:val="2"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style>
  </w:styles>`;

  const NUM_3 = `<w:numbering ${W}><w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl>
    <w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2.%3"/></w:lvl>
  </w:abstractNum><w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;

  it('shifts a grandchild from the level TWO above it, and collapses to one row', async () => {
    const head = para('L0', 'The Vendor covenants with the Purchaser as follows.');
    const kids = [
      para('L1', 'It shall maintain the insurances listed in the Schedule.'),
      para('L2', 'It shall renew each policy before it lapses.'),
      para('L2', 'It shall name the Purchaser as a co-insured.'),
      para('L1', 'It shall keep the Records for six years.'),
    ];
    const origXml = body(head, ...kids);
    const revXml = body(
      para('L0', 'A newly inserted division precedes the covenant.'), head, ...kids,
    );

    const read = (xml: string, policy: 'acceptAll' | 'rejectAll'): ParaRecord[] =>
      preprocessSide(xml, 'document', policy, config,
        buildNumberingContext(STYLES_3, NUM_3)).paras;
    const original = read(origXml, config.flattenOriginal);
    const revised = read(revXml, config.flattenRevised);
    const result = await runComparePipeline({
      originalParas: original, revisedParas: revised, config,
      author: 'Mutatis', date: '2026-09-12T00:00:00Z',
    });

    expect(original.map((r) => r.numbering?.label ?? '·')).toEqual(
      ['1.', '1.1', '1.1.1', '1.1.2', '1.2'],
    );
    expect(revised.map((r) => r.numbering?.label ?? '·')).toEqual(
      ['1.', '2.', '2.1', '2.1.1', '2.1.2', '2.2'],
    );

    // The level-2 paragraphs read all three levels; each level reads only as
    // far as its own. This is the field a scan of the counter path uses to stop
    // at the levels the label actually shows.
    expect(original.map((r) => r.numbering?.renders ?? null)).toEqual(
      [[0], [0, 1], [0, 1, 2], [0, 1, 2], [0, 1]],
    );

    // One row, not five, and it names the level that MOVED — 0 (1-based: 1) —
    // not each paragraph's own. The grandchildren's own counters never moved.
    const [cascade] = result.summary.renumberingCascades;
    expect(cascade?.tokenKind).toBe('Clause numbering (level 1)');
    expect(cascade?.shift).toBe(1);
    expect(cascade?.affectedCount).toBe(5);
    expect(cascade?.exampleRefs).toEqual(['1. → 2.', '1.1 → 2.1', '1.1.1 → 2.1.1']);

    const moved = properties(result.ops);
    expect(moved).toHaveLength(5);
    for (const op of moved) {
      expect(op.propertyChange?.numberingIdentity).toBe(false);
      expect(op.collapsedInto).toBe(cascade?.id);
    }
  });
});

/**
 * The key's field separator is load-bearing, and what a collision would cost.
 *
 * `numberingKey` joins every field with `U+0001` — a C0 control that is illegal
 * in XML 1.0 content, so a CONFORMING `numbering.xml` cannot put it inside a
 * `lvlText` or a `numFmt` and no two different definitions can land on one key.
 * Nothing held that: empty the constant and the whole suite stayed green.
 *
 * A collision cannot be built from conforming values — which is the point of
 * choosing that byte — so the probes below use `numFmt` values no producer would
 * write (`decimal1`, `fooX1`). That makes them tests of the ENCODING, not claims
 * about any producer's output, and they are labelled as such rather than dressed up as
 * corpus cases.
 *
 * **Both probes sit at level 1 and collide an ANCESTOR level's fields, and that
 * is not cosmetic.** They used to collide the paragraph's OWN `numFmt` against
 * its own `start`, at level 0. A later fix added the own level's
 * DECLARED `numFmt` as a trailing field; a difference in the last field has
 * nothing after it to absorb the boundary; and both probes stopped colliding on
 * the spot — still passing, still asserting, testing nothing. Running the suite
 * proved they passed. It could not say they had lost their premise. An ancestor
 * block ends at `lvlRestart`, so `numFmt`+`start` there still concatenates the
 * same, with the own level identical on both sides:
 *
 *     1 . %1.%2 . 0 . present . fooX  . 11 . 0 . 1 . present . decimal . 1 . 1 . decimal
 *     1 . %1.%2 . 0 . present . fooX1 . 1  . 0 . 1 . present . decimal . 1 . 1 . decimal
 *
 * When a change widens this key, re-derive every existing guard's PREMISE, not
 * just its pass/fail.
 *
 * There is no third probe for the level prefix, and deliberately so: the key is
 * `${ilvl}` then the fields, so a separator appearing only there can decide a
 * comparison only when one side's `ilvl` decimal string is a proper prefix of
 * the other's — impossible while both are single digits, which Word's nine
 * levels always are. It was a second constant that no ablation could turn red.
 * The level is now the first ordinary field behind the one separator these two
 * tests pin.
 */
describe("the comparison key's field separator", () => {
  const STYLES_SEP = `<w:styles ${W}>
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:pPr/></w:style>
    <w:style w:type="paragraph" w:styleId="H"><w:pPr>
      <w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style>
  </w:styles>`;

  /** Level 0 is the one that varies; the level the paragraph sits at never does. */
  const scheme = (numFmt: string, start: string) => `<w:numbering ${W}>
    <w:abstractNum w:abstractNumId="0">
      <w:lvl w:ilvl="0"><w:start w:val="${start}"/><w:numFmt w:val="${numFmt}"/>
        <w:lvlText w:val="%1."/></w:lvl>
      <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/>
        <w:lvlText w:val="%1.%2"/></w:lvl>
    </w:abstractNum>
    <w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>
  </w:numbering>`;

  const ONE = body(para('H', 'The Vendor shall comply with this Clause in all respects.'));

  const read = (numXml: string, policy: 'acceptAll' | 'rejectAll'): ParaRecord[] =>
    preprocessSide(ONE, 'document', policy, config, buildNumberingContext(STYLES_SEP, numXml)).paras;

  it('keeps two definitions apart that would otherwise concatenate the same', async () => {
    // `decimal` + start `11` and `decimal1` + start `1` differ field by field and
    // are IDENTICAL once the boundaries are removed. With the separator the keys
    // differ and the op says the LIST was redefined. Empty it and the key reads
    // identical, the op SURVIVES — `numberLabel` is compared from the two
    // RENDERED labels and never from the key — but it now says *"its number moved
    // because a sibling did"*, and §8.2 hides one of those two by default.
    //
    // Here, where the two labels differ, a collision misattributes and cannot
    // hide. That is only half the story, and the half that is easy to
    // over-generalise: the test below is the other half, where the labels are
    // identical and a collision drops the op outright.
    const original = read(scheme('decimal', '11'), config.flattenOriginal);
    const revised = read(scheme('decimal1', '1'), config.flattenRevised);
    const result = await runComparePipeline({
      originalParas: original, revisedParas: revised, config,
      author: 'Mutatis', date: '2026-09-12T00:00:00Z',
    });

    expect(original[0]!.numbering?.key).not.toBe(revised[0]!.numbering?.key);
    expect(original[0]!.numbering?.label).toBe('11.1');
    expect(revised[0]!.numbering?.label).toBe('.1');

    const [op] = properties(result.ops);
    expect(op?.propertyChange?.numberingIdentity).toBe(true);
    expect(op?.classification).toBe('formatting');
    expect(result.summary.byCategory.crossRefRenumber).toBe(0);
  });

  it('is the ONLY detector when a definition change moves no label', async () => {
    // The case that falsifies the unqualified form of this plan's central safety
    // argument. That argument — *"narrowing can misattribute a change, it cannot
    // hide one"* — was written into ADR-0042, into `numberingKey`'s doc comment
    // and into this suite, and it is true only of a label that MOVED. Here
    // neither label moves: both level-0 `numFmt` values are ones `formatInt`
    // cannot express, so the `%1` token renders empty on both sides and the label
    // is `.1` either way. The definition changed; the page did not.
    //
    // For that class the key is the only detector there is. With the separator
    // the fields stay apart and the change is reported — ADR-0042's *"err towards
    // reporting"*. Empty it and the two keys collide, the label comparison has
    // nothing to add, and the op disappears: `summary.total` 1 to 0. A change
    // found and then not reported, which ADR-0008 forbids.
    //
    // There is no bridging rule that makes the narrowing automatically safe; two
    // were written into this plan and measurement falsified both. The second
    // claimed the key's span is a superset of everything that can move this
    // paragraph's label, which ADR-0042 § Accepted limits TWO has always
    // contradicted. Two facts, no rule joining them: ON A PAIR THIS PASS
    // EXAMINES a label that MOVED is always reported whatever the key carries,
    // and a change that moves no label — this one — is reported only if the key
    // carries the field. (That qualifier is an earlier review's, falsified a third time by
    // Accepted limit FOUR: a clause whose words were edited AND whose number
    // moved reports the text change and nothing about the number. This comment
    // was the seventh home of the sentence and a review found it unqualified.)
    const original = read(scheme('fooX', '11'), config.flattenOriginal);
    const revised = read(scheme('fooX1', '1'), config.flattenRevised);
    const result = await runComparePipeline({
      originalParas: original, revisedParas: revised, config,
      author: 'Mutatis', date: '2026-09-12T00:00:00Z',
    });

    // Neither side renders the level-0 counter — the labels are provably equal.
    expect(original[0]!.numbering?.label).toBe('.1');
    expect(revised[0]!.numbering?.label).toBe('.1');
    expect(original[0]!.numbering?.key).not.toBe(revised[0]!.numbering?.key);

    const [op] = properties(result.ops);
    expect(op?.propertyChange?.numberingIdentity).toBe(true);
    expect(op?.propertyChange?.numberLabel).toBe(false); // nothing visible moved
    expect(result.summary.total).toBe(1);
  });
});

/**
 * A list built through a Word LIST STYLE — the third ADR-0008 suppression this
 * area produced, found in review.
 *
 * When numbering is attached to a list style, Word writes the levels on a second
 * `<w:abstractNum>` carrying `<w:styleLink>` and leaves the one the `<w:num>`
 * points at holding nothing but `<w:numStyleLink>` (ECMA-376 §17.9.21/§17.9.28).
 * `resolveLevel` did not follow that, so every level resolved to `undefined`,
 * every label was `null`, and the comparison key was the bare level on both
 * sides — no identity change, no label change, no op. Measured before the fix:
 *
 *     redefinition decimal -> lowerRoman   levels on the abstractNum  total 3
 *                                          behind a list style        total 0
 *     one clause inserted above            levels on the abstractNum  total 4, 1 cascade
 *                                          behind a list style        total 1, 0 cascades
 *
 * It survived 23 passes because a re-LEVELLING is still caught — `ilvl` is in the
 * key whether or not the level resolves — and every probe in this file re-levels.
 * It is the same shape as an earlier review's: a bound recorded as the RENDERER's
 * (*"style-linked numbering … handled upstream/deferred"*) is not confined to the
 * renderer while the comparison key resolves through the same module.
 */
describe('a list built through a Word LIST STYLE', () => {
  const STYLES_LS = `<w:styles ${W}>
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:pPr/></w:style>
    <w:style w:type="paragraph" w:styleId="Clause1"><w:pPr>
      <w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style>
    <w:style w:type="numbering" w:styleId="HouseList"><w:pPr>
      <w:numPr><w:numId w:val="9"/></w:numPr></w:pPr></w:style>
  </w:styles>`;

  /** The levels live on abstractNum 1, behind the style; abstractNum 0 delegates. */
  const viaStyle = (numFmt: string) => `<w:numbering ${W}>
    <w:abstractNum w:abstractNumId="0"><w:numStyleLink w:val="HouseList"/></w:abstractNum>
    <w:abstractNum w:abstractNumId="1"><w:styleLink w:val="HouseList"/>
      <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${numFmt}"/>
        <w:lvlText w:val="%1."/></w:lvl>
    </w:abstractNum>
    <w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>
    <w:num w:numId="9"><w:abstractNumId w:val="1"/></w:num>
  </w:numbering>`;

  const THREE = body(para('Clause1', C1), para('Clause1', C2), para('Clause1', C3));
  const FOUR = body(
    para('Clause1', 'The Purchaser shall pay the Consideration on Completion.'),
    para('Clause1', C1), para('Clause1', C2), para('Clause1', C3));

  const read = (doc: string, numXml: string, policy: 'acceptAll' | 'rejectAll'): ParaRecord[] =>
    preprocessSide(doc, 'document', policy, config, buildNumberingContext(STYLES_LS, numXml)).paras;

  const run = async (oDoc: string, oNum: string, rDoc: string, rNum: string) => {
    const original = read(oDoc, oNum, config.flattenOriginal);
    const revised = read(rDoc, rNum, config.flattenRevised);
    const result = await runComparePipeline({
      originalParas: original, revisedParas: revised, config,
      author: 'Mutatis', date: '2026-09-12T00:00:00Z',
    });
    return { original, revised, result };
  };

  it('REPORTS a redefinition of the list the style owns', async () => {
    const { original, revised, result } = await run(
      THREE, viaStyle('decimal'), THREE, viaStyle('lowerRoman'));

    // The labels resolve at all — which is the half that was missing.
    expect(original.map((p) => p.numbering?.label)).toEqual(['1.', '2.', '3.']);
    expect(revised.map((p) => p.numbering?.label)).toEqual(['i.', 'ii.', 'iii.']);

    const props = properties(result.ops);
    expect(props).toHaveLength(3);
    for (const op of props) {
      expect(op.propertyChange?.numberingIdentity).toBe(true);
      expect(op.classification).toBe('formatting');
    }
    expect(result.summary.total).toBe(3);
  });

  it('collapses the cascade when a clause is inserted above such a list', async () => {
    const { result } = await run(THREE, viaStyle('decimal'), FOUR, viaStyle('decimal'));

    // One insertion, three clauses whose numbers moved because of it, collapsed
    // into one §6.1 row. Before the fix this was `total 1` — the insertion alone,
    // with the three renumbered clauses reported by nothing at all.
    expect(result.summary.total).toBe(4);
    const [cascade] = result.summary.renumberingCascades;
    expect(cascade?.affectedCount).toBe(3);
    expect(cascade?.shift).toBe(1);
    expect(properties(result.ops)).toHaveLength(3);
  });
});

/**
 * Two `<w:num>` over ONE `<w:abstractNum>` — the fourth ADR-0008 suppression of
 * this area, found in review, and the next item on the
 * same list of "renderer bounds" that produced the second and third.
 *
 * Counters belong to the abstract DEFINITION, not to the instance. That is not a
 * claim about Word that this repo cannot test: `<w:startOverride>` exists to
 * restart a level for one instance, and it would be a no-op if every fresh
 * `<w:num>` restarted by itself. Word writes a second instance WITH an override
 * when the user asks to restart and WITHOUT one when the user asks to continue.
 *
 * While the renderer keyed counters per `numId`, the same bound broke Compare in
 * both directions, measured end to end:
 *
 *   - SILENCE. A clause inserted into the first of two groups sharing one
 *     definition moves the second group's numbers. Reported: NOTHING —
 *     `summary.total` 1 (the insert alone) against 3 for the identical edit on a
 *     single instance.
 *   - INVENTION. Re-pointing three sub-clauses at a second instance of the same
 *     definition reported a `-3` renumbering cascade — *"three clauses
 *     renumbered"* over numbers Word never renders.
 *
 * The second is the worse of the two, and the suite already said so about a
 * neighbouring case: *"the right kind of change under the wrong numbers, which is
 * worse than silence"*.
 */
describe('two list instances over one definition', () => {
  const STYLES_SH = `<w:styles ${W}>
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:pPr/></w:style>
    <w:style w:type="paragraph" w:styleId="GroupA"><w:pPr>
      <w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style>
    <w:style w:type="paragraph" w:styleId="GroupB"><w:pPr>
      <w:numPr><w:ilvl w:val="0"/><w:numId w:val="4"/></w:numPr></w:pPr></w:style>
  </w:styles>`;

  /** `numId` 3 and `numId` 4 both point at abstract 0. Neither overrides anything. */
  const SHARED = `<w:numbering ${W}>
    <w:abstractNum w:abstractNumId="0">
      <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>
        <w:lvlText w:val="%1."/></w:lvl>
    </w:abstractNum>
    <w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>
    <w:num w:numId="4"><w:abstractNumId w:val="0"/></w:num>
  </w:numbering>`;

  const A1 = 'The Vendor shall procure that the Company complies with this Clause.';
  const A2 = 'Settlement shall be made in cleared funds within ten Business Days.';
  const B1 = 'Each Party shall bear its own costs of the negotiation.';
  const B2 = 'Notices shall be given in writing to the address last notified.';
  const INSERTED = 'The Purchaser shall pay the Consideration on Completion.';

  const read = (doc: string, policy: 'acceptAll' | 'rejectAll'): ParaRecord[] =>
    preprocessSide(doc, 'document', policy, config,
      buildNumberingContext(STYLES_SH, SHARED)).paras;

  const run = async (o: string, r: string) => {
    const original = read(o, config.flattenOriginal);
    const revised = read(r, config.flattenRevised);
    const result = await runComparePipeline({
      originalParas: original, revisedParas: revised, config,
      author: 'Mutatis', date: '2026-09-12T00:00:00Z',
    });
    return { original, revised, result };
  };

  it('continues ONE sequence, so an insert in the first group moves the second', async () => {
    const { original, revised, result } = await run(
      body(para('GroupA', A1), para('GroupA', A2), para('GroupB', B1), para('GroupB', B2)),
      body(para('GroupA', A1), para('GroupA', A2), para('GroupA', INSERTED),
        para('GroupB', B1), para('GroupB', B2)),
    );

    // The second group counts on from the first, which is the half that was
    // missing: per-`numId` counters rendered it `1. 2.` on both sides.
    expect(original.map((r2) => r2.numbering?.label)).toEqual(['1.', '2.', '3.', '4.']);
    expect(revised.map((r2) => r2.numbering?.label)).toEqual(['1.', '2.', '3.', '4.', '5.']);

    // The insert, plus the two clauses whose numbers it moved.
    expect(result.summary.total).toBe(3);
    const moved = properties(result.ops);
    expect(moved).toHaveLength(2);
    for (const op of moved) {
      expect(op.classification).toBe('crossRefRenumber');
      expect(op.propertyChange?.numberingIdentity).toBe(false);
      expect(op.propertyChange?.numberLabel).toBe(true);
    }
  });

  it('is SILENT when a clause moves between them, because nothing renders differently', async () => {
    // The other direction, and the one that used to INVENT a change: the same
    // definition, no override, so the sequence is unbroken and the label a
    // reader sees is the label that was there before.
    //
    // The move is a direct `<w:numPr>` on the paragraph rather than a restyle,
    // so the ONLY thing that differs between the two sides is the instance id —
    // which the key excludes on purpose. A restyle would report a `styleId`
    // change, which is a different and real fact.
    const onInstance4 = (text: string) =>
      `<w:p><w:pPr><w:pStyle w:val="GroupA"/>` +
      `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="4"/></w:numPr>` +
      `</w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
    const { original, revised, result } = await run(
      body(para('GroupA', A1), para('GroupA', A2), para('GroupA', B1)),
      body(para('GroupA', A1), para('GroupA', A2), onInstance4(B1)),
    );

    expect(original.map((r2) => r2.numbering?.label)).toEqual(['1.', '2.', '3.']);
    expect(revised.map((r2) => r2.numbering?.label)).toEqual(['1.', '2.', '3.']);
    expect(properties(result.ops)).toHaveLength(0);
    expect(result.summary.total).toBe(0);
    expect(result.summary.renumberingCascades).toHaveLength(0);
  });
});

describe('a style that re-levels by <w:ilvl> alone', () => {
  // Found in review. A style `basedOn` a numbered base
  // style, overriding nothing but the LEVEL, is ordinary template construction —
  // and `resolveParaNumPr` skipped it, because it walked the chain looking for a
  // `numId` and dropped everything a `numId`-less `<w:numPr>` declared. Every
  // such paragraph resolved to level 0 on BOTH sides, so a re-levelling a reader
  // sees as `(a) (b)` → `(i) (ii)` produced no op at all: ADR-0008 suppression,
  // reached through the one branch of the style chain task 3's acceptance case
  // (a `pStyle` chain that REACHES a `numPr`) does not cover.
  const stylesWithSubAt = (ilvl: number) => `<w:styles ${W}>
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:pPr/></w:style>
    <w:style w:type="paragraph" w:styleId="Body"><w:pPr>
      <w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style>
    <w:style w:type="paragraph" w:styleId="Sub"><w:basedOn w:val="Body"/><w:pPr>
      <w:numPr><w:ilvl w:val="${ilvl}"/></w:numPr></w:pPr></w:style>
  </w:styles>`;

  const NUM_IL = `<w:numbering ${W}>
    <w:abstractNum w:abstractNumId="0">
      <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>
        <w:lvlText w:val="%1."/></w:lvl>
      <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/>
        <w:lvlText w:val="(%2)"/></w:lvl>
      <w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/>
        <w:lvlText w:val="(%3)"/></w:lvl>
    </w:abstractNum>
    <w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>
  </w:numbering>`;

  const DOC = body(para('Body', C1), para('Sub', C2), para('Sub', C3));

  const read = (ilvl: number, policy: 'acceptAll' | 'rejectAll'): ParaRecord[] =>
    preprocessSide(DOC, 'document', policy, config,
      buildNumberingContext(stylesWithSubAt(ilvl), NUM_IL)).paras;

  it('resolves the level the style declares, and reports the template edit', async () => {
    const original = read(1, config.flattenOriginal);
    const revised = read(2, config.flattenRevised);

    // The level the style declares, not the level of the instance it inherits.
    expect(original.map((r) => r.numbering?.label)).toEqual(['1.', '(a)', '(b)']);
    expect(revised.map((r) => r.numbering?.label)).toEqual(['1.', '(i)', '(ii)']);

    const result = await runComparePipeline({
      originalParas: original, revisedParas: revised, config,
      author: 'Mutatis', date: '2026-09-12T00:00:00Z',
    });

    // The body is byte-identical on both sides; only the STYLE moved, so every
    // op here exists because the numbering resolved through the chain.
    const props = properties(result.ops);
    expect(props).toHaveLength(2);
    for (const op of props) {
      expect(op.propertyChange?.numberingIdentity).toBe(true);
      expect(op.propertyChange?.numberLabel).toBe(true);
      expect(op.classification).toBe('formatting');
    }
    expect(props[0]?.propertyChange?.from.ilvl).toBe(1);
    expect(props[0]?.propertyChange?.to.ilvl).toBe(2);
    expect(result.summary.total).toBe(2);
  });
});

describe("ADR-0042 Accepted limit FOUR — a property change riding on a text change", () => {
  // The limit that scopes this plan's central claim, and the reason that claim
  // is written *"for the pairs `detectPropertyChanges` examines"* rather than
  // unqualified. `detectPropertyChanges` converts `equal` ops and nothing else,
  // so a paragraph whose words were edited AND whose number moved is reported
  // once — as the text change — and its renumbering is not reported at all.
  //
  // This test exists to make that limit REGRESS LOUDLY rather than silently. It
  // is the only one of ADR-0042's four accepted limits with no pinning test
  // until now, and it is the one a future pass is most likely to "fix" without
  // recording the trade (marking a `replace` as also-renumbered would put a
  // second row in the change list for one edit, which §6.1 exists to prevent).
  const STYLES_L4 = `<w:styles ${W}>
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:pPr/></w:style>
    <w:style w:type="paragraph" w:styleId="Cl"><w:pPr>
      <w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style>
  </w:styles>`;

  const NUM_L4 = `<w:numbering ${W}>
    <w:abstractNum w:abstractNumId="0">
      <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>
        <w:lvlText w:val="%1."/></w:lvl>
    </w:abstractNum>
    <w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>
  </w:numbering>`;

  const NEW = 'The Purchaser shall deliver the Completion Accounts within twenty Business Days.';
  const C3_EDITED = C3.replace('Reporting Accountant', 'Independent Expert');

  const read = (doc: string, policy: 'acceptAll' | 'rejectAll'): ParaRecord[] =>
    preprocessSide(doc, 'document', policy, config,
      buildNumberingContext(STYLES_L4, NUM_L4)).paras;

  it('is NOT reported — the renumbering rides on the text op and is lost', async () => {
    const original = read(
      body(para('Cl', C1), para('Cl', C2), para('Cl', C3)), config.flattenOriginal);
    const revised = read(
      body(para('Cl', C1), para('Cl', NEW), para('Cl', C2), para('Cl', C3_EDITED)),
      config.flattenRevised);

    // Both of the last two clauses renumber. The reader sees both moves.
    expect(original.map((r) => r.numbering?.label)).toEqual(['1.', '2.', '3.']);
    expect(revised.map((r) => r.numbering?.label)).toEqual(['1.', '2.', '3.', '4.']);

    const result = await runComparePipeline({
      originalParas: original, revisedParas: revised, config,
      author: 'Mutatis', date: '2026-09-12T00:00:00Z',
    });

    // C2 kept its words, so it pairs `equal` and its move IS reported.
    const props = properties(result.ops);
    expect(props).toHaveLength(1);
    expect(props[0]?.propertyChange?.from.numberLabel).toBe('2.');
    expect(props[0]?.propertyChange?.to.numberLabel).toBe('3.');

    // C3 was reworded, so it pairs as a text change — and NOTHING carries the
    // fact that its number moved 3. → 4. No op reports it, and no op for it
    // carries a `propertyChange` at all. That is the limit, stated as a test.
    const textOps = result.ops.filter((o) => o.type !== 'equal' && o.type !== 'property');
    expect(textOps.length).toBeGreaterThan(0);
    for (const op of textOps) expect(op.propertyChange).toBeUndefined();
  });
});
