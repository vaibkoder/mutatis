/**
 * A `property` op (ADR-0042) on every surface that shows a change.
 *
 * The op type was added on the argument that a change which is not an op "cannot
 * be counted or navigated to". Counting landed first and the rest did not: the
 * change-list row rendered as a bare category chip, clicking it did nothing, and
 * the exported report printed a row whose two sides were the same string — the
 * exact shape the AI export refuses on the grounds that it reads as a mistake.
 * Both `spec-validator` passes flagged it. These pin the surfaces.
 */
import { describe, it, expect, afterEach, vi } from 'vitest';

import { setActiveHost, _resetActiveHost } from '@shared/host/index.js';
import type { HostPort } from '@shared/host/index.js';
import { PizZipBundle } from '@shared/ooxml/bundle.js';
import { compareDriverState } from '@features/compare/driver.js';

import { previewText } from '@features/compare/components/ChangeList.js';
import { buildCompareReport } from '@features/compare/export/report-model.js';
import { serialiseChangeSet } from '@features/ai-export/serialise/change-set.js';
import { normaliseText } from '@features/compare/preprocess/index.js';
import { extractParaRecords } from '@features/compare/preprocess/index.js';
import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import type { DiffOp, ParaRecord, PropertyChange, ChangeSummary } from '@features/compare/model.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const config = getCompareConfig(SG_PROFILE);

const CLAUSE = 'Settlement of the balance shall be made in cleared funds within ten Business Days.';

function propertyChange(over: Partial<PropertyChange> = {}): PropertyChange {
  return {
    styleId: true,
    numberingIdentity: true,
    numberLabel: true,
    from: { styleId: 'Subclause3', numberLabel: '(vi)', ilvl: 3 },
    to: { styleId: 'Subclause2', numberLabel: '(d)', ilvl: 2 },
    ...over,
  };
}

function propertyOp(over: Partial<DiffOp> = {}): DiffOp {
  return {
    id: 'p1',
    type: 'property',
    origRange: { start: 0, end: 0 },
    revRange: { start: 0, end: 0 },
    partId: 'document',
    propertyChange: propertyChange(),
    classification: 'formatting',
    classes: ['formatting'],
    author: 'Mutatis',
    date: '2026-09-12T00:00:00Z',
    ...over,
  };
}

function record(): ParaRecord {
  return {
    index: 0, partId: 'document', rawXml: '', normText: CLAUSE, normTextAggressive: CLAUSE,
    numberingPrefix: '', hash: 'h', styleId: 'Subclause2', isTableCell: false,
    isSignatureBlock: false, isHeaderFooter: false, isFootnote: false,
  };
}

describe('the change-list row preview', () => {
  it('says what moved, instead of rendering a blank row', () => {
    // `truncate` collapses runs of whitespace, so the separator arrives single-spaced.
    expect(previewText(propertyOp())).toBe('(vi) → (d) · style Subclause3 → Subclause2');
  });

  it('shows the number alone when only the number moved (a cascade)', () => {
    const op = propertyOp({
      propertyChange: propertyChange({
        styleId: false,
        numberingIdentity: false,
        from: { styleId: 'C', numberLabel: '(d)', ilvl: 2 },
        to: { styleId: 'C', numberLabel: '(e)', ilvl: 2 },
      }),
      classification: 'crossRefRenumber',
      classes: ['crossRefRenumber', 'formatting'],
    });

    expect(previewText(op)).toBe('(d) → (e)');
  });

  it('shows the style alone when only the style changed', () => {
    const op = propertyOp({
      propertyChange: propertyChange({
        numberingIdentity: false,
        numberLabel: false,
        from: { styleId: 'Plain', numberLabel: null, ilvl: null },
        to: { styleId: 'PlainAlt', numberLabel: null, ilvl: null },
      }),
    });

    expect(previewText(op)).toBe('style Plain → PlainAlt');
  });
});


/**
 * `BuildCompareReportInput.summary` is REQUIRED, and these fixtures omitted it —
 * which is the drift bringing `tests/` under `tsc` exists to surface.
 *
 * It is also, today, READ BY NOTHING: `report-model.ts` computes `changeCount`
 * from `rows.length` and never touches `input.summary`. Measured — replace the
 * values below with nonsense and all eleven tests in this file still pass. So
 * this helper satisfies the type and asserts nothing; do not read it as coverage.
 * The field's own doc comment claims it "drives the total-count cross-check";
 * there is no such cross-check, and whether the field should exist at all is an
 * open question, not something to settle from a test fixture.
 */
function summaryFor(ops: readonly DiffOp[]): ChangeSummary {
  return {
    total: ops.length,
    byCategory: {
      structural: 0, obligationLanguage: 0, numericalValue: 0, definedTerm: 0,
      crossRefRenumber: 0, substantiveText: 0, whitespacePunct: 0,
      formatting: ops.length,
    },
    renumberingCascades: [], movedParagraphs: 0,
    partChanges: { headers: 0, footers: 0, footnotes: 0 },
    styleDefinitionNote: null,
    spellingVariantCount: 0,
    quarantinedCommentsOriginal: 0, quarantinedCommentsRevised: 0,
    filtersActive: { hideFormattingOnly: true, collapseRenumbering: true },
  };
}

describe('the exported candidate-differences report', () => {
  it('prints the property before/after, not the same paragraph text twice', () => {
    const report = buildCompareReport({
      mode: 'word',
      ops: [propertyOp()],
      summary: summaryFor([propertyOp()]),
      originalParas: [record()],
      revisedParas: [record()],
      generatedAt: '2026-09-12',
      disclosures: [],
    });

    expect(report.rows).toHaveLength(1);
    const row = report.rows[0]!;
    expect(row.type).toBe('property');
    expect(row.original).toBe('(vi) · style: Subclause3');
    expect(row.revised).toBe('(d) · style: Subclause2');
    // The defect this replaces: two identical strings side by side.
    expect(row.original).not.toBe(row.revised);
  });

  it('says what changed when the identity moved and the label did not', () => {
    // ADR-0042 § "Where the key still exceeds the label": a list can be
    // REDEFINED under a paragraph whose own number does not move — a
    // `w:lvlRestart` edit on a one-clause list, say. `numberLabel` is then
    // false, both sides resolve the SAME label, and the row printed that label
    // twice: exactly the artefact this function's doc comment says it exists to
    // prevent. Both surfaces had a fallback for "no label resolved" and none
    // for "label resolved identically".
    const same: PropertyChange = {
      styleId: false,
      numberingIdentity: true,
      numberLabel: false,
      from: { styleId: 'Clause2', numberLabel: '1.1', ilvl: 1 },
      to: { styleId: 'Clause2', numberLabel: '1.1', ilvl: 1 },
    };
    const report = buildCompareReport({
      mode: 'word',
      ops: [propertyOp({ propertyChange: same })],
      summary: summaryFor([propertyOp({ propertyChange: same })]),
      originalParas: [record()],
      revisedParas: [record()],
      generatedAt: '2026-09-12',
      disclosures: [],
    });

    const row = report.rows[0]!;
    expect(row.original).toBe('1.1');
    expect(row.revised).toBe('1.1 (list definition changed)');
    expect(row.original).not.toBe(row.revised);
    // And the on-screen row, which builds its own string from the same facts.
    expect(previewText(propertyOp({ propertyChange: same })))
      .toBe('1.1 (list definition changed)');
  });

  it('names a label that renders as NOTHING rather than printing a gap', () => {
    // A level at `numFmt="none"` (or with an empty `lvlText`) resolves to a
    // label that is the EMPTY STRING — which is not "no label resolved", and is
    // not a level change either. The export row printed an empty original cell
    // beside a revised cell reading " (list definition changed)": the blank-row
    // artefact this function exists to prevent, reached from the other end.
    // Found in review. , which walked a ten-shape matrix of
    // label/level/style combinations across both surfaces.
    const invisible: PropertyChange = {
      styleId: false,
      numberingIdentity: true,
      numberLabel: false,
      from: { styleId: 'Clause2', numberLabel: '', ilvl: 1 },
      to: { styleId: 'Clause2', numberLabel: '', ilvl: 1 },
    };
    const report = buildCompareReport({
      mode: 'word',
      ops: [propertyOp({ propertyChange: invisible })],
      summary: summaryFor([propertyOp({ propertyChange: invisible })]),
      originalParas: [record()],
      revisedParas: [record()],
      generatedAt: '2026-09-13',
      disclosures: [],
    });

    const row = report.rows[0]!;
    expect(row.original).toBe('(no visible number)');
    expect(row.revised).toBe('(no visible number) (list definition changed)');
    expect(previewText(propertyOp({ propertyChange: invisible })))
      .toBe('(no visible number) (list definition changed)');
  });

  it('does NOT call a re-levelling a definition change, when the label is the same', () => {
    // The wrong-fire direction of the test above, and it is a different fact
    // about the document, not a vaguer version of the same one: `ilvl` 0 → 1
    // with the SAME `numbering.xml` on both sides is a paragraph that moved
    // level, and whether its label moves depends only on the templates. Under
    // `%1.`-style levels it does not — so both sides read "1.", and a branch
    // keyed on "the two sides are equal" would announce a list definition
    // change that never happened.
    const relevelled: PropertyChange = {
      styleId: false,
      numberingIdentity: true,
      numberLabel: false,
      from: { styleId: null, numberLabel: '1.', ilvl: 0 },
      to: { styleId: null, numberLabel: '1.', ilvl: 1 },
    };
    const report = buildCompareReport({
      mode: 'word',
      ops: [propertyOp({ propertyChange: relevelled })],
      summary: summaryFor([propertyOp({ propertyChange: relevelled })]),
      originalParas: [record()],
      revisedParas: [record()],
      generatedAt: '2026-09-12',
      disclosures: [],
    });

    const row = report.rows[0]!;
    expect(row.original).toBe('1. · level 1');
    expect(row.revised).toBe('1. · level 2');
    expect(row.revised).not.toContain('list definition changed');
    expect(previewText(propertyOp({ propertyChange: relevelled })))
      .toBe('1. (level 1) → 1. (level 2)');
  });

  it('falls back to the LEVEL when no label resolved, and says so when there is none', () => {
    // A `numId` missing from `numbering.xml`, or an `ilvl` past the levels the
    // `abstractNum` defines, resolves to no label on EITHER side — and then a
    // label-only row prints "— → —" for a real re-levelling. The level is always
    // known, because it is what the style chain resolved, so the row says
    // "level 4 → level 3" instead. `(unnumbered)` is the remaining case: a
    // paragraph that carries no numbering at all on that side.
    const report = buildCompareReport({
      mode: 'word',
      ops: [
        propertyOp({
          propertyChange: propertyChange({
            styleId: false,
            from: { styleId: 'X', numberLabel: null, ilvl: 3 },
            to: { styleId: 'X', numberLabel: null, ilvl: 2 },
          }),
        }),
        propertyOp({
          id: 'p2',
          propertyChange: propertyChange({
            styleId: false,
            from: { styleId: 'X', numberLabel: null, ilvl: null },
            to: { styleId: 'X', numberLabel: '2.', ilvl: 0 },
          }),
        }),
      ],
      summary: summaryFor([
        propertyOp({
          propertyChange: propertyChange({
            styleId: false,
            from: { styleId: 'X', numberLabel: null, ilvl: 3 },
            to: { styleId: 'X', numberLabel: null, ilvl: 2 },
          }),
        }),
        propertyOp({
          id: 'p2',
          propertyChange: propertyChange({
            styleId: false,
            from: { styleId: 'X', numberLabel: null, ilvl: null },
            to: { styleId: 'X', numberLabel: '2.', ilvl: 0 },
          }),
        }),
      ]),
      originalParas: [record()],
      revisedParas: [record()],
      generatedAt: '2026-09-12',
      disclosures: [],
    });

    expect(report.rows[0]!.original).toBe('level 4');
    expect(report.rows[0]!.revised).toBe('level 3');
    // Gained numbering it did not have: the honest word, not a fabricated level.
    expect(report.rows[1]!.original).toBe('(unnumbered)');
    expect(report.rows[1]!.revised).toBe('2.');
  });
});

describe('the AI-readable change set', () => {
  it('excludes property ops DELIBERATELY — rule F1 is before/after text', () => {
    // Not an oversight: `ChangeEntry.type` is a closed four-member union in
    // spec 11 §4.4, and a property change has no text delta to put in
    // before/after. Recorded in ADR-0042 as feature 11's to add. This test
    // exists so the exclusion cannot be "fixed" into a schema violation by
    // someone who reads it as a bug.
    const set = serialiseChangeSet(
      [propertyOp(), propertyOp({ id: 'i1', type: 'insert', origRange: null })],
      {
        baseline: 'a.docx',
        revised: 'b.docx',
        clauseFor: () => 'Clause 7.3',
        revParaText: () => CLAUSE,
      },
    );

    expect(set.changes.map((c) => c.id)).toEqual(['i1']);
  });
});

describe('field instruction text never enters the comparable text', () => {
  it('collects w:t only, so a REF field contributes its RESULT and not its code', () => {
    // Task 5's surviving acceptance. The class it belonged to measured zero and
    // was dropped, but this is an invariant of `extractRunText`, not a feature of
    // that class — and it was correct by OMISSION rather than by design, which is
    // exactly the kind of correctness that disappears in a refactor.
    const xml =
      `<w:document ${W}><w:body><w:p>` +
      '<w:r><w:t xml:space="preserve">See </w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> REF _Ref70001 \\h \\w </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:t>Clause 7.3</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
      '</w:p></w:body></w:document>';

    const [para] = extractParaRecords(xml, 'document', config);

    expect(para!.normText).toBe(normaliseText('See Clause 7.3'));
    expect(para!.normText).not.toContain('REF');
    expect(para!.normText).not.toContain('_Ref70001');
  });
});

describe('clicking a property row navigates somewhere', () => {
  afterEach(() => {
    _resetActiveHost();
    vi.restoreAllMocks();
  });

  /** A .docx whose one clause carries `styleId`, so a style change is all that differs. */
  function makeDocx(name: string, styleId: string): File {
    const bundle = new PizZipBundle();
    bundle.setPart(
      'word/document.xml',
      `<w:document ${W}><w:body>` +
        `<w:p><w:pPr><w:pStyle w:val="${styleId}"/></w:pPr>` +
        `<w:r><w:t xml:space="preserve">${CLAUSE}</w:t></w:r></w:p>` +
        '</w:body></w:document>',
    );
    bundle.setPart('word/styles.xml', `<w:styles ${W}><w:docDefaults/></w:styles>`);
    return new File([new Uint8Array(bundle.toUint8Array())], name);
  }

  it('anchors on the paragraph\'s own text — it has no changed text to aim at', async () => {
    // `goToChange` anchors on `op.wordDiffs`, and a property op has none, so the
    // row used to return early and the click did nothing at all — silently.
    const findAndReveal = vi.fn().mockResolvedValue(undefined);
    setActiveHost({
      writeBack: vi.fn().mockResolvedValue(undefined),
      findAndReveal,
    } as unknown as HostPort);

    const d = compareDriverState;
    d.uploadDocx('original', makeDocx('orig.docx', 'Plain'));
    d.uploadDocx('revised', makeDocx('rev.docx', 'PlainAlt'));
    await d.runCompare();

    const property = d.ops.value!.find((op) => op.type === 'property');
    expect(property, 'the style change produced no property op').toBeDefined();

    d.goToChange(property!.id);

    expect(findAndReveal).toHaveBeenCalledTimes(1);
    expect(findAndReveal.mock.calls[0]![0]).toBe(CLAUSE);
  });
});
