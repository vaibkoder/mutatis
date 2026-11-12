/**
 * Office-context paths: REF-field/bookmark read (T5/A12/A13), Assign-source
 * replacement (T9/A15), and the layer-(c) OOXML probe (T7/B6/B8).
 *
 * These exercise the modules that need a Word.RequestContext, using a minimal
 * hand-built mock. No real Office JS is involved.
 */
import { describe, it, expect, vi } from 'vitest';
import type { ParaNode } from '@shared/doc-model/types.js';
import type { ReferenceMatch, StaticNumberingResult } from '@features/cross-reference/types.js';
import type { SectionNode } from '@shared/section-index/types.js';
import { readFieldsAndBookmarks } from '@features/cross-reference/fields.js';
import { applyAssignSource } from '@features/cross-reference/assign-source.js';
import { probeOoxmlLayer } from '@features/cross-reference/static-numbering-ooxml.js';
import { analyseCrossReference } from '@features/cross-reference/driver.js';
import { UK_PROFILE } from '@shared/config/index.js';

// ---- Mock Word context helpers ----

function mockField(code: string, resultText: string) {
  return { load: () => {}, code, result: { text: resultText } };
}

function mockRange(opts: {
  fields?: ReturnType<typeof mockField>[];
  ooxml?: string;
  insertText?: ReturnType<typeof vi.fn>;
} = {}) {
  return {
    fields: { load: () => {}, items: opts.fields ?? [] },
    getOoxml: () => ({ value: opts.ooxml ?? '' }),
    insertText: opts.insertText ?? vi.fn(),
    search: () => ({ load: () => {}, items: [] }),
  };
}

function mockCtx(bookmarkNames: string[] = [], searchHit?: ReturnType<typeof mockRange>) {
  const bodyRange = mockRange();
  return {
    document: {
      bookmarks: { load: () => {}, items: bookmarkNames.map(n => ({ name: n })) },
      // resolveAnchor re-resolves via a document-wide body.search(); return the
      // provided hit so navigation/replacement targets the matched span.
      body: {
        getRange: () => bodyRange,
        search: () => ({ load: () => {}, items: searchHit ? [searchHit] : [] }),
      },
    },
    sync: async () => {},
  } as unknown as Word.RequestContext;
}

function paraWithRange(index: number, text: string, range: ReturnType<typeof mockRange>): ParaNode {
  return {
    id: `body-para-${index}`,
    index,
    text,
    styleName: 'Normal',
    container: 'body',
    getRange: () => range as unknown as Word.Range,
  };
}

function makeMatch(paraId: string, rawText: string): ReferenceMatch {
  return {
    matchId: `m-${paraId}-${rawText}`,
    rawText,
    anchor: { paragraphId: paraId, startOffset: 0, endOffset: rawText.length },
    paragraphId: paraId,
    unitType: 'clause',
    normalizedKey: 'clause:4.2',
    expandedKeys: ['clause:4.2'],
    classification: 'ok',
    candidateTargets: [],
    isFieldBacked: false,
    staticBadge: true,
    severity: 'info',
  };
}

// ---- A12: field-backed reference ----

describe('fields.ts — REF field detection (spec §5.1.4)', () => {
  it('A12: reference backed by a live REF field → isFieldBacked, no static badge', async () => {
    const range = mockRange({ fields: [mockField('REF _Ref1 \\h', 'Clause 4.2')] });
    const p = paraWithRange(0, 'See Clause 4.2 for details.', range);
    const match = makeMatch(p.id, 'Clause 4.2');

    const res = await readFieldsAndBookmarks([match], [p], mockCtx(['_Ref1']));
    expect(res.fieldsApiAvailable).toBe(true);
    expect(match.isFieldBacked).toBe(true);
    expect(match.staticBadge).toBe(false);
    expect(match.fieldId).toBe('_Ref1');
  });

  it('A13: REF field whose bookmark was deleted → Not Found, high confidence (error)', async () => {
    const range = mockRange({ fields: [mockField('REF _Ref1 \\h', 'Clause 4.2')] });
    const p = paraWithRange(0, 'See Clause 4.2 for details.', range);
    const match = makeMatch(p.id, 'Clause 4.2');
    match.classification = 'ok'; // index still has clause 4.2

    // Bookmark _Ref1 absent from the document → broken REF.
    await readFieldsAndBookmarks([match], [p], mockCtx([]));
    expect(match.isFieldBacked).toBe(true);
    expect(match.classification).toBe('notFound');
    expect(match.severity).toBe('error');
  });

  it('C2: paragraph with one REF field + one manual ref → only the matching ref is field-backed', async () => {
    // Field result is "Clause 7.1"; the paragraph also contains manual "Clause 4.2".
    const range = mockRange({ fields: [mockField('REF _Ref7 \\h', 'Clause 7.1')] });
    const p = paraWithRange(0, 'See Clause 4.2 and Clause 7.1.', range);
    const manual = makeMatch(p.id, 'Clause 4.2');
    const field = makeMatch(p.id, 'Clause 7.1');

    await readFieldsAndBookmarks([manual, field], [p], mockCtx(['_Ref7']));
    expect(field.isFieldBacked).toBe(true);
    expect(field.staticBadge).toBe(false);
    // The manually-typed one must KEEP its static badge.
    expect(manual.isFieldBacked).toBe(false);
    expect(manual.staticBadge).toBe(true);
  });

  it('no REF field in paragraph → static badge retained', async () => {
    const range = mockRange({ fields: [] });
    const p = paraWithRange(0, 'See Clause 4.2.', range);
    const match = makeMatch(p.id, 'Clause 4.2');
    await readFieldsAndBookmarks([match], [p], mockCtx([]));
    expect(match.isFieldBacked).toBe(false);
    expect(match.staticBadge).toBe(true);
  });

  it('A13 end-to-end: broken REF survives the driver even though clause 4.2 exists in the index', async () => {
    // Regression guard for the second-resolveAll clobber: the driver must NOT
    // re-resolve from the index after the field read, or the broken-REF verdict
    // would be silently reverted to OK.
    const clausePara: ParaNode = {
      id: 'body-para-0', index: 0, text: 'Payment terms', styleName: 'Normal',
      container: 'body', listString: '4.2', listLevel: 1,
      getRange: () => mockRange() as unknown as Word.Range,
    };
    const refRange = mockRange({ fields: [mockField('REF _Ref1 \\h', 'Clause 4.2')] });
    const prosePara = paraWithRange(1, 'As set out in Clause 4.2, payment is due.', refRange);

    // Bookmark _Ref1 absent → broken REF, even though clause:4.2 is in the index.
    const result = await analyseCrossReference([clausePara, prosePara], UK_PROFILE, mockCtx([]));
    const notFound = result.xrefFindings.filter(f => f.category === 'cross-reference/not-found');
    expect(notFound.length).toBe(1);
    expect(notFound[0]?.severity).toBe('error');
  });
});

// ---- A15: Assign-source replacement ----

describe('assign-source.ts — static text replacement (spec §8.3, A15)', () => {
  it('A15: replaces matched range text with the target displayLabel', async () => {
    const insertText = vi.fn();
    const hit = mockRange({ insertText });
    const p = paraWithRange(0, 'see Clause 9.9 here', mockRange());
    const anchor = { paragraphId: p.id, startOffset: 4, endOffset: 13 }; // "Clause 9.9"
    const target: SectionNode = {
      key: 'clause:4.3', kind: 'clause', rawNumber: '4.3', displayLabel: 'Clause 4.3',
      level: 0, paraId: 'x', isStaticNumbered: false, bookmarkNames: [],
      getRange: () => { throw new Error('unused'); },
    };

    await applyAssignSource(mockCtx([], hit), [p], anchor, target);
    expect(insertText).toHaveBeenCalledWith('Clause 4.3', 'Replace');
  });

  it('replacement label never emits "Section"/"Exhibit" (uses configured displayLabel)', async () => {
    const insertText = vi.fn();
    const hit = mockRange({ insertText });
    const p = paraWithRange(0, 'see Exhibit B here', mockRange());
    const anchor = { paragraphId: p.id, startOffset: 4, endOffset: 13 };
    const target: SectionNode = {
      key: 'schedule:B', kind: 'schedule', rawNumber: 'B', displayLabel: 'Schedule B',
      level: 0, paraId: 'x', isStaticNumbered: false, bookmarkNames: [],
      getRange: () => { throw new Error('unused'); },
    };
    await applyAssignSource(mockCtx([], hit), [p], anchor, target);
    const [label] = insertText.mock.calls[0] ?? [];
    expect(label).toBe('Schedule B');
    expect(label).not.toContain('Exhibit');
    expect(label).not.toContain('Section');
  });
});

// ---- B6 / B8: layer (c) OOXML probe ----

function staticResult(paraId: string): StaticNumberingResult {
  return {
    paragraphId: paraId,
    leadingNumberText: '4.2',
    hasListItem: false,
    listStringIfAny: null,
    verdict: 'static',
    confidence: 'high',
    severity: 'info',
  };
}

describe('static-numbering-ooxml.ts — layer (c) probe (spec §5.2.3)', () => {
  it('B6: field-driven number → OOXML probe reclassifies static→dynamic', async () => {
    const fieldRange = mockRange({ ooxml: '<w:p><w:fldSimple w:instr="SEQ"><w:r><w:t>4.2</w:t></w:r></w:fldSimple></w:p>' });
    const tRunRange = mockRange({ ooxml: '<w:p><w:r><w:t>4.2 The Buyer…</w:t></w:r></w:p>' });
    const pField = paraWithRange(0, '4.2 …', fieldRange);
    const pStatic = paraWithRange(1, '4.2 The Buyer…', tRunRange);
    const rField = staticResult(pField.id);
    const rStatic = staticResult(pStatic.id);

    await probeOoxmlLayer(
      [{ para: pField, result: rField }, { para: pStatic, result: rStatic }],
      mockCtx(),
      UK_PROFILE,
    );
    expect(rField.verdict).toBe('dynamic'); // field → not a manually-typed number
    expect(rStatic.verdict).toBe('static'); // <w:t> run, no field → confirmed static
    expect(rStatic.ooxmlVerdict).toBe('tRun');
  });

  it('B8: probes capped at maxOoxmlProbes; beyond-cap residuals KEEP static verdict', async () => {
    const profile = { ...UK_PROFILE, staticNumbering: { ...UK_PROFILE.staticNumbering, maxOoxmlProbes: 2 } };
    // All five would probe as field-driven (→dynamic) IF probed.
    const residuals = Array.from({ length: 5 }, (_, i) => {
      const range = mockRange({ ooxml: '<w:p><w:fldSimple><w:r><w:t>4.2</w:t></w:r></w:fldSimple></w:p>' });
      const para = paraWithRange(i, '4.2 …', range);
      return { para, result: staticResult(para.id) };
    });

    await probeOoxmlLayer(residuals, mockCtx(), profile);

    const dynamic = residuals.filter(r => r.result.verdict === 'dynamic');
    const stillStatic = residuals.filter(r => r.result.verdict === 'static');
    // Only the first 2 (cap) were probed → dynamic; the other 3 stay flagged static.
    expect(dynamic).toHaveLength(2);
    expect(stillStatic).toHaveLength(3);
  });
});
