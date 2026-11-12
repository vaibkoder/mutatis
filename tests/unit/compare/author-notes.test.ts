/**
 * Footnotes and endnotes authored into the redline, every reference pointing at
 * its own note (ADR-0047; retires ADR-0045 limit 9(a)).
 *
 * Driven from the body pipeline's own authored body: the notes follow the
 * references the body's diff kept, inserted and struck.
 */
import { describe, it, expect } from 'vitest';

import { runPipeline } from '@features/compare/pipeline.js';
import { extractParts } from '@features/compare/io/parts.js';
import { authorNotes, resolvedNotes, referenceIds, type NoteKind } from '@features/compare/output/author-notes.js';
import { makeAuthorContext } from '@features/compare/output/index.js';
import { getCompareConfig, getActiveProfile } from '@shared/config/index.js';

import { partsPackage, p, fnRef, enRef, type PackageSpec } from '../../fixtures/parts-docx.js';

const CONFIG = getCompareConfig(getActiveProfile());
const AUTHOR = 'Compare';
const DATE = '2026-10-26T00:00:00Z';

/** Run the body pipeline over two packages, then author their notes of `kind`. */
async function redline(original: PackageSpec, revised: PackageSpec, kind: NoteKind = 'footnote') {
  const o = extractParts(partsPackage(original));
  const r = extractParts(partsPackage(revised));
  const body = await runPipeline({ originalParts: o, revisedParts: r, config: CONFIG, author: AUTHOR, date: DATE });
  const notes = (side: typeof o) => (kind === 'footnote' ? side.footnotesXml : side.endnotesXml);
  const result = await authorNotes({
    kind,
    bodyXml: body.authoredBodyXml,
    original: { documentXml: o.documentXml, notesXml: notes(o) },
    revised: { documentXml: r.documentXml, notesXml: notes(r) },
    config: CONFIG, author: AUTHOR, date: DATE,
    ctx: makeAuthorContext(AUTHOR, DATE, 1000),
  });
  return { o, r, body, result, notes };
}

/** What a reader sees under each reference, both ways — and what each source document shows. */
function views(run: Awaited<ReturnType<typeof redline>>, kind: NoteKind = 'footnote') {
  const { result, o, r, notes } = run;
  return {
    accepted: resolvedNotes(result!.bodyXml, result!.notesXml, kind, 'acceptAll'),
    revised: resolvedNotes(r.documentXml, notes(r), kind, 'acceptAll'),
    rejected: resolvedNotes(result!.bodyXml, result!.notesXml, kind, 'rejectAll'),
    original: resolvedNotes(o.documentXml, notes(o), kind, 'rejectAll'),
  };
}

const TWO_NOTES: PackageSpec = {
  body: [p('The Seller', fnRef(1), ' shall deliver', fnRef(2), ' the Goods on the Completion Date.'), 'Clause two.'],
  footnotes: { 1: ['As defined in the Framework Agreement.'], 2: ['Delivery is to the Buyer’s premises.'] },
};

describe('a note whose text changed', () => {
  it('is tracked inside the note, and the body is untouched', async () => {
    const run = await redline(TWO_NOTES, {
      ...TWO_NOTES,
      footnotes: { 1: ['As defined in the Master Agreement.'], 2: ['Delivery is to the Buyer’s premises.'] },
    });
    expect(run.result).not.toBeNull();
    expect(run.result!.bodyXml).toBe(run.body.authoredBodyXml);
    const note1 = /<w:footnote w:id="1">[\s\S]*?<\/w:footnote>/.exec(run.result!.notesXml!)![0];
    expect(note1).toMatch(/<w:ins [^>]*><w:r>[\s\S]*Master/);
    expect(note1).toMatch(/<w:del [^>]*><w:r>[\s\S]*Framework/);
    expect(run.result!.changed.map((c) => [c.partId, c.ordinal, c.status])).toEqual([['footnote:1', 1, 'kept']]);
    const v = views(run);
    expect(v.accepted).toEqual(v.revised);
    expect(v.rejected).toEqual(v.original);
  });

  it('a note unchanged is written as the revised note, unmarked', async () => {
    const run = await redline(TWO_NOTES, TWO_NOTES);
    expect(run.result!.changed).toEqual([]);
    expect(run.result!.notesXml).not.toMatch(/<w:(ins|del) /);
  });
});

describe('a note added with its reference', () => {
  it('inserts the reference and the note, and the reference resolves', async () => {
    const run = await redline(TWO_NOTES, {
      body: [p('The Seller', fnRef(1), ' shall deliver', fnRef(2), ' the Goods', fnRef(3), ' on the Completion Date.'), 'Clause two.'],
      footnotes: { ...TWO_NOTES.footnotes, 3: ['Including spare parts.'] },
    });
    expect(run.result!.changed.map((c) => [c.partId, c.ordinal, c.status])).toEqual([['footnote:3', 3, 'inserted']]);
    const note3 = /<w:footnote w:id="3">[\s\S]*?<\/w:footnote>/.exec(run.result!.notesXml!)![0];
    expect(note3).toMatch(/<w:ins [^>]*>[\s\S]*spare parts/);
    const v = views(run);
    expect(v.accepted).toEqual(v.revised);
    expect(v.rejected).toEqual(v.original);
  });
});

describe('a note removed with its reference', () => {
  // Word renumbers on save: the surviving note is id 1 in the revised package,
  // the SAME id the struck one has in the original.
  const REMOVED: PackageSpec = {
    body: [p('The Seller shall deliver', fnRef(1), ' the Goods on the Completion Date.'), 'Clause two.'],
    footnotes: { 1: ['Delivery is to the Buyer’s premises.'] },
  };

  it('strikes the reference and the note, the struck reference retargeted to its own note', async () => {
    const run = await redline(TWO_NOTES, REMOVED);
    // Struck: numbered as in the original (1), its part id the fresh output id (3).
    expect(run.result!.changed.map((c) => [c.partId, c.ordinal, c.status])).toEqual([['footnote:3', 1, 'deleted']]);
    // The struck reference came from the original paragraph with the original's
    // id 1 — which in the output is the KEPT note. It must point at its own.
    const ids = referenceIds(run.result!.bodyXml, 'footnote');
    expect(ids).toHaveLength(2);
    expect(new Set(ids).size).toBe(2);
    const struck = /<w:footnote w:id="(\d+)">((?:(?!<\/w:footnote>)[\s\S])*Framework[\s\S]*?)<\/w:footnote>/.exec(run.result!.notesXml!)!;
    expect(ids).toContain(struck[1]);
    expect(struck[2]).toMatch(/<w:del /);
    const v = views(run);
    expect(v.accepted).toEqual(v.revised);
    expect(v.rejected).toEqual(v.original);
  });
});

describe('a clause rewritten whole, its note with it', () => {
  it('a struck note and an inserted note of the same number have distinct part ids', async () => {
    const run = await redline(TWO_NOTES, {
      body: [p('Title to the Goods passes on payment in full', fnRef(1), '.'), 'Clause two.'],
      footnotes: { 1: ['Retention of title applies until then.'] },
    });
    const sets = run.result!.changed;
    expect(sets.map((c) => c.status).sort()).toEqual(['deleted', 'deleted', 'inserted']);
    expect(new Set(sets.map((c) => c.partId)).size).toBe(sets.length);
    const v = views(run);
    expect(v.accepted).toEqual(v.revised);
    expect(v.rejected).toEqual(v.original);
  });
});

describe('two references that swapped notes (ADR-0045 limit 9(a), retired)', () => {
  it('diffs the notes under each place, so the swap is shown', async () => {
    const run = await redline(TWO_NOTES, {
      ...TWO_NOTES,
      footnotes: { 1: ['Delivery is to the Buyer’s premises.'], 2: ['As defined in the Framework Agreement.'] },
    });
    expect(run.result!.changed.map((c) => c.status)).toEqual(['kept', 'kept']);
    expect(run.result!.notesXml).toMatch(/<w:del /);
    const v = views(run);
    expect(v.accepted).toEqual(v.revised);
    expect(v.rejected).toEqual(v.original);
  });
});

describe('endnotes take the same road', () => {
  it('an endnote edited is tracked inside the endnote', async () => {
    const base: PackageSpec = {
      body: [p('The Price', enRef(1), ' is payable in arrears.'), 'Clause two.'],
      endnotes: { 1: ['Exclusive of GST.'] },
    };
    const run = await redline(base, { ...base, endnotes: { 1: ['Exclusive of GST and stamp duty.'] } }, 'endnote');
    expect(run.result!.changed.map((c) => c.partId)).toEqual(['endnote:1']);
    expect(run.result!.notesXml).toMatch(/<w:endnote w:id="1">[\s\S]*<w:ins /);
    const v = views(run, 'endnote');
    expect(v.accepted).toEqual(v.revised);
    expect(v.rejected).toEqual(v.original);
  });
});

describe('a reference only an mc:Fallback carries', () => {
  it('is in neither view, and keeps the id it had', async () => {
    const alt = '<w:r><mc:AlternateContent><mc:Choice Requires="wps">' + fnRef(1) +
      '</mc:Choice><mc:Fallback>' + fnRef(1) + '</mc:Fallback></mc:AlternateContent></w:r>';
    const spec: PackageSpec = {
      body: [p('The Seller shall deliver', alt, ' the Goods.'), 'Clause two.'],
      footnotes: { 1: ['On the Completion Date.'] },
    };
    const run = await redline(spec, spec);
    expect(run.result).not.toBeNull();
    expect(run.result!.bodyXml).toContain(`<mc:Fallback>${fnRef(1)}</mc:Fallback>`);
  });

  it('is not counted when a later note is numbered', async () => {
    const alt = '<w:r><mc:AlternateContent><mc:Choice Requires="wps">' + fnRef(1) +
      '</mc:Choice><mc:Fallback>' + fnRef(1) + '</mc:Fallback></mc:AlternateContent></w:r>';
    const spec = (second: string): PackageSpec => ({
      body: [p('The Seller shall deliver', alt, ' the Goods.'), p('The Buyer shall pay', fnRef(2), '.'), 'Clause three.'],
      footnotes: { 1: ['On the Completion Date.'], 2: [second] },
    });
    const run = await redline(spec('Within 30 days.'), spec('Within 45 days.'));
    expect(run.result!.changed.map((c) => c.ordinal)).toEqual([2]);
  });
});

describe('the self-check refuses what it cannot resolve', () => {
  it('a revised notes part holding two notes under one id → null', async () => {
    const run = await redline(TWO_NOTES, {
      ...TWO_NOTES,
      footnotes: { 1: ['As defined in the Master Agreement.'], 2: ['Delivery is to the Buyer’s premises.'] },
    });
    expect(run.result).not.toBeNull();
    const o = extractParts(partsPackage(TWO_NOTES));
    const r = extractParts(partsPackage(TWO_NOTES));
    const twice = r.footnotesXml!.replace('</w:footnotes>', '<w:footnote w:id="2"><w:p><w:r><w:t>Stray.</w:t></w:r></w:p></w:footnote></w:footnotes>');
    const body = await runPipeline({ originalParts: o, revisedParts: { ...r, footnotesXml: twice }, config: CONFIG, author: AUTHOR, date: DATE });
    const result = await authorNotes({
      kind: 'footnote', bodyXml: body.authoredBodyXml,
      original: { documentXml: o.documentXml, notesXml: o.footnotesXml },
      revised: { documentXml: r.documentXml, notesXml: twice },
      config: CONFIG, author: AUTHOR, date: DATE, ctx: makeAuthorContext(AUTHOR, DATE, 1000),
    });
    expect(result).toBeNull();
  });

  it('one revised note under two references that had two different notes → null', async () => {
    const run = await redline(TWO_NOTES, {
      body: [p('The Seller', fnRef(1), ' shall deliver', fnRef(1), ' the Goods on the Completion Date.'), 'Clause two.'],
      footnotes: { 1: ['As defined in the Framework Agreement.'] },
    });
    expect(run.result).toBeNull();
  });


  it('a reference to a note the part does not hold → null (the original part stays)', async () => {
    const run = await redline(TWO_NOTES, {
      ...TWO_NOTES,
      footnotes: { 1: ['As defined in the Master Agreement.'] }, // note 2 missing
    });
    expect(run.result).toBeNull();
  });
});
