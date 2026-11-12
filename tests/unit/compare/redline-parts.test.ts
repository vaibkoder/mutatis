/**
 * The redline PACKAGE carries the changes of its auxiliary parts as tracked
 * changes inside those parts.
 *
 * Driven through the Word-vs-Word driver to the host's `writeBack`, so what is
 * asserted is the `.docx` a user receives — not an intermediate string.
 */
import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';

import { setActiveHost, _resetActiveHost } from '@shared/host/index.js';
import type { HostPort } from '@shared/host/index.js';
import type { DocxBundle } from '@shared/ooxml/index.js';
import { resolvedNotes } from '@features/compare/output/author-notes.js';
import { extractParts, parseRelationships, resolveTarget } from '@features/compare/io/parts.js';
import { runPipeline } from '@features/compare/pipeline.js';

import { partsPackage, p, para, fnRef, pageField, type PackageSpec } from '../../fixtures/parts-docx.js';
import { reconstruct, isClean, summarise } from '../../fixtures/reconstruct.js';

let config: Record<string, unknown> = {};
vi.mock('@shared/config/index.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@shared/config/index.js')>();
  return {
    ...actual,
    getCompareConfig: (profile: Parameters<typeof actual.getCompareConfig>[0]) => ({
      ...actual.getCompareConfig(profile),
      ...config,
    }),
  };
});

const { createCompareDriver } = await import('@features/compare/driver.js');
const { getCompareConfig, getActiveProfile } = await import('@shared/config/index.js');

let written: DocxBundle | null = null;

beforeEach(() => {
  written = null;
  setActiveHost({
    writeBack: vi.fn().mockImplementation(async (bundle: DocxBundle) => { written = bundle; }),
  } as unknown as HostPort);
});

afterEach(() => {
  _resetActiveHost();
  vi.restoreAllMocks();
  config = {};
});

const file = (spec: PackageSpec, name: string): File =>
  new File([new Uint8Array(partsPackage(spec).toUint8Array())], name);

/** Run the driver over two package specs; return the written package and the driver. */
async function compare(original: PackageSpec, revised: PackageSpec) {
  const driver = createCompareDriver();
  driver.uploadDocx('original', file(original, 'original.docx'));
  driver.uploadDocx('revised', file(revised, 'revised.docx'));
  await driver.runCompare();
  expect(driver.errorMessage.value).toBeNull();
  expect(written).not.toBeNull();
  return { pkg: written!, driver };
}

const NOTES: PackageSpec = {
  body: [
    p('The Seller', fnRef(1), ' shall deliver', fnRef(2), ' the Goods on the Completion Date.'),
    p('The Buyer shall pay the Price', fnRef(3), ' within 30 days.'),
    'Clause three.',
  ],
  footnotes: {
    1: ['As defined in the Framework Agreement.'],
    2: ['Delivery is to the Buyer’s premises.'],
    3: ['Exclusive of GST.'],
  },
};

/** One note edited (2), one removed (1), one added (after "Clause three"). */
const NOTES_REVISED: PackageSpec = {
  body: [
    p('The Seller shall deliver', fnRef(1), ' the Goods on the Completion Date.'),
    p('The Buyer shall pay the Price', fnRef(2), ' within 30 days.'),
    p('Clause three', fnRef(3), '.'),
  ],
  footnotes: {
    1: ['Delivery is to the Seller’s premises.'],
    2: ['Exclusive of GST.'],
    3: ['Inserted at the Buyer’s request.'],
  },
};

describe('footnotes in the redline package', () => {
  it('carries an edit, an addition and a removal as tracked changes inside footnotes.xml', async () => {
    const { pkg } = await compare(NOTES, NOTES_REVISED);
    const notes = pkg.part('word/footnotes.xml')!;
    expect(notes).toMatch(/<w:ins [^>]*>(?:(?!<\/w:ins>)[\s\S])*Seller/);  // the edit, inside the note
    expect(notes).toMatch(/<w:del [^>]*>(?:(?!<\/w:del>)[\s\S])*Framework/); // the removed note, struck
    expect(notes).toMatch(/<w:ins [^>]*>(?:(?!<\/w:ins>)[\s\S])*Buyer’s request/); // the added note, inserted
  });

  it('every reference resolves to its own note in both views of the package', async () => {
    const { pkg } = await compare(NOTES, NOTES_REVISED);
    const body = pkg.part('word/document.xml')!;
    const notes = pkg.part('word/footnotes.xml')!;
    const src = (spec: PackageSpec) => partsPackage(spec);
    const o = src(NOTES);
    const r = src(NOTES_REVISED);
    expect(resolvedNotes(body, notes, 'footnote', 'acceptAll'))
      .toEqual(resolvedNotes(r.part('word/document.xml')!, r.part('word/footnotes.xml')!, 'footnote', 'acceptAll'));
    expect(resolvedNotes(body, notes, 'footnote', 'rejectAll'))
      .toEqual(resolvedNotes(o.part('word/document.xml')!, o.part('word/footnotes.xml')!, 'footnote', 'rejectAll'));
  });

  it('a package whose original has no footnotes part gains one, wired by relationship and content type', async () => {
    const plain: PackageSpec = { body: ['The Seller shall deliver the Goods.', 'Clause two.'] };
    const { pkg } = await compare(plain, {
      body: [p('The Seller shall deliver the Goods', fnRef(1), '.'), 'Clause two.'],
      footnotes: { 1: ['On the Completion Date.'] },
    });
    expect(pkg.part('word/footnotes.xml')).toMatch(/<w:ins [^>]*>[\s\S]*Completion Date/);
    expect(pkg.part('word/_rels/document.xml.rels')).toMatch(/Type="[^"]*\/footnotes" Target="footnotes.xml"/);
    expect(pkg.part('[Content_Types].xml')).toContain('PartName="/word/footnotes.xml"');
  });

  it('the rendered preview reads the redline’s own notes part, which the body’s references point into', async () => {
    const { pkg, driver } = await compare(NOTES, NOTES_REVISED);
    expect(driver.renderSource.value?.footnotesXml).toBe(pkg.part('word/footnotes.xml'));
  });

  it('notes the self-check refuses keep the pre-ADR output, and their changes are still counted', async () => {
    // The revised body references a note its part does not hold.
    const broken: PackageSpec = { ...NOTES_REVISED, footnotes: { 1: ['Delivery is to the Seller’s premises.'] } };
    const { pkg, driver } = await compare(NOTES, broken);
    expect(pkg.part('word/footnotes.xml')).toBe(partsPackage(NOTES).part('word/footnotes.xml'));
    expect(driver.summary.value!.partChanges.footnotes).toBeGreaterThan(0);
  });

  it('the body gets its last sectPr back (page setup and the last section’s references)', async () => {
    const { pkg } = await compare(NOTES, NOTES);
    expect(pkg.part('word/document.xml')).toMatch(/<w:sectPr>[\s\S]*<w:pgSz w:w="11906"[\s\S]*<\/w:sectPr><\/w:body>/);
  });
});

// ---------------------------------------------------------------------------
// Task 3 — headers and footers
// ---------------------------------------------------------------------------

/** Every header/footer reference of the written body resolves to a part of its kind. */
function referencesResolve(pkg: DocxBundle): void {
  const rels = parseRelationships(pkg.part('word/_rels/document.xml.rels'));
  const doc = pkg.part('word/document.xml')!;
  for (const m of doc.matchAll(/<w:(header|footer)Reference\b[^>]*r:id="([^"]*)"/g)) {
    const rel = rels.find((r) => r.id === m[2]);
    expect(rel?.type, `${m[0]} resolves`).toBe(m[1]);
    expect(pkg.hasPart(resolveTarget('word', rel!.target))).toBe(true);
  }
}

/** The 08i yardstick on one part: authored part both ways against the two sources. */
function partReconstructs(authored: string, original: string, revised: string): void {
  const report = reconstruct(authored, original, revised);
  expect(isClean(report.accept), summarise(report.accept)).toBe(true);
  expect(isClean(report.reject), summarise(report.reject)).toBe(true);
}

const FOOTED = (stamp: string): PackageSpec => ({
  body: ['The Seller shall deliver the Goods.', 'Clause two.'],
  sections: [{ headers: { default: 'header1.xml' }, footers: { default: 'footer1.xml' } }],
  parts: {
    'header1.xml': [para('Project Sample')],
    'footer1.xml': [para(stamp), p('Page ', pageField('1'))],
  },
});

describe('headers and footers in the redline package', () => {
  it('a footer stamp edited is tracked inside the footer part, and reconstructs both ways', async () => {
    config = { excludeHeaders: false, excludeFooters: false };
    const { pkg } = await compare(FOOTED('Draft for discussion'), FOOTED('Execution version'));
    const footer = pkg.part('word/footer1.xml')!;
    expect(footer).toMatch(/<w:ins [^>]*>(?:(?!<\/w:ins>)[\s\S])*Execution/);
    expect(footer).toMatch(/<w:del [^>]*>(?:(?!<\/w:del>)[\s\S])*Draft/);
    expect(footer).toContain('PAGE'); // the field is kept
    partReconstructs(footer,
      partsPackage(FOOTED('Draft for discussion')).part('word/footer1.xml')!,
      partsPackage(FOOTED('Execution version')).part('word/footer1.xml')!);
    // The unchanged header is the original's part, byte for byte.
    expect(pkg.part('word/header1.xml')).toBe(partsPackage(FOOTED('x')).part('word/header1.xml'));
    referencesResolve(pkg);
  });

  it('w:ids stay unique across the body and the parts', async () => {
    config = { excludeHeaders: false, excludeFooters: false };
    const body = (word: string): PackageSpec => ({ ...FOOTED(`Draft ${word}`), body: [`The Seller shall ${word} the Goods.`, 'Clause two.'] });
    const { pkg } = await compare(body('deliver'), body('supply'));
    const ids = [pkg.part('word/document.xml')!, pkg.part('word/footer1.xml')!]
      .flatMap((x) => [...x.matchAll(/<w:(?:ins|del) w:id="(\d+)"/g)].map((m) => m[1]));
    expect(ids.length).toBeGreaterThan(2);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it('with no part changed, the body is what the body pipeline wrote, plus its last sectPr', async () => {
    config = { excludeHeaders: false, excludeFooters: false };
    const original = { ...FOOTED('Draft'), body: ['The Seller shall deliver the Goods.', 'Clause two.'] };
    const revised = { ...FOOTED('Draft'), body: ['The Seller shall supply the Goods.', 'Clause two.'] };
    const { pkg } = await compare(original, revised);
    const o = extractParts(partsPackage(original));
    const r = extractParts(partsPackage(revised));
    const cfg = { ...getCompareConfig(getActiveProfile()), excludeHeaders: false, excludeFooters: false };
    const alone = await runPipeline({ originalParts: o, revisedParts: r, config: cfg, author: cfg.author, date: 'x' });
    const written = /<w:body>([\s\S]*)<w:sectPr>[\s\S]*<\/w:sectPr><\/w:body>/.exec(pkg.part('word/document.xml')!)![1];
    const strip = (x: string): string => x.replace(/ w:date="[^"]*"/g, '');
    expect(strip(written!)).toBe(strip(alone.authoredBodyXml.replace(/^<w:body>|<\/w:body>$/g, '')));
  });

  it('a first-page header only the revised document has is added, wired, referenced, and inserted', async () => {
    config = { excludeHeaders: false };
    const revised: PackageSpec = {
      ...FOOTED('Draft'),
      sections: [{ headers: { default: 'header1.xml', first: 'header2.xml' }, footers: { default: 'footer1.xml' } }],
      parts: { ...FOOTED('Draft').parts, 'header2.xml': [para('Strictly private and confidential')] },
    };
    const { pkg, driver } = await compare(FOOTED('Draft'), revised);
    const rels = parseRelationships(pkg.part('word/_rels/document.xml.rels'));
    const added = rels.find((r) => r.id.startsWith('rIdDp'))!;
    expect(added.type).toBe('header');
    const part = pkg.part(resolveTarget('word', added.target))!;
    expect(part).toMatch(/<w:ins [^>]*>[\s\S]*Strictly private/);
    expect(pkg.part('word/document.xml')).toContain(`<w:headerReference w:type="first" r:id="${added.id}"/>`);
    expect(pkg.part('[Content_Types].xml')).toContain(`PartName="/${resolveTarget('word', added.target)}"`);
    expect(driver.summary.value!.partChanges.headers).toBe(1);
    referencesResolve(pkg);
  });

  it('a footer only the original has stays referenced, its paragraphs struck', async () => {
    config = { excludeFooters: false };
    const revised: PackageSpec = { ...FOOTED('Draft'), sections: [{ headers: { default: 'header1.xml' } }] };
    const { pkg } = await compare(FOOTED('Draft for discussion'), revised);
    const footer = pkg.part('word/footer1.xml')!;
    expect(footer).toMatch(/<w:del [^>]*>(?:(?!<\/w:del>)[\s\S])*Draft for discussion/);
    expect(pkg.part('word/document.xml')).toMatch(/<w:footerReference w:type="default" r:id="rId\d+"\/>/);
    referencesResolve(pkg);
  });

  it('a section break’s references, written from the revised side, point into the output package even when excluded', async () => {
    // The revised package numbers its relationships differently (an extra part first).
    const original: PackageSpec = {
      body: ['Clause one.', '§', 'Schedule one.'],
      sections: [{ footers: { default: 'footer1.xml' } }, { footers: { default: 'footer2.xml' } }],
      parts: { 'footer1.xml': [para('Agreement')], 'footer2.xml': [para('Schedule')] },
    };
    const revised: PackageSpec = {
      ...original,
      sections: [{ headers: { default: 'header9.xml' }, footers: { default: 'footer1.xml' } }, { footers: { default: 'footer2.xml' } }],
      parts: { ...original.parts, 'header9.xml': [para('Cover')] },
    };
    expect(partsPackage(revised).part('word/document.xml')).toContain('r:id="rId11"'); // ids shifted
    config = { excludeHeaders: true, excludeFooters: true };
    const { pkg } = await compare(original, revised);
    referencesResolve(pkg);
    const doc = pkg.part('word/document.xml')!;
    expect(doc.match(/<w:footerReference [^>]*>/g)).toEqual([
      '<w:footerReference w:type="default" r:id="rId10"/>',
      '<w:footerReference w:type="default" r:id="rId11"/>',
    ]);
    expect(doc).not.toContain('headerReference'); // excluded, and the original has no such header
  });

  it('one original footer two sections share, paired with two different revised footers, is written twice', async () => {
    config = { excludeFooters: false };
    const original: PackageSpec = {
      body: ['Clause one.', '§', 'Schedule one.'],
      sections: [{ footers: { default: 'footer1.xml' } }, { footers: { default: 'footer1.xml' } }],
      parts: { 'footer1.xml': [para('Agreement')] },
    };
    const revised: PackageSpec = {
      ...original,
      sections: [{ footers: { default: 'footer1.xml' } }, { footers: { default: 'footer2.xml' } }],
      parts: { 'footer1.xml': [para('Agreement')], 'footer2.xml': [para('Schedule')] },
    };
    const { pkg } = await compare(original, revised);
    // The first section's footer did not change: the original's part, untouched.
    expect(pkg.part('word/footer1.xml')).toBe(partsPackage(original).part('word/footer1.xml'));
    const refs = [...pkg.part('word/document.xml')!.matchAll(/<w:footerReference [^>]*r:id="([^"]*)"/g)].map((m) => m[1]);
    expect(refs[0]).toBe('rId10');
    expect(refs[1]).toMatch(/^rIdDp/);
    referencesResolve(pkg);
  });

  it('a header table keeps its structure, its changed cell tracked', async () => {
    config = { excludeHeaders: false };
    const tbl = (ref: string): string =>
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${para('Sample Holdings')}</w:tc>` +
      `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${para(ref)}</w:tc></w:tr></w:tbl>`;
    const spec = (ref: string): PackageSpec => ({
      body: ['Clause one.'], sections: [{ headers: { default: 'header1.xml' } }],
      parts: { 'header1.xml': [tbl(ref), para('')] },
    });
    const { pkg } = await compare(spec('Ref: MH/001'), spec('Ref: MH/002'));
    const header = pkg.part('word/header1.xml')!;
    expect(header.match(/<w:tbl>/g)).toHaveLength(1);
    expect(header).toMatch(/<w:tc>(?:(?!<\/w:tc>)[\s\S])*<w:ins /);
    partReconstructs(header, partsPackage(spec('Ref: MH/001')).part('word/header1.xml')!, partsPackage(spec('Ref: MH/002')).part('word/header1.xml')!);
  });

  it('an added header’s own relationships and the image they target are copied, a clashing name renamed', async () => {
    config = { excludeHeaders: false };
    const original = partsPackage({ body: ['Clause one.'] });
    original.setPartBytes('word/media/image1.png', new Uint8Array([1, 2, 3]));
    const revised = partsPackage({
      body: ['Clause one.'], sections: [{ headers: { default: 'header1.xml' } }],
      parts: { 'header1.xml': ['<w:p><w:r><w:drawing><a:blip xmlns:a="a" r:embed="rIdLogo"/></w:drawing></w:r></w:p>'] },
    });
    revised.setPart('word/_rels/header1.xml.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rIdLogo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>');
    revised.setPartBytes('word/media/image1.png', new Uint8Array([9, 9, 9]));
    revised.setPart('[Content_Types].xml', revised.part('[Content_Types].xml')!.replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>'));
    const driver = createCompareDriver();
    driver.uploadDocx('original', new File([new Uint8Array(original.toUint8Array())], 'o.docx'));
    driver.uploadDocx('revised', new File([new Uint8Array(revised.toUint8Array())], 'r.docx'));
    await driver.runCompare();
    const pkg = written!;
    const rels = parseRelationships(pkg.part('word/_rels/document.xml.rels'));
    const header = resolveTarget('word', rels.find((r) => r.type === 'header')!.target);
    const own = pkg.part(header.replace(/^word\//, 'word/_rels/') + '.rels')!;
    expect(own).toContain('Target="media/image1-dp1.png"');
    expect([...pkg.partBytes!('word/media/image1-dp1.png')!]).toEqual([9, 9, 9]);
    expect([...pkg.partBytes!('word/media/image1.png')!]).toEqual([1, 2, 3]);
    expect(pkg.part('[Content_Types].xml')).toContain('Extension="png"');
  });
});
