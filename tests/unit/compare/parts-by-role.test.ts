/**
 * Headers and footers are found and matched by ROLE.
 *
 * A role is the section whose `<w:sectPr>` references the part plus the
 * reference's `w:type`. Word numbers `header1.xml`, `header2.xml` … as it saves,
 * so the same file name can be a different header on the other side; matching by
 * name compared two different things. And a part one side has and the other
 * does not was skipped by a bare `continue` — a suppression (ADR-0008).
 */
import { describe, it, expect } from 'vitest';

import { extractParts, sectionSpans, sectionReferences } from '@features/compare/io/parts.js';
import { pairRoles } from '@features/compare/parts-diff.js';
import { getCompareConfig, getActiveProfile } from '@shared/config/index.js';
import type { CompareConfig } from '@shared/config/index.js';

import { partsPackage, para, compareHeadersFooters } from '../../fixtures/parts-docx.js';

const CONFIG: CompareConfig = { ...getCompareConfig(getActiveProfile()), excludeHeaders: false, excludeFooters: false };
const AUTHOR = 'Compare';
const DATE = '2026-10-26T00:00:00Z';

const changed = (r: Awaited<ReturnType<typeof compareHeadersFooters>>): number =>
  r.ops.filter((o) => o.type !== 'equal').length;

describe('extractParts finds headers and footers through the relationships', () => {
  it('keys each part by its role, carrying its path', async () => {
    const parts = extractParts(partsPackage({
      body: ['The Buyer shall pay.'],
      sections: [{ headers: { default: 'header1.xml', first: 'header2.xml' }, footers: { default: 'footer1.xml' } }],
      parts: {
        'header1.xml': [para('Running header')],
        'header2.xml': [para('Cover header')],
        'footer1.xml': [para('Footer stamp')],
      },
    }));
    expect([...parts.headers.keys()]).toEqual(['header:s0:default', 'header:s0:first']);
    expect(parts.headers.get('header:s0:first')!.path).toBe('word/header2.xml');
    expect(parts.footers.get('footer:s0:default')!.xml).toContain('Footer stamp');
    expect(parts.sectionCount).toBe(1);
    expect(parts.roles.map((r) => r.rId)).toEqual(['rId10', 'rId11', 'rId12']);
  });

  it('does not take a header file no section references', async () => {
    const parts = extractParts(partsPackage({
      body: ['The Buyer shall pay.'],
      parts: { 'header1.xml': [para('Orphan header')] },
    }));
    expect(parts.headers.size).toBe(0);
  });

  it('numbers sections by their sectPr, skipping the old properties a sectPrChange carries', async () => {
    const xml =
      '<w:body><w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rId7"/>' +
      '<w:sectPrChange w:id="1"><w:sectPr><w:pgSz w:w="1"/></w:sectPr></w:sectPrChange></w:sectPr></w:pPr></w:p>' +
      '<w:sectPr/></w:body>';
    const spans = sectionSpans(xml);
    expect(spans).toHaveLength(2);
    expect(spans[0]!.xml.endsWith('</w:sectPrChange></w:sectPr>')).toBe(true);
    expect(sectionReferences(spans[0]!.xml)).toEqual([
      { kind: 'header', type: 'default', rId: 'rId7', xml: '<w:headerReference w:type="default" r:id="rId7"/>' },
    ]);
    expect(spans[1]!.xml).toBe('<w:sectPr/>');
  });
});

describe('parts are matched by role, never by file name', () => {
  it('two packages whose header files are swapped between roles show no change', async () => {
    const original = extractParts(partsPackage({
      body: ['The Buyer shall pay.'],
      sections: [{ headers: { default: 'header1.xml', first: 'header2.xml' } }],
      parts: { 'header1.xml': [para('Running header')], 'header2.xml': [para('Cover header')] },
    }));
    const revised = extractParts(partsPackage({
      body: ['The Buyer shall pay.'],
      sections: [{ headers: { default: 'header2.xml', first: 'header1.xml' } }],
      parts: { 'header1.xml': [para('Cover header')], 'header2.xml': [para('Running header')] },
    }));
    const result = await compareHeadersFooters(original, revised, CONFIG, AUTHOR, DATE);
    expect(changed(result)).toBe(0);
    expect(pairRoles(original, revised, 'header').map((x) => [x.roleId, x.original?.xml === x.revised?.xml]))
      .toEqual([['header:s0:default', true], ['header:s0:first', true]]);
    expect(pairRoles(original, revised, 'header').map((x) => [x.original?.path, x.revised?.path]))
      .toEqual([['word/header1.xml', 'word/header2.xml'], ['word/header2.xml', 'word/header1.xml']]);
  });

  it('a header role on one side only is reported, its paragraphs inserted', async () => {
    const original = extractParts(partsPackage({
      body: ['The Buyer shall pay.'],
      sections: [{ headers: { default: 'header1.xml' } }],
      parts: { 'header1.xml': [para('Running header')] },
    }));
    const revised = extractParts(partsPackage({
      body: ['The Buyer shall pay.'],
      sections: [{ headers: { default: 'header1.xml' }, footers: { default: 'footer1.xml' } }],
      parts: { 'header1.xml': [para('Running header')], 'footer1.xml': [para('Draft for discussion')] },
    }));
    const result = await compareHeadersFooters(original, revised, CONFIG, AUTHOR, DATE);
    expect(result.footerChanges).toBe(1);
    const footer = result.parts.find((s) => s.partId === 'footer:s0:default')!;
    expect(footer.originalParas).toEqual([]);
    expect(footer.ops.map((o) => o.type)).toEqual(['insert']);
  });

  it('a footer role removed is reported, its paragraphs deleted', async () => {
    const original = extractParts(partsPackage({
      body: ['The Buyer shall pay.'],
      sections: [{ footers: { default: 'footer1.xml' } }],
      parts: { 'footer1.xml': [para('Draft for discussion'), para('Subject to contract')] },
    }));
    const revised = extractParts(partsPackage({ body: ['The Buyer shall pay.'] }));
    const result = await compareHeadersFooters(original, revised, CONFIG, AUTHOR, DATE);
    expect(result.footerChanges).toBe(2);
    expect(result.parts[0]!.ops.map((o) => o.type)).toEqual(['delete', 'delete']);
  });

  it('a section that gained its own copy of the header it inherited shows no change', async () => {
    const original = extractParts(partsPackage({
      body: ['Clause one.', '§', 'Schedule one.'],
      sections: [{ headers: { default: 'header1.xml' } }, {}],
      parts: { 'header1.xml': [para('Running header')] },
    }));
    const revised = extractParts(partsPackage({
      body: ['Clause one.', '§', 'Schedule one.'],
      sections: [{ headers: { default: 'header1.xml' } }, { headers: { default: 'header2.xml' } }],
      parts: { 'header1.xml': [para('Running header')], 'header2.xml': [para('Running header')] },
    }));
    const pairs = pairRoles(original, revised, 'header');
    expect(pairs.map((x) => [x.roleId, x.original?.path, x.revised?.path])).toEqual([
      ['header:s0:default', 'word/header1.xml', 'word/header1.xml'],
      ['header:s1:default', 'word/header1.xml', 'word/header2.xml'],
    ]);
    expect(changed(await compareHeadersFooters(original, revised, CONFIG, AUTHOR, DATE))).toBe(0);
  });

  it('every part op id is prefixed with its role, so it cannot collide with a body op', async () => {
    const original = extractParts(partsPackage({
      body: ['x'], sections: [{ headers: { default: 'header1.xml' } }], parts: { 'header1.xml': [para('Alpha')] },
    }));
    const revised = extractParts(partsPackage({
      body: ['x'], sections: [{ headers: { default: 'header1.xml' } }], parts: { 'header1.xml': [para('Beta')] },
    }));
    const ids = (await compareHeadersFooters(original, revised, CONFIG, AUTHOR, DATE)).ops.map((o) => o.id);
    expect(ids.every((id) => id.startsWith('header:s0:default/'))).toBe(true);
  });
});
