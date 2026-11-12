/**
 * ADR-0045 § Amendment 1 — the zero-text objects of a paragraph, and how the
 * change list and the report print a change to them.
 *
 * `zeroTextObjects` decides what is compared; every sort it reads and every
 * place it computes is pinned here, directly, so a regression shows as a wrong
 * list rather than as a missing op three stages later. Synthetic XML only.
 */
import { describe, it, expect } from 'vitest';
import { zeroTextObjects } from '@features/compare/preprocess/inline-units.js';
import { objectEdits, textWithObjects } from '@features/compare/diff/object-edits.js';
import { previewText } from '@features/compare/components/ChangeList.js';
import type { DiffOp, PropertyChange } from '@features/compare/model.js';
import {
  para, run, footnoteRef, endnoteRef, inlinePicture, embeddedObject, textBox, alternateTextBox, hyperlink,
} from '../../fixtures/inline-objects-docx.js';

const legacyPicture = '<w:r><w:pict><v:shape id="_x0000_s1026"><v:imagedata r:id="rId7"/></v:shape></w:pict></w:r>';

describe('zeroTextObjects: the sorts', () => {
  it('reads a footnote and an endnote reference, a picture, an embedded object, in document order', () => {
    const xml = para(run('A') + footnoteRef() + run('B') + endnoteRef() + inlinePicture() + run('C') + embeddedObject());
    expect(zeroTextObjects(xml, '').map((x) => x.sort)).toEqual([
      'footnote reference', 'endnote reference', 'picture', 'embedded object',
    ]);
  });

  it('reads a legacy VML picture as a picture', () => {
    expect(zeroTextObjects(para(run('A') + legacyPicture), '')).toEqual([{ sort: 'picture', at: 1 }]);
  });

  it('reads a text box as a text box, and a DrawingML choice with its VML fallback ONCE', () => {
    expect(zeroTextObjects(para(run('A ') + textBox('Note')), '').map((x) => x.sort)).toEqual(['text box']);
    expect(zeroTextObjects(para(run('A ') + alternateTextBox('Note')), '').map((x) => x.sort)).toEqual(['text box']);
  });

  it('reads a footnote nested inside a link', () => {
    const link = `<w:hyperlink r:id="rId9"><w:r><w:t>the Seller</w:t></w:r>${footnoteRef()}</w:hyperlink>`;
    expect(zeroTextObjects(para(run('Pay ') + link), '')).toEqual([{ sort: 'footnote reference', at: 14 }]);
  });

  it('reads nothing from a paragraph without such a tag (the pre-filter), or from a link or bookmark', () => {
    expect(zeroTextObjects(para(run('Plain words.')), '')).toEqual([]);
    expect(zeroTextObjects(para(run('See ') + hyperlink('the portal')), '')).toEqual([]);
  });

  it('never reads which footnote or image it is', () => {
    const a = zeroTextObjects(para(run('A') + footnoteRef('1') + inlinePicture('rId20')), '');
    const b = zeroTextObjects(para(run('A') + footnoteRef('7') + inlinePicture('rId99')), '');
    expect(a).toEqual(b);
  });
});

describe('zeroTextObjects: the places', () => {
  it('counts characters of the normalised text before it', () => {
    expect(zeroTextObjects(para(run('The Seller') + footnoteRef() + run(' shall pay.')), '')).toEqual([
      { sort: 'footnote reference', at: 10 },
    ]);
  });

  it('is not moved by a whitespace run written differently', () => {
    const one = zeroTextObjects(para(run('The Seller') + footnoteRef() + run(' shall pay.')), '');
    const two = zeroTextObjects(para(run('The   Seller') + footnoteRef() + run('  shall pay.')), '');
    expect(two).toEqual(one);
  });

  it('is not moved by trailing space after the object', () => {
    expect(zeroTextObjects(para(run('Pay.') + footnoteRef() + run('  ')), '')).toEqual(
      zeroTextObjects(para(run('Pay.') + footnoteRef()), ''),
    );
  });

  it('counts after a manual number from the first character after it, so renumbering does not move it', () => {
    expect(zeroTextObjects(para(run('9. ') + footnoteRef() + run('Pay.')), '9. ')).toEqual([{ sort: 'footnote reference', at: 0 }]);
    expect(zeroTextObjects(para(run('10. ') + footnoteRef() + run('Pay.')), '10. ')).toEqual([{ sort: 'footnote reference', at: 0 }]);
  });

  it('places an object inside a text box exactly, and says which box paragraph holds it', () => {
    const box = (inner: string) =>
      '<w:r><w:drawing><wp:anchor><a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent>' +
      `<w:p>${inner}</w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
    const start = zeroTextObjects(para(run('See ') + box(footnoteRef() + run('Plan'))), '');
    const second = zeroTextObjects(para(run('See ') + box(run('P') + footnoteRef() + run('lan'))), '');
    const outside = zeroTextObjects(para(run('See ') + footnoteRef() + box(run('Plan'))), '');
    expect(start[1]).toMatchObject({ sort: 'footnote reference', within: 1 });
    expect(second[1]!.at).toBe(start[1]!.at + 1);
    expect(outside[0]).toEqual({ sort: 'footnote reference', at: start[1]!.at });
    expect(outside[0]).not.toEqual(start[1]); // same count, different place
  });

  it('keeps the place of an object inside an equation or box that has no text', () => {
    const empty = '<w:r><w:drawing><wp:anchor><a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent>' +
      `<w:p>${inlinePicture()}</w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
    expect(zeroTextObjects(para(run('Ab') + empty + run('cd')), '').map((x) => x.at)).toEqual([2, 2]);
  });

  it('codes a place in the manual number: -2 before it, -1 inside it', () => {
    expect(zeroTextObjects(para(footnoteRef() + run('9. Pay.')), '9. ')).toEqual([{ sort: 'footnote reference', at: -2 }]);
    expect(zeroTextObjects(para(run('9.') + footnoteRef() + run(' Pay.')), '9. ')).toEqual([{ sort: 'footnote reference', at: -1 }]);
  });
});

describe('the surfaces', () => {
  const fn = (at: number) => ({ sort: 'footnote reference' as const, at });
  const pic = (at: number) => ({ sort: 'picture' as const, at });
  type Objects = NonNullable<PropertyChange['from']['objects']>;
  const change = (from: Objects, to: Objects, extra: Partial<PropertyChange> = {}): PropertyChange => ({
    styleId: false, numberingIdentity: false, numberLabel: false, objects: true,
    from: { styleId: null, numberLabel: null, ilvl: null, objects: from },
    to: { styleId: null, numberLabel: null, ilvl: null, objects: to },
    ...extra,
  });

  it('objectEdits aligns on sort and place: one removed of three is the only edit, a move is both', () => {
    expect(objectEdits(change([fn(3), fn(9), fn(20)], [fn(3), fn(20)]))).toEqual({ removed: [fn(9)], added: [] });
    expect(objectEdits(change([fn(3)], [fn(9)]))).toEqual({ removed: [fn(3)], added: [fn(9)] });
    expect(objectEdits(change([], [pic(0)]))).toEqual({ removed: [], added: [pic(0)] });
  });

  it('textWithObjects writes each object in at its place, a place in the manual number at the start', () => {
    expect(textWithObjects('The Seller shall pay.', [fn(10), pic(21)])).toBe('The Seller[footnote reference] shall pay.[picture]');
    expect(textWithObjects('Pay.', [fn(-2)])).toBe('[footnote reference]Pay.');
    expect(textWithObjects('', [pic(0)])).toBe('[picture]');
  });

  it('the row names the objects first, then the number that moved with them', () => {
    const op = {
      id: 'op-1', type: 'property', origRange: { start: 0, end: 0 }, revRange: { start: 0, end: 0 }, partId: 'document',
      classification: 'substantiveText', classes: ['substantiveText', 'formatting'], author: 'A', date: 'D',
      propertyChange: {
        ...change([fn(4)], []),
        numberLabel: true,
        from: { styleId: null, numberLabel: '(c)', ilvl: 1, objects: [fn(4)] },
        to: { styleId: null, numberLabel: '(d)', ilvl: 1, objects: [] },
      },
    } as DiffOp;
    expect(previewText(op)).toBe('-[footnote reference] · (c) → (d)');
  });
});
