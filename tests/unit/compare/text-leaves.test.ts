/**
 * `textLeaves` names the source of every character
 * `collectText` reads, one for one, so an edited paragraph can be rebuilt from
 * its own runs. The side map refuses the source rebuild when the two disagree
 * (a fast path ahead of the rebuild's own dry-run check); this pins that they
 * agree on every shape the market's drafting uses.
 */
import { describe, it, expect } from 'vitest';

import { childrenOf } from '@shared/ooxml/po-xml.js';
import { paragraphNode, paragraphText, textLeaves } from '@features/compare/preprocess/inline-units.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
const r = (t: string): string => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const box = (inner: string): string =>
  `<w:r><mc:AlternateContent><mc:Choice><w:drawing><w:txbxContent>${inner}</w:txbxContent></w:drawing></mc:Choice><mc:Fallback><w:pict><w:txbxContent>${inner}</w:txbxContent></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;

const SHAPES: [string, string][] = [
  ['plain runs', r('The Seller ') + r('shall deliver.')],
  ['a tab and a break', '<w:r><w:t>(a)</w:t><w:tab/><w:t>Price</w:t><w:br/><w:t>payable</w:t></w:r>'],
  ['a non-breaking hyphen inside a run and alone', r('co') + '<w:r><w:noBreakHyphen/></w:r>' + '<w:r><w:t>re</w:t><w:softHyphen/><w:t>sign</w:t></w:r>'],
  ['a hyperlink, a simple field and a content control', `<w:hyperlink w:history="1">${r('the Website')}</w:hyperlink><w:fldSimple w:instr=" PAGE ">${r('3')}</w:fldSimple><w:sdt><w:sdtContent>${r('[Date]')}</w:sdtContent></w:sdt>`],
  ['an equation', r('rate ') + `<m:oMath><m:r><m:t>r=3%</m:t></m:r></m:oMath>` + r(' per annum')],
  ['a text box (choice read, fallback not)', r('See ') + box(`<w:p>${r('Box text')}</w:p>`) + r(' below.')],
  ['an empty text box', r('See') + box('<w:p></w:p>') + r(' below.')],
  ['a character outside the Basic Multilingual Plane', r(`The ${String.fromCodePoint(0x20bb7)} family.`)],
];

describe('textLeaves reads collectText\'s characters one for one', () => {
  for (const [what, inner] of SHAPES) {
    it(what, () => {
      const p = paragraphNode(`<w:p ${W}>${inner}</w:p>`)!;
      const leaves = textLeaves(childrenOf(p));
      expect(leaves.map((l) => l?.char ?? ' ').join('')).toBe(paragraphText(p));
      expect(leaves).toHaveLength(paragraphText(p).length); // one per UTF-16 code unit, as the text is indexed
    });
  }

  it('silent content travels with a character: inside its run, or as its own run', () => {
    const p = paragraphNode(`<w:p ${W}>${r('co')}<w:r><w:noBreakHyphen/></w:r><w:r><w:t>re</w:t><w:softHyphen/><w:t>sign</w:t></w:r></w:p>`)!;
    const leaves = textLeaves(childrenOf(p));
    expect(leaves[2]!.runsBefore).toHaveLength(1); // the hyphen's own run, before "r"
    expect(leaves[4]!.before).toHaveLength(1); // the soft hyphen, before "s"
  });

  it('a run of only silent content at the end of a paragraph stays after its last character', () => {
    const p = paragraphNode(`<w:p ${W}>${r('re')}<w:r><w:noBreakHyphen/></w:r></w:p>`)!;
    expect(textLeaves(childrenOf(p))[1]!.runsAfter).toHaveLength(1);
  });
});
