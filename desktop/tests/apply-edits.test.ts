import { describe, it, expect } from 'vitest';
import { PizZipBundle } from '@shared/ooxml/bundle.js';
import { applyEdits } from '../src/renderer/src/edits/apply-edits.js';
import { makeDocxBytes } from './helpers/make-docx.js';

/**
 * The refusal behaviour is the point of this module. A model that quotes text
 * slightly wrong, or quotes a phrase that appears twice, must get a reason back
 * — never a guess at which clause it meant.
 */

function textOf(bytes: Uint8Array): string {
  const xml = new PizZipBundle(bytes).part('word/document.xml') ?? '';
  return [...xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)]
    .map((match) => match[1])
    .join('|');
}

const DOC = [
  'The Purchase Price is GBP 1,000,000.',
  'The Seller shall indemnify the Buyer in full.',
  'Notices must be given in writing.',
];

describe('applyEdits', () => {
  it('replaces text that occurs exactly once', () => {
    const result = applyEdits(makeDocxBytes(DOC), [
      { find: 'GBP 1,000,000', replace: 'GBP 1,250,000', reason: 'agreed uplift' },
    ]);
    expect(result.refused).toEqual([]);
    expect(result.applied).toHaveLength(1);
    expect(textOf(result.bytes)).toContain('GBP 1,250,000');
    expect(textOf(result.bytes)).not.toContain('1,000,000');
  });

  it('leaves every other paragraph untouched', () => {
    const result = applyEdits(makeDocxBytes(DOC), [
      { find: 'GBP 1,000,000', replace: 'GBP 1,250,000' },
    ]);
    const text = textOf(result.bytes);
    expect(text).toContain('Notices must be given in writing.');
    expect(text).toContain('The Seller shall indemnify the Buyer in full.');
  });

  it('REFUSES text it cannot find, with a reason', () => {
    const result = applyEdits(makeDocxBytes(DOC), [
      { find: 'the purchase price is GBP 1,000,000', replace: 'x' },
    ]);
    expect(result.applied).toEqual([]);
    expect(result.refused[0]?.refusal).toMatch(/does not appear/);
  });

  it('REFUSES ambiguous text rather than picking one', () => {
    const result = applyEdits(makeDocxBytes([...DOC, 'Notices must be given in writing.']), [
      { find: 'Notices must be given in writing.', replace: 'Notices may be given by email.' },
    ]);
    expect(result.applied).toEqual([]);
    expect(result.refused[0]?.refusal).toMatch(/appears 2 times/);
    expect(textOf(result.bytes)).not.toContain('by email');
  });

  it('refuses an empty find rather than matching everywhere', () => {
    const result = applyEdits(makeDocxBytes(DOC), [{ find: '   ', replace: 'x' }]);
    expect(result.refused[0]?.refusal).toMatch(/empty/);
  });

  it('deletes when the replacement is empty', () => {
    const result = applyEdits(makeDocxBytes(DOC), [
      { find: ' in full', replace: '' },
    ]);
    expect(result.applied).toHaveLength(1);
    expect(textOf(result.bytes)).toContain('The Seller shall indemnify the Buyer.');
  });

  it('applies several edits, and one refusal does not block the others', () => {
    const result = applyEdits(makeDocxBytes(DOC), [
      { find: 'GBP 1,000,000', replace: 'GBP 2,000,000' },
      { find: 'a clause that is not there', replace: 'x' },
      { find: 'in writing', replace: 'in writing or by email' },
    ]);
    expect(result.applied).toHaveLength(2);
    expect(result.refused).toHaveLength(1);
    const text = textOf(result.bytes);
    expect(text).toContain('GBP 2,000,000');
    expect(text).toContain('in writing or by email');
  });

  it('produces a document the engine can still read', () => {
    const result = applyEdits(makeDocxBytes(DOC), [
      { find: 'GBP 1,000,000', replace: 'GBP 1,250,000' },
    ]);
    const bundle = new PizZipBundle(result.bytes);
    expect(bundle.hasPart('word/document.xml')).toBe(true);
    expect(bundle.part('word/document.xml')).toMatch(/<\/w:document>/);
  });

  it('escapes XML in the replacement rather than injecting markup', () => {
    const result = applyEdits(makeDocxBytes(DOC), [
      { find: 'in writing', replace: 'in writing <w:br/> & signed' },
    ]);
    expect(result.applied).toHaveLength(1);
    const xml = new PizZipBundle(result.bytes).part('word/document.xml') ?? '';
    expect(xml).toContain('&lt;w:br/&gt; &amp; signed');
    expect(xml).not.toContain('<w:br/>');
  });

  it('writes no revision markup — tracked changes come from the compare engine', () => {
    const result = applyEdits(makeDocxBytes(DOC), [
      { find: 'GBP 1,000,000', replace: 'GBP 1,250,000' },
    ]);
    const xml = new PizZipBundle(result.bytes).part('word/document.xml') ?? '';
    expect(xml).not.toContain('<w:ins');
    expect(xml).not.toContain('<w:del');
  });

  it('keeps trailing whitespace by adding xml:space when it matters', () => {
    const result = applyEdits(makeDocxBytes(['Payment is due on Completion.']), [
      { find: 'Completion.', replace: 'Completion, ' },
    ]);
    const xml = new PizZipBundle(result.bytes).part('word/document.xml') ?? '';
    expect(xml).toMatch(/xml:space="preserve"/);
  });
});
