/**
 * Pairing sketches (ADR-0037; spec §6.2).
 *
 * A sketch is the only thing pairing ever reads, so the tests here are about
 * the two properties the rest of the feature leans on: that the sketch
 * *discriminates* (two versions of one agreement share most of their hashes;
 * two different agreements share almost none) and that it is *plain data* a
 * human could check, because ADR-0008 forbids an opaque judgement about a
 * document.
 */
import { describe, it, expect } from 'vitest';

import { getCompareConfig, UK_PROFILE } from '@shared/config/index.js';
import { synthesiseDocx } from '@features/compare/io/text/synthesise-docx.js';
import {
  buildPairingSketch,
  isSketchWorthy,
  type BundleSide,
} from '@features/compare/batch/pairing/sketch.js';
import { jaccard } from '@features/compare/batch/pairing/signals.js';
import { PAIRING_MIN_SKETCH_TOKENS } from '@features/compare/batch/constants.js';
import type { ParaRecord } from '@features/compare/model.js';

const CONFIG = getCompareConfig(UK_PROFILE);

/** A body long enough that paragraph hashes actually discriminate. */
const SPA = [
  'This Share Purchase Agreement is made between the Seller and the Buyer.',
  'The Seller shall sell and the Buyer shall purchase the Sale Shares.',
  'The Purchase Price shall be paid in full on the Completion Date.',
  'Completion shall take place at the offices of the Seller’s solicitors.',
  'The Seller warrants that it has good title to the Sale Shares.',
  'This Agreement is governed by the laws of England and Wales.',
].join('\n');

const TAX_DEED = [
  'This Deed of Tax Covenant is made between the Covenantor and the Covenantee.',
  'The Covenantor covenants to pay an amount equal to any Tax Liability.',
  'A claim under this Deed must be notified within seven years.',
  'Payment shall be made in cleared funds within five Business Days.',
  'Nothing in this Deed limits any other remedy available at law.',
  'This Deed is governed by the laws of Singapore.',
].join('\n');

function bytesOf(text: string): ArrayBuffer {
  const raw = new Uint8Array(synthesiseDocx(text).toUint8Array());
  return raw.buffer.slice(raw.byteOffset, raw.byteOffset + raw.byteLength) as ArrayBuffer;
}

function sketch(docId: string, text: string, side: BundleSide = 'a', fileName = `${docId}.docx`) {
  return buildPairingSketch({
    docId,
    side,
    fileName,
    bytes: bytesOf(text),
    config: CONFIG,
    profile: UK_PROFILE,
    flatten: 'acceptAll',
  });
}

describe('isSketchWorthy — boilerplate exclusion', () => {
  const record = (normText: string): ParaRecord =>
    ({ normText }) as unknown as ParaRecord;

  it('excludes paragraphs below the content-token floor', () => {
    // Empty paragraphs, one-word headings and numbering stubs are identical in
    // every document in a bundle; left in, they pull every Jaccard score toward
    // every other and flatten the distinction pairing exists to make.
    expect(isSketchWorthy(record(''), CONFIG)).toBe(false);
    expect(isSketchWorthy(record('Definitions'), CONFIG)).toBe(false);
    expect(isSketchWorthy(record('1.'), CONFIG)).toBe(false);
  });

  it('includes a paragraph at or above the floor', () => {
    const atFloor = Array.from({ length: PAIRING_MIN_SKETCH_TOKENS }, (_, i) => `word${i}`).join(' ');
    expect(isSketchWorthy(record(atFloor), CONFIG)).toBe(true);
    expect(isSketchWorthy(record(SPA.split('\n')[0] ?? ''), CONFIG)).toBe(true);
  });
});

describe('buildPairingSketch', () => {
  it('is plain data — every field is inspectable, none of it is a verdict', () => {
    const s = sketch('spa', SPA, 'a', 'Project Falcon SPA v6.docx');

    expect(s.docId).toBe('spa');
    expect(s.side).toBe('a');
    expect(s.sourceKind).toBe('docx');
    expect(s.fileName).toBe('Project Falcon SPA v6.docx');
    expect(s.byteLength).toBeGreaterThan(0);
    expect(s.paraHashes.size).toBeGreaterThan(0);
    expect(s.quarantinedCount).toBe(0);

    // Sets of strings and numbers only. Nothing here is a score, a band or a
    // decision — those are computed downstream from these, in the open.
    for (const hash of s.paraHashes) expect(typeof hash).toBe('string');
    for (const term of s.definedTerms) expect(typeof term).toBe('string');
    for (const name of s.partyNames) expect(typeof name).toBe('string');
  });

  it('excludes boilerplate paragraphs from the content signal', () => {
    // Six substantive paragraphs plus three stubs; only the six earn a hash.
    const withStubs = `Definitions\n1.\n\n${SPA}`;
    expect(sketch('stubs', withStubs).paraHashes.size).toBe(sketch('clean', SPA).paraHashes.size);
  });

  it('keeps the hashes and DROPS the records', () => {
    // Spec §4.3. The sketch used to carry `paras` so a confirmed pairing could
    // run without re-reading the file — and at the §4.4 cap of fifty documents
    // a side that is a hundred document-sized `ParaRecord[]` live from the
    // moment a pairing is proposed until the driver is discarded, which is the
    // out-of-memory shape §4.3 opens by naming. The run preprocesses each
    // pair's two documents in the lane that needs them instead.
    const s = sketch('spa', SPA);
    expect(Object.keys(s)).not.toContain('paras');
    expect(s.paraHashes.size).toBeGreaterThan(0);
    expect([...s.paraHashes].every(h => typeof h === 'string')).toBe(true);
  });

  it('discriminates two versions of one agreement from two different agreements', () => {
    const v6 = sketch('v6', SPA);
    const v7 = sketch('v7', SPA.replace('paid in full on the Completion Date', 'paid in cleared funds on Completion'));
    const deed = sketch('deed', TAX_DEED);

    // The property the whole of pairing rests on.
    expect(jaccard(v6.paraHashes, v7.paraHashes)).toBeGreaterThan(0.6);
    expect(jaccard(v6.paraHashes, deed.paraHashes)).toBe(0);
    expect(jaccard(v6.paraHashes, v7.paraHashes))
      .toBeGreaterThan(jaccard(v6.paraHashes, deed.paraHashes));
  });

  it('gives an identical document an identical hash set, whatever it is called', () => {
    // Filename is a separate, secondary signal; it must not leak into content.
    const a = sketch('a', SPA, 'a', 'SPA execution copy.docx');
    const b = sketch('b', SPA, 'b', 'final version FOR SIGNING.docx');
    expect([...a.paraHashes].sort()).toEqual([...b.paraHashes].sort());
  });

  it('survives a document whose cast cannot be harvested', () => {
    // Best-effort by design (spec §6.3): the cast is a tie-breaker, so losing
    // it must never cost a pairing that content and filename can decide.
    const s = sketch('bare', 'The quick brown fox jumped over the lazy dog again.');
    expect(s.paraHashes.size).toBeGreaterThan(0);
    expect(s.definedTerms.size).toBe(0);
  });
});
