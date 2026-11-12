import { describe, it, expect } from 'vitest';
import { buildPairing } from '../src/renderer/src/pairing/build-pairing.js';
import { makeDocxBytes, SAMPLE_ORIGINAL, SAMPLE_REVISED } from './helpers/make-docx.js';

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const copy = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(copy).set(bytes);
  return copy;
}

describe('pairing proposal', () => {
  it('gives every document exactly one row, even with nothing to pair against', () => {
    const result = buildPairing(
      [
        {
          docId: 'a1',
          side: 'a',
          fileName: 'only-original.docx',
          bytes: toArrayBuffer(makeDocxBytes(SAMPLE_ORIGINAL)),
        },
      ],
      'UK',
      '2026-01-01T00:00:00.000Z',
    );

    // An unmatched document is a real answer — an added or removed document —
    // not a failure, and it must still appear.
    expect(result.proposal.rows).toHaveLength(1);
    expect(result.proposal.rows[0]?.band).toBe('unmatched');
    expect(result.proposal.totals.unmatchedA).toBe(1);
  });

  it('pairs an obvious 1:1 and counts it', () => {
    const result = buildPairing(
      [
        {
          docId: 'a1',
          side: 'a',
          fileName: 'services.docx',
          bytes: toArrayBuffer(makeDocxBytes(SAMPLE_ORIGINAL)),
        },
        {
          docId: 'b1',
          side: 'b',
          fileName: 'services-v2.docx',
          bytes: toArrayBuffer(makeDocxBytes(SAMPLE_REVISED)),
        },
      ],
      'UK',
      '2026-01-01T00:00:00.000Z',
    );

    expect(result.proposal.totals.pairs).toBe(1);
    const row = result.proposal.rows.find((r) => r.a !== null && r.b !== null);
    expect(row?.a).toBe('a1');
    expect(row?.b).toBe('b1');
  });
});
