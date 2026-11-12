/**
 * The many-to-many runner (ADR-0037; spec §6.7).
 *
 * Two properties carry this module, and neither is visible anywhere else:
 *
 *  - **It takes confirmed pairs and nothing else.** There is no path from two
 *    bundles to a run inside the runner. A runner that could pair for itself
 *    would be a runner that could run an unconfirmed pairing, and an unconfirmed
 *    pairing produces a complete, plausible, worthless redline.
 *  - **No document is preprocessed twice.** The pairing sketches already built
 *    the records while proposing the pairing; re-deriving them for the run would
 *    double the most expensive part of a bundle comparison.
 */
import { describe, it, expect } from 'vitest';

import { getCompareConfig, UK_PROFILE } from '@shared/config/index.js';
import { synthesiseDocx } from '@features/compare/io/text/synthesise-docx.js';
import {
  runManyToManyBatch,
  CROSS_FORMAT_REASON,
  type ConfirmedPairInput,
  type PairRecords,
} from '@features/compare/batch/runner.js';
import type { BatchRow } from '@features/compare/batch/model.js';
import {
  createBatchPool,
  createInProcessWorker,
  type BatchPool,
} from '@features/compare/batch/worker/index.js';
import { buildPairingSketch, preprocessBundleDocument } from '@features/compare/batch/pairing/sketch.js';
import { proposePairing } from '@features/compare/batch/pairing/propose.js';
import {
  MAX_BATCH_TOTAL_BYTES,
  MAX_BUNDLE_DOCUMENTS,
  MAX_DOCX_BYTES,
} from '@features/compare/batch/limits.js';
import type { PairingSketch, BundleSide } from '@features/compare/batch/pairing/sketch.js';
import { PizZipBundle } from '@shared/ooxml/bundle.js';

const CONFIG = getCompareConfig(UK_PROFILE);
const DATE = '2026-09-10T09:00:00.000Z';

const SPA = [
  'This Share Purchase Agreement is made between the Seller and the Buyer.',
  'The Seller shall sell and the Buyer shall purchase the Sale Shares.',
  'The Purchase Price shall be paid in full on the Completion Date.',
  'The Seller warrants that it has good title to the Sale Shares.',
  'This Agreement is governed by the laws of England and Wales.',
].join('\n');

const DEED = [
  'This Deed of Tax Covenant is made between the Covenantor and the Covenantee.',
  'The Covenantor covenants to pay an amount equal to any Tax Liability.',
  'A claim under this Deed must be notified within seven years.',
  'Payment shall be made in cleared funds within five Business Days.',
  'This Deed is governed by the laws of Singapore.',
].join('\n');

function bytesOf(text: string): ArrayBuffer {
  const raw = new Uint8Array(synthesiseDocx(text).toUint8Array());
  return raw.buffer.slice(raw.byteOffset, raw.byteOffset + raw.byteLength) as ArrayBuffer;
}

function sketchOf(docId: string, fileName: string, text: string, side: BundleSide): PairingSketch {
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

/**
 * One side's half of `PairRecords`. Two exact `Pick`s rather than one function
 * returning a union: the union's branches each carry `?: never` for the other
 * side's keys, which does not spread to `PairRecords` — and the obvious repairs
 * (`Partial<PairRecords>`, or `as PairRecords` on the spread) both type-check
 * unconditionally, so adding a required field to `PairRecords` would break
 * `bundle-driver.ts` and leave the test that MIRRORS it compiling clean.
 */
function originalRecords(
  bytes: ArrayBuffer,
): Pick<PairRecords, 'originalParas' | 'quarantinedOriginal'> {
  const { paras, quarantinedCount } = preprocessBundleDocument(bytes, CONFIG, 'acceptAll');
  return { originalParas: paras, quarantinedOriginal: quarantinedCount };
}

function revisedRecords(
  bytes: ArrayBuffer,
): Pick<PairRecords, 'revisedParas' | 'quarantinedRevised'> {
  const { paras, quarantinedCount } = preprocessBundleDocument(bytes, CONFIG, 'acceptAll');
  return { revisedParas: paras, quarantinedRevised: quarantinedCount };
}

/** Turn a confirmed proposal into runnable pairs, the way the driver does. */
function confirmedPairs(
  sideA: readonly PairingSketch[],
  sideB: readonly PairingSketch[],
  bytesFor: Map<string, ArrayBuffer>,
): ConfirmedPairInput[] {
  const proposal = proposePairing({ sideA, sideB, generatedAt: DATE });
  const byId = new Map([...sideA, ...sideB].map(s => [s.docId, s]));

  return proposal.rows
    .filter(r => r.a !== null && r.b !== null)
    .map((r): ConfirmedPairInput => {
      const a = byId.get(r.a as string) as PairingSketch;
      const b = byId.get(r.b as string) as PairingSketch;
      return {
        id: r.id,
        label: r.label,
        artefactStem: a.fileName.replace(/\.docx$/i, ''),
        // Built per pair, when the lane asks — the way the driver does it, and
        // for the reason it does: a sketch carries hashes, never records.
        records: (): Promise<PairRecords> => Promise.resolve({
          ...originalRecords(bytesFor.get(a.docId) as ArrayBuffer),
          ...revisedRecords(bytesFor.get(b.docId) as ArrayBuffer),
        }),
        originalBytes: () => Promise.resolve(bytesFor.get(a.docId) as ArrayBuffer),
        originalByteLength: a.byteLength,
        revisedByteLength: b.byteLength,
        ...(r.runnable ? {} : { unrunnableReason: r.reason as string }),
      };
    });
}

function pool(size = 2): BatchPool {
  return createBatchPool(Array.from({ length: size }, () => createInProcessWorker()));
}

function run(pairs: readonly ConfirmedPairInput[], extra = {}) {
  return runManyToManyBatch({
    pairs,
    config: CONFIG,
    author: 'Mutatis Compare',
    date: DATE,
    flattenA: 'acceptAll',
    flattenB: 'acceptAll',
    pool: pool(),
    ...extra,
  });
}

/** A realistic two-bundle scenario, with the bytes kept for artefact assembly. */
function bundles() {
  const bytesFor = new Map<string, ArrayBuffer>();
  const make = (id: string, name: string, text: string, side: BundleSide): PairingSketch => {
    bytesFor.set(id, bytesOf(text));
    return sketchOf(id, name, text, side);
  };

  const sideA = [
    make('a-spa', 'Project Falcon SPA v6.docx', SPA, 'a'),
    make('a-deed', 'Project Falcon Tax Deed v2.docx', DEED, 'a'),
  ];
  const sideB = [
    make('b-spa', 'Project Falcon SPA v7.docx',
      SPA.replace('paid in full on the Completion Date', 'paid in cleared funds on Completion'), 'b'),
    make('b-deed', 'Project Falcon Tax Deed (execution).docx',
      DEED.replace('within seven years', 'within five years'), 'b'),
  ];
  return { sideA, sideB, bytesFor };
}

describe('a confirmed pairing produces one redline per pair', () => {
  it('runs every confirmed pair and reports a row for each', async () => {
    const { sideA, sideB, bytesFor } = bundles();
    const pairs = confirmedPairs(sideA, sideB, bytesFor);
    expect(pairs).toHaveLength(2);

    const index = await run(pairs).index;

    expect(index.mode).toBe('many-to-many');
    expect(index.rows).toHaveLength(2);
    expect(index.rows.every(r => r.status === 'done')).toBe(true);
    expect(index.totals.done).toBe(2);
    // Each pair found its own counterpart's single edit.
    for (const row of index.rows) expect(row.changeCount).toBeGreaterThan(0);
  });

  it('labels each row "A ▸ B" and names the artefact after the side-A document', async () => {
    const { sideA, sideB, bytesFor } = bundles();
    const index = await run(confirmedPairs(sideA, sideB, bytesFor)).index;

    for (const row of index.rows) {
      expect(row.label).toMatch(/ ▸ /);
      expect(row.artefacts[0]?.kind).toBe('redlineDocx');
      expect(row.artefacts[0]?.fileName).toMatch(/ \(redline\)\.docx$/);
    }
    const spaRow = index.rows.find(r => r.label.startsWith('Project Falcon SPA v6'));
    expect(spaRow?.artefacts[0]?.fileName).toBe('Project Falcon SPA v6 (redline).docx');
  });

  it('produces a real .docx package for each pair', async () => {
    const { sideA, sideB, bytesFor } = bundles();
    const index = await run(confirmedPairs(sideA, sideB, bytesFor)).index;

    for (const row of index.rows) {
      const bytes = new Uint8Array(await (row.artefacts[0]?.blob as Blob).arrayBuffer());
      // ZIP local header — a real package, not an empty Blob.
      expect(Array.from(bytes.slice(0, 4))).toEqual([0x50, 0x4b, 0x03, 0x04]);
      const bundle = new PizZipBundle(bytes);
      expect(bundle.part('word/document.xml')).toContain('w:ins');
    }
  });

  it('states the flatten choice for both sides on the index', async () => {
    const { sideA, sideB, bytesFor } = bundles();
    const index = await runManyToManyBatch({
      pairs: confirmedPairs(sideA, sideB, bytesFor),
      config: CONFIG,
      author: 'Mutatis Compare',
      date: DATE,
      flattenA: 'acceptAll',
      flattenB: 'rejectAll',
      pool: pool(),
    }).index;

    expect(index.flatten).toEqual({ template: 'acceptAll', variant: 'rejectAll' });
    // No template label: there is no template in a many-to-many batch.
    expect('templateLabel' in index).toBe(false);
  });
});

describe('the cap a bundle run is refused by is the bundle one', () => {
  it('names documents, not variants, when there are too many pairs', () => {
    // A pairing gives each side-A document at most one counterpart, so the pair
    // count is bounded by one side's document count — `maxVariants` (100) can
    // never fire here, and telling someone who handed over two folders that
    // they have too many "variants" names a concept their run does not contain.
    const { sideA, sideB, bytesFor } = bundles();
    const one = confirmedPairs(sideA, sideB, bytesFor)[0] as ConfirmedPairInput;
    const many = Array.from({ length: MAX_BUNDLE_DOCUMENTS + 1 }, (_, i) => ({
      ...one,
      id: `pair-${i}`,
      label: `Pair ${i}`,
    }));

    expect(() => run(many)).toThrow(/document/i);
  });
});

describe('a pair\'s records live only as long as its lane', () => {
  /**
   * Track how many pairs' records are alive at once.
   *
   * "Alive" starts when `records()` is called and ends when that pair's row
   * reaches a terminal state — which is when the lane lets go of them. An
   * earlier version of this decremented one microtask after `records()`
   * resolved, which meant a sequential eager loop never overlapped and the
   * whole point went unpinned: the eager shape §4.3 forbids passed it.
   */
  function withRecordAccounting(pairs: readonly ConfirmedPairInput[]) {
    const alive = new Set<string>();
    let peak = 0;
    const instrumented = pairs.map(p => ({
      ...p,
      records: async () => {
        alive.add(p.id);
        peak = Math.max(peak, alive.size);
        return p.records();
      },
    }));
    return {
      pairs: instrumented,
      peak: () => peak,
      release: (row: BatchRow) => { alive.delete(row.id); },
    };
  }

  it('never holds more pairs\' records at once than the pool has lanes', async () => {
    // Spec §4.3 step 4. Passing the arrays in meant every confirmed pair's
    // records were live from the moment the batch was assembled — at the §4.4
    // cap, a hundred document-sized `ParaRecord[]` at once, which §4.3 opens by
    // naming as an out-of-memory crash rather than a slowdown.
    const { sideA, sideB, bytesFor } = bundles();
    const tracked = withRecordAccounting(confirmedPairs(sideA, sideB, bytesFor));

    await run(tracked.pairs, { pool: pool(1), onRowChanged: tracked.release }).index;

    expect(tracked.peak()).toBe(1);
  });

  it('scales that ceiling with the pool, and no further', async () => {
    // The control: a peak of one under a one-lane pool could equally mean the
    // records are never built at all. Two lanes must reach two, and stop.
    const { sideA, sideB, bytesFor } = bundles();
    const tracked = withRecordAccounting(confirmedPairs(sideA, sideB, bytesFor));

    await run(tracked.pairs, { pool: pool(2), onRowChanged: tracked.release }).index;

    expect(tracked.peak()).toBeGreaterThan(0);
    expect(tracked.peak()).toBeLessThanOrEqual(2);
  });

  it('a sketch carries hashes, never records', async () => {
    // The property the two tests above depend on. If a sketch started holding
    // `ParaRecord[]` again the peak would go back to the whole bundle's, and
    // neither of them would notice.
    const { sideA } = bundles();
    const spa = sideA.find(s => s.docId === 'a-spa') as PairingSketch;

    expect(Object.keys(spa)).not.toContain('paras');
    expect(spa.paraHashes.size).toBeGreaterThan(0);
  });

  it('reads a side-A package only when its artefact is assembled', async () => {
    // Holding M open packages across a batch is the memory failure §4.3 exists
    // to prevent, so the bytes are a lazy callback.
    const { sideA, sideB, bytesFor } = bundles();
    const reads: string[] = [];
    const pairs = confirmedPairs(sideA, sideB, bytesFor).map(p => ({
      ...p,
      originalBytes: () => { reads.push(p.id); return p.originalBytes(); },
    }));

    expect(reads).toHaveLength(0);
    await run(pairs, { pool: pool(1) }).index;
    expect(reads).toHaveLength(2);
    expect(new Set(reads).size).toBe(2);
  });

  it('does not read the package at all when the artefact is not retained', async () => {
    const { sideA, sideB, bytesFor } = bundles();
    let reads = 0;
    const pairs = confirmedPairs(sideA, sideB, bytesFor).map(p => ({
      ...p,
      originalBytes: () => { reads += 1; return p.originalBytes(); },
    }));

    const index = await run(pairs, { pool: pool(1), artefactBudgetBytes: 1 }).index;

    expect(index.rows.every(r => r.artefactOnDemand)).toBe(true);
    // Still compared, still counted — the package simply was never needed.
    expect(index.rows.every(r => r.changeCount > 0)).toBe(true);
    expect(reads).toBe(0);
  });
});

describe('an unrunnable pair is a row, not an absence', () => {
  it('keeps a cross-format pair as a row with its reason', async () => {
    const { sideA, sideB, bytesFor } = bundles();
    const pairs = confirmedPairs(sideA, sideB, bytesFor);
    const blocked: ConfirmedPairInput[] = [
      { ...(pairs[0] as ConfirmedPairInput), unrunnableReason: CROSS_FORMAT_REASON },
      pairs[1] as ConfirmedPairInput,
    ];

    const index = await run(blocked).index;

    expect(index.rows).toHaveLength(2);
    expect(index.totals.unrunnable).toBe(1);
    const row = index.rows.find(r => r.status === 'unrunnable');
    expect(row?.reason).toBe(CROSS_FORMAT_REASON);
    expect(row?.label).toMatch(/ ▸ /);
  });

  it('refuses duplicate pair ids rather than collapsing two rows into one', () => {
    const { sideA, sideB, bytesFor } = bundles();
    const pairs = confirmedPairs(sideA, sideB, bytesFor);
    const clashing = [pairs[0] as ConfirmedPairInput, { ...(pairs[1] as ConfirmedPairInput), id: (pairs[0] as ConfirmedPairInput).id }];

    expect(() => run(clashing)).toThrow(/duplicate document id/i);
  });

  it('reports a failure with its reason and keeps going', async () => {
    const { sideA, sideB, bytesFor } = bundles();
    const pairs = confirmedPairs(sideA, sideB, bytesFor);
    const broken = [
      { ...(pairs[0] as ConfirmedPairInput), originalBytes: () => Promise.resolve(new Uint8Array([1, 2]).buffer as ArrayBuffer) },
      pairs[1] as ConfirmedPairInput,
    ];

    const index = await run(broken, { pool: pool(1) }).index;

    expect(index.rows).toHaveLength(2);
    expect(index.totals.failed).toBe(1);
    expect(index.rows.find(r => r.status === 'failed')?.reason).toBeTruthy();
    expect(index.totals.done).toBe(1);
  });

  it('handles an empty confirmed set', async () => {
    const index = await run([]).index;
    expect(index.rows).toEqual([]);
    expect(index.totals.pairs).toBe(0);
    expect(index.mode).toBe('many-to-many');
  });
});

describe('an oversize document costs its own pair and nothing else', () => {
  // The runner used to call `checkDocxByteCap`, which THROWS, over every
  // confirmed pair. One document over the cap therefore ended the whole run
  // before a single pair had been dispatched — and the grid had already shown
  // that pair with the cap named, so the user had confirmed a run knowing it
  // would not be compared. Turning an accepted, disclosed limitation into a
  // batch-wide failure is the many-to-many form of losing ninety-nine good
  // redlines to one bad file.
  // A third deal, so the extra pair cannot be confused with the SPA or the
  // deed. Two documents with the same body would let the assignment cross them
  // over and the test would be measuring the wrong pair.
  const LEASE = [
    'This Lease is granted by the Landlord to the Tenant.',
    'The Term commences on the Rent Commencement Date and runs for ten years.',
    'The Tenant shall keep the Premises in good and substantial repair.',
    'The Rent is reviewed on each fifth anniversary of the Term.',
    'This Lease is governed by the laws of England and Wales.',
  ].join('\n');

  /**
   * The pairs the driver would build, with one document's byte length pushed
   * over the cap AFTER the proposal — so the runner meets a pair the grid did
   * NOT block, which is the case that used to throw.
   */
  function withOversize(oversizeId: string) {
    const { sideA, sideB, bytesFor } = bundles();
    const a = [...sideA, sketchOf('a-big', 'Big Lease v1.docx', LEASE, 'a')];
    const b = [...sideB, sketchOf('b-big', 'Big Lease v2.docx',
      LEASE.replace('ten years', 'fifteen years'), 'b')];
    bytesFor.set('a-big', bytesOf(LEASE));

    return confirmedPairs(a, b, bytesFor).map(p => p.id.includes('big')
      ? {
        ...p,
        ...(oversizeId === 'a'
          ? { originalByteLength: MAX_DOCX_BYTES + 1 }
          : { revisedByteLength: MAX_DOCX_BYTES + 1 }),
      }
      : p);
  }

  it('keeps every other pair when one side-A document is over the cap', async () => {
    const index = await run(withOversize('a')).index;

    expect(index.totals.pairs).toBe(3);
    expect(index.totals.done).toBe(2);
    expect(index.totals.unrunnable).toBe(1);
    const blocked = index.rows.find(r => r.status === 'unrunnable');
    expect(blocked?.label).toMatch(/Big Lease/);
    expect(blocked?.reason).toMatch(/per-document cap/);
  });

  it('checks side B too — a cap enforced on one side of a pair is not a cap', async () => {
    // Side B's bytes are never read (its records came from the sketch), but its
    // SIZE is still a limit the pair has to satisfy. This was unchecked
    // entirely: the runner looked only at side A.
    const index = await run(withOversize('b')).index;

    expect(index.totals.done).toBe(2);
    expect(index.totals.unrunnable).toBe(1);
    expect(index.rows.find(r => r.status === 'unrunnable')?.reason).toMatch(/per-document cap/);
  });

  it('leaves the grid\'s own reason in place rather than restating it', async () => {
    // A pair the proposal already blocked keeps the sentence the user read and
    // accepted. Two different sentences for one refusal would read as two
    // failures.
    const { sideA, sideB, bytesFor } = bundles();
    const pairs = confirmedPairs(sideA, sideB, bytesFor).map((p, i) =>
      i === 0 ? { ...p, unrunnableReason: 'A reason the grid already gave.' } : p);

    const index = await run(pairs).index;

    expect(index.rows.find(r => r.status === 'unrunnable')?.reason)
      .toBe('A reason the grid already gave.');
  });

  it('does not count the bytes of a pair it has just decided not to compare', async () => {
    // The aggregate cap is justified by what the run HOLDS AT ONCE (spec §6.6).
    // Summing before the demotions counted documents the very next statement
    // had decided not to read, so one over-cap document could refuse a batch
    // that fitted comfortably.
    const { sideA, sideB, bytesFor } = bundles();
    const one = confirmedPairs(sideA, sideB, bytesFor)[0] as ConfirmedPairInput;
    const pairs = [
      // Alone, far over the aggregate cap — but demoted, so never read.
      { ...one, id: 'huge', originalByteLength: MAX_BATCH_TOTAL_BYTES * 2 },
      { ...one, id: 'small' },
    ];

    const index = await run(pairs).index;

    expect(index.totals.unrunnable).toBe(1);
    expect(index.totals.done).toBe(1);
  });

  it('still refuses a batch whose total input size is over the aggregate cap', async () => {
    // The per-document demotion above must not have removed the whole-batch
    // wall. This one IS a refusal: it is about how much the run would hold at
    // once, not about one unusable document.
    // Every pair is UNDER the per-document cap, so none is demoted — the
    // aggregate cap has to be what refuses this, not the per-document one. At
    // 50 MiB a side it takes eleven pairs to clear 1 GiB.
    const { sideA, sideB, bytesFor } = bundles();
    const one = confirmedPairs(sideA, sideB, bytesFor)[0] as ConfirmedPairInput;
    const pairs = Array.from({ length: 11 }, (_, i) => ({
      ...one,
      id: `pair-${i}`,
      originalByteLength: MAX_DOCX_BYTES,
      revisedByteLength: MAX_DOCX_BYTES,
    }));
    expect(11 * 2 * MAX_DOCX_BYTES).toBeGreaterThan(MAX_BATCH_TOTAL_BYTES);

    // Thrown from the call itself, not from the returned promise: a cap is a
    // pre-flight refusal, so nothing has been started to fail.
    expect(() => run(pairs)).toThrow(/totalBytes/i);
  });
});

describe('cancellation keeps what was already produced', () => {
  it('keeps completed rows and marks the rest cancelled', async () => {
    const { sideA, sideB, bytesFor } = bundles();
    const pairs = confirmedPairs(sideA, sideB, bytesFor);

    let running: ReturnType<typeof run> | null = null;
    running = runManyToManyBatch({
      pairs,
      config: CONFIG,
      author: 'Mutatis Compare',
      date: DATE,
      flattenA: 'acceptAll',
      flattenB: 'acceptAll',
      pool: pool(1),
      onRowChanged: row => { if (row.status === 'done') running?.cancel(); },
    });

    const index = await running.index;

    expect(index.rows).toHaveLength(2);
    expect(index.totals.done).toBeGreaterThanOrEqual(1);
    expect(index.totals.done + index.totals.cancelled).toBe(2);
    // A completed pair keeps its artefact — cancelling stops the batch, it does
    // not discard what the batch already produced.
    for (const row of index.rows.filter(r => r.status === 'done')) {
      expect(row.artefacts[0]?.byteLength).toBeGreaterThan(0);
    }
  });
});
