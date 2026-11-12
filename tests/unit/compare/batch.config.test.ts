/**
 * `compare.batch.*` resolution (ADR-0037; spec §7).
 *
 * The keys travel on `CompareConfig` so a batch setting is reached the way
 * every other Compare setting is; the defaults stay in the feature's own
 * `constants.ts` because `shared/` must not depend on a feature. This file
 * tests the seam between the two, and one property in particular: an unset key
 * means "use the constant", never "use zero".
 */
import { describe, it, expect } from 'vitest';

import { getCompareConfig, UK_PROFILE } from '@shared/config/index.js';
import type { CompareConfig } from '@shared/config/index.js';
import { resolveBatchConfig } from '@features/compare/batch/config.js';
import { MAX_BATCH_VARIANTS, MAX_BUNDLE_DOCUMENTS } from '@features/compare/batch/limits.js';
import {
  PAIRING_AUTO_PAIR_MARGIN,
  PAIRING_AUTO_PAIR_SCORE,
  PAIRING_MIN_SKETCH_TOKENS,
  PAIRING_NO_CONTENT_MATCH_SCORE,
  PAIRING_NO_MATCH_SCORE,
  PAIRING_TIE_EPSILON,
  PAIRING_WEIGHT_CONTENT,
  PAIRING_WEIGHT_FILENAME,
} from '@features/compare/batch/constants.js';
import { runOneToManyBatch, type BatchDocumentInput } from '@features/compare/batch/runner.js';
import { BatchInputTooLargeError } from '@features/compare/batch/limits.js';
import { synthesiseDocx } from '@features/compare/io/text/synthesise-docx.js';
import {
  createBatchPool,
  createInProcessWorker,
} from '@features/compare/batch/worker/index.js';
import { resolvePairingSettings } from '@features/compare/batch/config.js';
import { proposePairing } from '@features/compare/batch/pairing/propose.js';
import { rankCandidates, type PairScore } from '@features/compare/batch/pairing/score.js';
import { isSketchWorthy, type PairingSketch } from '@features/compare/batch/pairing/sketch.js';
import type { ParaRecord } from '@features/compare/model.js';

const BASE = getCompareConfig(UK_PROFILE);

describe('resolveBatchConfig', () => {
  it('falls back to the feature constants when nothing is set', () => {
    const resolved = resolveBatchConfig(BASE);

    expect(resolved.maxVariants).toBe(MAX_BATCH_VARIANTS);
    expect(resolved.maxBundleDocuments).toBe(MAX_BUNDLE_DOCUMENTS);
    expect(resolved.poolSize).toBe(0);
    expect(resolved.pairing).toEqual({
      weightContent: PAIRING_WEIGHT_CONTENT,
      weightFilename: PAIRING_WEIGHT_FILENAME,
      tieEpsilon: PAIRING_TIE_EPSILON,
      autoPairScore: PAIRING_AUTO_PAIR_SCORE,
      autoPairMargin: PAIRING_AUTO_PAIR_MARGIN,
      noMatchScore: PAIRING_NO_MATCH_SCORE,
      noContentMatchScore: PAIRING_NO_CONTENT_MATCH_SCORE,
      minSketchTokens: PAIRING_MIN_SKETCH_TOKENS,
    });
  });

  it('takes every key an override supplies', () => {
    const config: CompareConfig = {
      ...BASE,
      batch: {
        maxVariants: 12,
        maxBundleDocuments: 7,
        poolSize: 3,
        pairing: { weightContent: 0.7, weightFilename: 0.3, minSketchTokens: 9 },
      },
    };
    const resolved = resolveBatchConfig(config);

    expect(resolved.maxVariants).toBe(12);
    expect(resolved.maxBundleDocuments).toBe(7);
    expect(resolved.poolSize).toBe(3);
    expect(resolved.pairing.weightContent).toBe(0.7);
    expect(resolved.pairing.minSketchTokens).toBe(9);
    // Untouched keys still come from the constants, not from zero.
    expect(resolved.pairing.tieEpsilon).toBe(PAIRING_TIE_EPSILON);
    expect(resolved.pairing.autoPairScore).toBe(PAIRING_AUTO_PAIR_SCORE);
  });

  it('keeps an explicit 0 distinct from an absent key', () => {
    // `poolSize: 0` means auto-size — a real value, not "unset". Resolving with
    // a spread instead of `??` would make the two indistinguishable.
    expect(resolveBatchConfig({ ...BASE, batch: { poolSize: 0 } }).poolSize).toBe(0);
    expect(resolveBatchConfig({ ...BASE, batch: {} }).poolSize).toBe(0);
    expect(resolveBatchConfig({ ...BASE, batch: { maxVariants: 0 } }).maxVariants).toBe(0);
    expect(resolveBatchConfig({ ...BASE, batch: {} }).maxVariants).toBe(MAX_BATCH_VARIANTS);
  });

  it('rejects a nonsensical value rather than passing it through', () => {
    // Config comes from a profile a human edits. A NaN or a negative reaching a
    // cap does not throw — it silently refuses every batch or accepts every
    // batch, depending on which comparison it lands in. Validated at the seam.
    const bad = resolveBatchConfig({
      ...BASE,
      batch: {
        maxVariants: Number.NaN,
        maxBundleDocuments: -5,
        pairing: { weightContent: 1.4, tieEpsilon: -0.1, minSketchTokens: Number.NaN },
      },
    });

    expect(bad.maxVariants).toBe(MAX_BATCH_VARIANTS);
    expect(bad.maxBundleDocuments).toBe(MAX_BUNDLE_DOCUMENTS);
    expect(bad.pairing.weightContent).toBe(PAIRING_WEIGHT_CONTENT);
    expect(bad.pairing.tieEpsilon).toBe(PAIRING_TIE_EPSILON);
    expect(bad.pairing.minSketchTokens).toBe(PAIRING_MIN_SKETCH_TOKENS);
  });

  it('floors a fractional count rather than half-comparing a document', () => {
    expect(resolveBatchConfig({ ...BASE, batch: { maxVariants: 7.9 } }).maxVariants).toBe(7);
  });

  it('accepts the boundary values, which are legitimate settings', () => {
    const edge = resolveBatchConfig({
      ...BASE,
      batch: { pairing: { weightContent: 0, weightFilename: 1, autoPairScore: 1 } },
    });
    expect(edge.pairing.weightContent).toBe(0);
    expect(edge.pairing.weightFilename).toBe(1);
    expect(edge.pairing.autoPairScore).toBe(1);
  });

  it('resolves the same values from a bare profile as from an empty override', () => {
    expect(resolveBatchConfig({ ...BASE, batch: {} })).toEqual(resolveBatchConfig(BASE));
  });
});

describe('the runner reads the resolved caps, not the constants directly', () => {
  function doc(id: string): BatchDocumentInput {
    const raw = new Uint8Array(synthesiseDocx('The Seller shall pay on Completion.').toUint8Array());
    const buffer = raw.buffer.slice(
      raw.byteOffset,
      raw.byteOffset + raw.byteLength,
    ) as ArrayBuffer;
    return {
      id,
      label: `${id}.docx`,
      byteLength: buffer.byteLength,
      bytes: () => Promise.resolve(buffer),
    };
  }

  it('refuses a batch over a profile-tightened variant cap', () => {
    const config: CompareConfig = { ...BASE, batch: { maxVariants: 2 } };

    expect(() => runOneToManyBatch({
      template: doc('template'),
      variants: [doc('a'), doc('b'), doc('c')],
      config,
      author: 'Mutatis Compare',
      date: '2026-09-10T09:00:00.000Z',
      flattenTemplate: 'acceptAll',
      flattenVariant: 'acceptAll',
      pool: createBatchPool([createInProcessWorker()]),
    })).toThrow(BatchInputTooLargeError);
  });

  it('accepts the same batch under the default cap', async () => {
    const index = await runOneToManyBatch({
      template: doc('template'),
      variants: [doc('a'), doc('b'), doc('c')],
      config: BASE,
      author: 'Mutatis Compare',
      date: '2026-09-10T09:00:00.000Z',
      flattenTemplate: 'acceptAll',
      flattenVariant: 'acceptAll',
      pool: createBatchPool([createInProcessWorker()]),
    }).index;

    expect(index.rows).toHaveLength(3);
  });
});

describe('the pairing keys reach the pairing maths (spec §7)', () => {
  // A review replaced four of these values with zero and ran the whole suite
  // excluding this file: 195 files, 2120 tests, all green. The resolver
  // resolved them and nothing consumed the result — every pairing module
  // imported the raw constants, so `compare.batch.pairing.autoPairScore` in a
  // profile did nothing whatsoever. These tests assert the values ARRIVE, at
  // the outcome rather than at the seam.
  const hashes = (n: number, prefix = 'p'): string[] =>
    Array.from({ length: n }, (_, i) => `${prefix}-${i}`);

  function sketch(
    docId: string,
    fileName: string,
    paraHashes: readonly string[],
  ): PairingSketch {
    return {
      docId,
      side: docId.startsWith('a') ? 'a' : 'b',
      sourceKind: 'docx',
      fileName,
      byteLength: 2048,
      paraHashes: new Set(paraHashes),
      definedTerms: new Set(['seller']),
      partyNames: new Set(['seller']),
      quarantinedCount: 0,
    };
  }

  const AT = '2026-09-10T09:00:00.000Z';

  /** One pair that lands in `propose` under the defaults: good, but contested. */
  function contested(config?: CompareConfig) {
    return proposePairing({
      sideA: [sketch('a1', 'SPA v6.docx', hashes(40))],
      sideB: [
        sketch('b1', 'SPA v7.docx', [...hashes(34), ...hashes(6, 'x')]),
        sketch('b2', 'SPA draft.docx', [...hashes(30), ...hashes(10, 'y')]),
      ],
      generatedAt: AT,
      ...(config === undefined ? {} : { config }),
    });
  }

  it('a profile that lowers autoPairScore and autoPairMargin pre-ticks a pair the defaults do not', () => {
    expect(contested().rows.find(r => r.a === 'a1')?.band).toBe('propose-with-rivals');

    const relaxed: CompareConfig = {
      ...BASE,
      batch: { pairing: { autoPairScore: 0.5, autoPairMargin: 0.01 } },
    };
    expect(contested(relaxed).rows.find(r => r.a === 'a1')?.band).toBe('auto-pair');
  });

  it('a profile that raises noMatchScore sends a real pair to the dummies instead', () => {
    // The threshold below which a document would rather go unmatched. Raised
    // above the pair's score, the assignment must stop proposing it — which
    // also proves `noMatchCost` is derived from the CONFIGURED value, since the
    // dummy price is what the solver actually compares against.
    expect(contested().rows.find(r => r.a === 'a1')?.b).toBe('b1');

    const strict: CompareConfig = { ...BASE, batch: { pairing: { noMatchScore: 0.99 } } };
    const row = contested(strict).rows.find(r => r.a === 'a1');
    expect(row?.b).toBeNull();
    expect(row?.band).toBe('unmatched');
  });

  it('never reports a score above 100%, however far a profile pushes noMatchScore', () => {
    // `noContentCeiling` is `noMatchScore + headroom`, and `noMatchScore` is a
    // unit value — so the top of its own range puts the ceiling at 1.05 and the
    // grid prints `105%`. A score is a proportion; the derivation is clamped.
    const edge: CompareConfig = { ...BASE, batch: { pairing: { noMatchScore: 1 } } };
    expect(resolvePairingSettings(edge).noContentCeiling).toBe(1);
  });

  it('a profile that raises noContentMatchScore stops a filename-only pair forming', () => {
    // The second assignment phase's bar (spec §6.4). It is reached only by a
    // pair with NO content evidence, so it needs its own fixture: a side-B
    // document nothing could be sketched from, whose name is close enough to
    // pair on and nothing else. Under the defaults phase 2 places it; raised
    // above the filename similarity, phase 2 must leave it unmatched.
    const nameOnly = (config?: CompareConfig) =>
      proposePairing({
        sideA: [sketch('a1', 'Tax Deed.docx', hashes(40))],
        sideB: [sketch('b1', 'Tax Deed of Covenant.docx', [])],
        generatedAt: AT,
        ...(config === undefined ? {} : { config }),
      }).rows.find(r => r.a === 'a1');

    expect(nameOnly()?.b).toBe('b1');

    const strict: CompareConfig = { ...BASE, batch: { pairing: { noContentMatchScore: 0.8 } } };
    expect(nameOnly(strict)?.b).toBeNull();
    expect(nameOnly(strict)?.band).toBe('unmatched');
  });

  it('a profile that reweights the signals changes the score itself', () => {
    const byDefault = contested().rows.find(r => r.a === 'a1')?.score?.score as number;

    const nameLed: CompareConfig = {
      ...BASE,
      batch: { pairing: { weightContent: 0.2, weightFilename: 0.8 } },
    };
    const reweighted = contested(nameLed).rows.find(r => r.a === 'a1')?.score?.score as number;

    expect(reweighted).not.toBeCloseTo(byDefault, 6);
  });

  it('a profile that widens tieEpsilon lets the cast reorder rivals it otherwise could not', () => {
    // The tie band is the only place the cast is consulted at all, so widening
    // it is the only way a profile can change what the cast does.
    const cast = (b: string, score: number, castValue: number): PairScore =>
      ({ a: 'a1', b, content: score, filename: score, cast: castValue, contentAvailable: true, score });

    const candidates = [cast('b1', 0.80, 0.1), cast('b2', 0.70, 0.9)];

    expect(rankCandidates(candidates)[0]?.b).toBe('b1');
    expect(
      rankCandidates(candidates, resolvePairingSettings({
        ...BASE,
        batch: { pairing: { tieEpsilon: 0.5 } },
      }))[0]?.b,
    ).toBe('b2');
  });

  it('a profile that raises minSketchTokens excludes paragraphs a sketch would otherwise keep', () => {
    const record = { normText: 'The Seller shall pay the Purchase Price on Completion.' } as ParaRecord;

    expect(isSketchWorthy(record, BASE)).toBe(true);
    expect(isSketchWorthy(record, { ...BASE, batch: { pairing: { minSketchTokens: 99 } } }))
      .toBe(false);
  });
});
