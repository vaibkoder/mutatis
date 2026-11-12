/**
 * Compare at scale — the acceptance sweep, both modes (ADR-0037; spec §9).
 *
 * B1–B13 (one-to-many) and M1–M8 (many-to-many) are each proven in the file
 * that owns the behaviour; this file is the index of where, plus the invariant
 * checks (I1, I3, and the two structural ones below) that have no natural home
 * elsewhere. Keeping the map in code rather than in a comment means a renamed
 * or deleted acceptance file fails a test instead of quietly leaving a spec
 * claim unbacked.
 */
import { describe, it, expect } from 'vitest';
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { getCompareConfig, UK_PROFILE } from '@shared/config/index.js';
import { synthesiseDocx } from '@features/compare/io/text/synthesise-docx.js';
import { runOneToManyBatch, type BatchDocumentInput } from '@features/compare/batch/runner.js';
import {
  createBatchPool,
  createInProcessWorker,
} from '@features/compare/batch/worker/index.js';
import { buildBatchIndex } from '@features/compare/batch/index-model.js';
import { isDownloadable } from '@features/compare/batch/artefacts.js';
import type { BatchRow } from '@features/compare/batch/model.js';

const ROOT = resolve(__dirname, '../../..');
const BATCH_SRC = resolve(ROOT, 'src/features/compare/batch');

/** Where each acceptance criterion is proven. */
const COVERAGE: Record<string, { claim: string; provenIn: string }> = {
  B1: {
    claim: 'two driver instances hold independent state',
    provenIn: 'tests/unit/compare.driver.instances.test.ts',
  },
  B2: {
    claim: 'the template is preprocessed exactly once per batch',
    provenIn: 'tests/unit/compare/batch.runner.test.ts',
  },
  B3: {
    claim: 'the shared template records are not mutated',
    provenIn: 'tests/unit/compare/batch.runner.test.ts',
  },
  B4: {
    claim: 'no PipelineResult outlives its own pair',
    provenIn: 'tests/unit/compare/batch.runner.test.ts',
  },
  B5: {
    claim: 'the index is filter-independent and loses no pair',
    provenIn: 'tests/unit/compare/batch.index-model.test.ts',
  },
  B6: {
    claim: 'cancelling keeps completed rows and their artefacts',
    provenIn: 'tests/unit/compare/batch.runner.test.ts',
  },
  B7: {
    claim: 'a corrupt file is one failed row and the batch continues',
    provenIn: 'tests/unit/compare/batch.runner.test.ts',
  },
  B8: {
    claim: 'each cap throws before any pair runs',
    provenIn: 'tests/unit/compare/batch.limits.test.ts',
  },
  B9: {
    claim: 'pooled and in-process runs produce identical output, job and row alike',
    provenIn: 'tests/unit/compare/batch.runner.test.ts',
  },
  B10: {
    claim: 'an unchanged variant is a done row with zero changes',
    provenIn: 'tests/unit/compare/batch.runner.test.ts',
  },
  B11: {
    claim: 'a row past the artefact ceiling is still compared, counted and downloadable',
    provenIn: 'tests/unit/compare/batch.runner.test.ts',
  },
  B12: {
    claim: 'download-all re-runs on-demand rows lazily, one at a time',
    provenIn: 'tests/unit/compare/batch.driver.test.ts',
  },
  B13: {
    claim: 'the flatten choice is stated on the index and duplicate ids are refused',
    provenIn: 'tests/unit/compare/batch.index-model.test.ts',
  },
  M1: {
    claim: 'the assignment recovers a known pairing exactly',
    provenIn: 'tests/unit/compare/batch.pairing.assign.test.ts',
  },
  M2: {
    claim: 'the near-identical v6/v7 pair lands in propose-with-rivals, with rival scores',
    provenIn: 'tests/unit/compare/batch.pairing.propose.test.ts',
  },
  M3: {
    claim: 'surplus documents come out unmatched rather than force-matched',
    provenIn: 'tests/unit/compare/batch.pairing.assign.test.ts',
  },
  M4: {
    claim: 'shuffling both bundles gives an identical assignment',
    provenIn: 'tests/unit/compare/batch.pairing.assign.test.ts',
  },
  M5: {
    claim: 'one shared cast does not decide a pairing — it applies inside the tie band only',
    provenIn: 'tests/unit/compare/batch.pairing.score.test.ts',
  },
  M6: {
    claim: 'filename normalisation strips every version token and day-first date form',
    provenIn: 'tests/unit/compare/batch.pairing.signals.test.ts',
  },
  M7: {
    claim: 'no comparison runs before the grid is confirmed',
    provenIn: 'tests/unit/web/pairing-grid.test.tsx',
  },
  M8: {
    claim: 'a cross-format pair is surfaced with its pairing, never dropped',
    provenIn: 'tests/unit/compare/batch.pairing.propose.test.ts',
  },
  I2: {
    claim: 'the Word add-in is untouched and cannot reach the batch surface',
    provenIn: 'tests/unit/web/web-batch-compare.test.tsx',
  },
};

/**
 * Behaviours that are not numbered acceptance criteria but are load-bearing
 * enough that losing their test should fail this file too.
 */
const ALSO_PROVEN: Record<string, string> = {
  'the worker fallback ladder ends in a working in-process pool':
    'tests/unit/compare/batch.worker.test.ts',
  'a confirmed pairing produces one redline per pair':
    'tests/unit/compare/batch.many.test.ts',
  "a pair's records live only as long as the lane running it":
    'tests/unit/compare/batch.many.test.ts',
  'the M7 gate is enforced in the driver, not only in the component':
    'tests/unit/compare/batch.bundle-driver.test.ts',
  'the bundle surface injects a live driver rather than rendering against the stub':
    'tests/unit/web/web-bundle-compare.test.tsx',
};

describe('acceptance coverage map', () => {
  for (const [id, { claim, provenIn }] of Object.entries(COVERAGE)) {
    it(`${id} — ${claim} — is proven in ${provenIn}`, () => {
      expect(existsSync(resolve(ROOT, provenIn))).toBe(true);
    });
  }

  for (const [claim, provenIn] of Object.entries(ALSO_PROVEN)) {
    it(`${claim} — is proven in ${provenIn}`, () => {
      expect(existsSync(resolve(ROOT, provenIn))).toBe(true);
    });
  }
});

describe('I1 — every batch source is host-free and network-free', () => {
  /** Every `.ts`/`.tsx` under the batch tree, plus the worker entry. */
  function batchSources(dir: string): string[] {
    const out: string[] = [];
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      const full = resolve(dir, entry.name);
      if (entry.isDirectory()) out.push(...batchSources(full));
      else if (/\.tsx?$/.test(entry.name)) out.push(full);
    }
    return out;
  }

  const files = [
    ...batchSources(BATCH_SRC),
    resolve(ROOT, 'src/workers/compare.worker.ts'),
    resolve(ROOT, 'src/web/WebBatchCompare.tsx'),
    resolve(ROOT, 'src/web/WebBundleCompare.tsx'),
  ];

  it('covers the whole batch tree, not a sample of it', () => {
    expect(files.length).toBeGreaterThanOrEqual(10);
  });

  for (const file of files) {
    const name = file.slice(ROOT.length + 1);
    it(`${name} touches no host API and issues no network call`, () => {
      // Comments legitimately name the banned symbols to explain why they are
      // absent, so only executable code is checked.
      const code = readFileSync(file, 'utf8')
        .replace(/\/\*[\s\S]*?\*\//g, '')
        .replace(/\/\/[^\n]*/g, '');

      expect(/\bOffice\./.test(code)).toBe(false);
      expect(/\bWord\./.test(code)).toBe(false);
      expect(/\bfetch\s*\(/.test(code)).toBe(false);
      expect(/\bnew\s+(XMLHttpRequest|WebSocket|EventSource)\b/.test(code)).toBe(false);
      expect(/sendBeacon/.test(code)).toBe(false);
    });
  }
});

describe('M6 — no US date convention reaches the pairing sources', () => {
  // ADR-0002. The spec states this as an acceptance criterion rather than
  // leaving it to `no-us-literals`, because the lint rule catches literals in
  // strings and this is really a claim about the *patterns*: a normaliser that
  // stripped `MM/DD/YYYY` would silently accept a date order neither
  // jurisdiction uses, and would do it while looking like thoroughness.
  const PAIRING_SRC = resolve(BATCH_SRC, 'pairing');

  function pairingSources(): string[] {
    return readdirSync(PAIRING_SRC)
      .filter(n => /\.tsx?$/.test(n))
      .map(n => resolve(PAIRING_SRC, n));
  }

  it('has pairing sources to check', () => {
    expect(pairingSources().length).toBeGreaterThanOrEqual(6);
  });

  for (const file of pairingSources()) {
    const name = file.slice(ROOT.length + 1);
    it(`${name} contains no month-first date pattern`, () => {
      const code = readFileSync(file, 'utf8')
        // Comments name the forbidden form in order to forbid it.
        .replace(/\/\*[\s\S]*?\*\//g, '')
        .replace(/\/\/[^\n]*/g, '');

      // Escapes are stripped first: inside a regex literal the separator is
      // written `\/`, so a check that only matched a bare slash would pass over
      // the very place a month-first pattern would actually live. Getting that
      // wrong is how this test would look thorough and prove nothing — the
      // first draft of it did exactly that, and a mutation caught it.
      const bare = code.replace(/\\/g, '');

      // Three shapes a month-first convention takes in source. The middle one
      // is the one that matters: a regex whose 1–12 group comes BEFORE its
      // 1–31 group is a month-first parser however it is named. The live
      // patterns use `\d{1,2}` for both fields and are order-agnostic, so
      // neither group appears here at all.
      expect(/\bMM\s*[/\-.]\s*DD\b/i.test(bare)).toBe(false);
      expect(/1\[0-2\][^\n]{0,24}?[/\-.][^\n]{0,24}?3\[01\]/.test(bare)).toBe(false);
      // A month-first literal, disambiguated by a day no month can be.
      expect(/\b(0?[1-9]|1[0-2])\/(1[3-9]|2\d|3[01])\/\d{2,4}\b/.test(bare)).toBe(false);
    });
  }
});

describe('the pairing tunables are read from config, not from the constants', () => {
  // Spec §7 requires `compare.batch.pairing.*` to be resolved through the
  // config layer. A review found that although `resolveBatchConfig` DID resolve
  // all seven, nothing consumed the result: every pairing module imported the
  // raw constants, so the resolver only ever talked to itself. Setting
  // `autoPairScore` in a profile did nothing, and the test that covered the
  // resolver passed either way.
  //
  // `batch.config.test.ts` proves the values now reach the outcome. This
  // proves they cannot stop reaching it: the tunables may be imported by
  // `config.ts`, which owns the defaults, and nowhere else under the batch
  // tree. A future module that reaches for a constant instead of its settings
  // fails here rather than silently ignoring a profile.
  const TUNABLE = [
    'PAIRING_WEIGHT_CONTENT',
    'PAIRING_WEIGHT_FILENAME',
    'PAIRING_TIE_EPSILON',
    'PAIRING_AUTO_PAIR_SCORE',
    'PAIRING_AUTO_PAIR_MARGIN',
    'PAIRING_NO_MATCH_SCORE',
    'PAIRING_MIN_SKETCH_TOKENS',
    'PAIRING_NO_CONTENT_HEADROOM',
    'PAIRING_NO_CONTENT_MATCH_SCORE',
  ];

  function batchSourcesUnder(dir: string): string[] {
    const out: string[] = [];
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      const full = resolve(dir, entry.name);
      if (entry.isDirectory()) out.push(...batchSourcesUnder(full));
      else if (/\.tsx?$/.test(entry.name)) out.push(full);
    }
    return out;
  }

  const OWNER = resolve(BATCH_SRC, 'config.ts');
  const files = batchSourcesUnder(BATCH_SRC)
    .filter(f => f !== OWNER && f !== resolve(BATCH_SRC, 'constants.ts'));

  it('has files to check', () => {
    expect(files.length).toBeGreaterThanOrEqual(10);
  });

  for (const file of files) {
    const name = file.slice(ROOT.length + 1);
    it(`${name} imports no tunable pairing constant`, () => {
      const code = readFileSync(file, 'utf8')
        .replace(/\/\*[\s\S]*?\*\//g, '')
        .replace(/\/\/[^\n]*/g, '');

      // Only the import statements: a comment or a string that happens to name
      // one of these is not a use of it.
      const imports = code.match(/import\s[\s\S]*?from\s+'[^']+';/g) ?? [];
      for (const statement of imports) {
        for (const tunable of TUNABLE) {
          expect(statement.includes(tunable), `${name} imports ${tunable}`).toBe(false);
        }
      }
    });
  }
});

describe('nothing can run before confirmation — structurally, not just by policy', () => {
  // Guardrail 1 of the batch work. The confirmation gate is enforced twice already
  // (the button's `disabled`, and `runConfirmed`'s own refusal), but both live
  // downstream of the proposal. This is the third wall and the only static
  // one: the pairing tree cannot IMPORT anything that runs a comparison, so an
  // "auto-run when confidence is high" path cannot be added inside it without
  // first failing this test.
  //
  // A mis-pairing does not fail loudly. It produces a plausible-looking
  // whole-document rewrite — the pairing analogue of suppression — which is
  // why this gate gets a structural guard and the others do not.
  const PAIRING_SRC = resolve(BATCH_SRC, 'pairing');
  const FORBIDDEN = /from\s+'[^']*(?:\/|^)(runner|driver|bundle-driver|pipeline)\.js'/;

  for (const name of readdirSync(PAIRING_SRC).filter(n => /\.tsx?$/.test(n))) {
    it(`pairing/${name} imports no module that can start a comparison`, () => {
      const code = readFileSync(resolve(PAIRING_SRC, name), 'utf8');
      expect(FORBIDDEN.test(code)).toBe(false);
    });
  }

  it('and the grid component cannot either', () => {
    const code = readFileSync(
      resolve(BATCH_SRC, 'components/PairingGrid.tsx'),
      'utf8',
    );
    expect(FORBIDDEN.test(code)).toBe(false);
  });

  it('the bundle driver runs a pipeline from exactly one place', () => {
    // The wall above keeps the gate out of the pairing tree. This one keeps it
    // from being bypassed INSIDE the driver: a review found `openPair` starting
    // a comparison on an unconfirmed pairing, because the gate had been written
    // at `runConfirmed` rather than at the thing being gated. One call site
    // means one guard — a second `runComparePipeline(` here is a second door,
    // and the person adding it has to read this test to add it.
    const code = readFileSync(resolve(BATCH_SRC, 'bundle-driver.ts'), 'utf8');
    expect(code.match(/runComparePipeline\(/g) ?? []).toHaveLength(1);
  });
});

describe('I3 — nothing is suppressed anywhere in the batch path', () => {
  const CONFIG = getCompareConfig(UK_PROFILE);
  const TEXT = 'The Seller shall pay the Purchase Price on Completion.\nGoverned by English law.';

  function doc(id: string, text: string, label = `${id}.docx`): BatchDocumentInput {
    const raw = new Uint8Array(synthesiseDocx(text).toUint8Array());
    const buffer = raw.buffer.slice(
      raw.byteOffset,
      raw.byteOffset + raw.byteLength,
    ) as ArrayBuffer;
    return { id, label, byteLength: buffer.byteLength, bytes: () => Promise.resolve(buffer) };
  }

  it('reports a row for every pair requested, whatever happened to it', async () => {
    const variants = [
      doc('changed', TEXT.replace('on Completion.', 'within five days.')),
      doc('same', TEXT),
      { ...doc('broken', TEXT), bytes: () => Promise.resolve(new Uint8Array([1, 2]).buffer as ArrayBuffer) },
      doc('pdf', TEXT, 'execution-copy.pdf'),
    ];

    const index = await runOneToManyBatch({
      template: doc('template', TEXT, 'house-nda.docx'),
      variants,
      config: CONFIG,
      author: 'Mutatis Compare',
      date: '2026-09-10T09:00:00.000Z',
      flattenTemplate: 'acceptAll',
      flattenVariant: 'acceptAll',
      pool: createBatchPool([createInProcessWorker()]),
    }).index;

    // Four in, four out. Not "four in, the two that worked out".
    expect(index.rows).toHaveLength(4);
    expect(index.totals.pairs).toBe(4);
    expect(index.rows.map(r => r.status).sort())
      .toEqual(['done', 'done', 'failed', 'unrunnable']);

    // Every non-success carries a reason, so no row is a silent shrug.
    for (const row of index.rows.filter(r => r.status !== 'done')) {
      expect(row.reason).toBeTruthy();
    }
  });

  it('never reports a filtered count as a change count', async () => {
    const index = await runOneToManyBatch({
      template: doc('template', TEXT, 'house-nda.docx'),
      variants: [doc('a', TEXT.replace('on Completion.', 'within five Business Days.'))],
      config: CONFIG,
      author: 'Mutatis Compare',
      date: '2026-09-10T09:00:00.000Z',
      flattenTemplate: 'acceptAll',
      flattenVariant: 'acceptAll',
      pool: createBatchPool([createInProcessWorker()]),
    }).index;

    const row = index.rows[0];
    const byClass = Object.values(row?.byClassification ?? {}).reduce((n, v) => n + v, 0);

    // The classification tally covers the whole change set; if a filter had
    // reached this code the two would diverge.
    expect(byClass).toBe(row?.changeCount);
    expect(index.totals.changes).toBe(row?.changeCount);
  });

  it('keeps every pair obtainable even when the artefact ceiling degrades the batch', async () => {
    // The ceiling is a memory trade-off, not a filter. A row past it keeps no
    // Blob — but it is still compared, still counted, still disclosed, and its
    // redline is still downloadable by re-running that one pair. A ceiling that
    // quietly cost the user a hundred redlines would be suppression by another
    // name.
    const index = await runOneToManyBatch({
      template: doc('template', TEXT, 'house-nda.docx'),
      variants: [doc('a', TEXT.replace('on Completion.', 'within five days.'))],
      config: CONFIG,
      author: 'Mutatis Compare',
      date: '2026-09-10T09:00:00.000Z',
      flattenTemplate: 'acceptAll',
      flattenVariant: 'acceptAll',
      artefactBudgetBytes: 1,
      pool: createBatchPool([createInProcessWorker()]),
    }).index;

    const row = index.rows[0];
    expect(row?.status).toBe('done');
    expect(row?.changeCount).toBeGreaterThan(0);
    expect(row?.artefactOnDemand).toBe(true);
    expect(isDownloadable(row as BatchRow)).toBe(true);
    expect(index.disclosures).toHaveLength(1);
  });

  it('gives the index builder no parameter through which a change could be dropped', () => {
    // The structural half of the ADR-0008 wall: `buildCompareReport` has no
    // filter parameter and neither does this. A future filter flag would have
    // to be added deliberately, in a diff a reviewer sees.
    const source = readFileSync(resolve(BATCH_SRC, 'index-model.ts'), 'utf8');
    expect(source).not.toMatch(/hideFormattingOnly|collapseRenumbering|filters?\s*[:?]/);

    // And the type is satisfied by exactly the fields below — nothing else is accepted.
    const index = buildBatchIndex({
      mode: 'one-to-many',
      rows: [],
      generatedAt: '2026-09-10T09:00:00.000Z',
      flatten: { template: 'acceptAll', variant: 'acceptAll' },
    });
    expect(index.rows).toEqual([]);
    expect(index.totals.pairs).toBe(0);
  });
});
