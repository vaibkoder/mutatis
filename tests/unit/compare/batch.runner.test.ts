/**
 * Compare at scale — the one-to-many runner (ADR-0037; spec §5.1, §4.3).
 *
 * Acceptance B2, B3, B4, B6, B7, B10.
 *
 * The runner's job is structural, so the tests are structural too. Three
 * properties carry the feature, and each is invisible to every other test in
 * the suite:
 *
 *  - the template is preprocessed ONCE and its records are not mutated, because
 *    the whole point of the batch seam is that N runs share one preprocess;
 *  - no `PipelineResult` outlives its own pair, because retaining N of them is
 *    an out-of-memory crash and not a slowdown;
 *  - a bad file, or a cancel, costs exactly what it should and no more.
 *
 * Fixtures are synthesised `.docx` packages rather than files on disk, so a
 * "hundred-variant" batch is cheap and deterministic.
 */
import { describe, it, expect, vi } from 'vitest';

import { getCompareConfig, UK_PROFILE } from '@shared/config/index.js';
import { synthesiseDocx } from '@features/compare/io/text/synthesise-docx.js';
import {
  styleNumberedDocx,
  styleNumberedNumbering,
  HEAD,
  SUB_A,
  SUB_B,
  INSERTED,
  type StyledPara,
} from '../../fixtures/style-numbered-docx.js';
import {
  runOneToManyBatch,
  CROSS_FORMAT_REASON,
  type BatchDocumentInput,
} from '@features/compare/batch/runner.js';
import {
  createBatchPool,
  createInProcessWorker,
  type BatchPool,
} from '@features/compare/batch/worker/index.js';
import { BatchInputTooLargeError, MAX_BATCH_VARIANTS } from '@features/compare/batch/limits.js';
import * as preprocess from '@features/compare/preprocess/index.js';

/** The real `preprocessSide`, captured before any spy replaces it. */
const preprocessActual = preprocess.preprocessSide;

const CONFIG = getCompareConfig(UK_PROFILE);
const DATE = '2026-09-10T09:00:00.000Z';

const TEMPLATE_TEXT = [
  'The Seller shall pay the Purchase Price on Completion.',
  'The Buyer shall deliver the Disclosure Letter on Completion.',
  'This Agreement is governed by the laws of England and Wales.',
].join('\n');

function docx(text: string): ArrayBuffer {
  const bytes = new Uint8Array(synthesiseDocx(text).toUint8Array());
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

function doc(id: string, text: string, label = `${id}.docx`): BatchDocumentInput {
  const bytes = docx(text);
  return {
    id,
    label,
    byteLength: bytes.byteLength,
    bytes: () => Promise.resolve(bytes),
  };
}

const template = (): BatchDocumentInput => doc('template', TEMPLATE_TEXT, 'house-nda.docx');

/** The runner prices each artefact at the template's own byte length. */
const TEMPLATE_BYTES = template().byteLength;

function variant(id: string, changed: string): BatchDocumentInput {
  return doc(id, TEMPLATE_TEXT.replace('on Completion.', changed));
}

function pool(size = 2): BatchPool {
  return createBatchPool(Array.from({ length: size }, () => createInProcessWorker()));
}

function run(
  variants: readonly BatchDocumentInput[],
  extra: Partial<Parameters<typeof runOneToManyBatch>[0]> = {},
): ReturnType<typeof runOneToManyBatch> {
  return runOneToManyBatch({
    template: template(),
    variants,
    config: CONFIG,
    author: 'Mutatis Compare',
    date: DATE,
    flattenTemplate: 'acceptAll',
    flattenVariant: 'acceptAll',
    pool: pool(),
    ...extra,
  });
}

describe('B2 — the template is preprocessed once, whatever N is', () => {
  it('calls preprocessSide once for the template and once per variant', async () => {
    const spy = vi.spyOn(preprocess, 'preprocessSide');
    spy.mockClear();

    const variants = Array.from({ length: 6 }, (_, i) => variant(`v${i}`, `within ${i} days.`));
    const index = await run(variants).index;

    // 1 template + 6 variants. Anything more means runPipeline() crept back in
    // and the template is being re-flattened and re-extracted per variant.
    expect(spy).toHaveBeenCalledTimes(7);
    expect(index.rows).toHaveLength(6);
    spy.mockRestore();
  });

  it('produces a row for every variant, in submission order', async () => {
    const variants = ['a', 'b', 'c'].map(id => variant(id, `by ${id}.`));
    const index = await run(variants).index;

    expect(index.rows.map(r => r.id)).toEqual(['a', 'b', 'c']);
    expect(index.templateLabel).toBe('house-nda.docx');
    expect(index.mode).toBe('one-to-many');
  });
});

describe('B3 — the shared template records are read, never written', () => {
  it('leaves the template ParaRecord[] deep-equal after a batch', async () => {
    // One array is shared by every job. If any stage mutated it, variant 2
    // onwards would be diffed against a template that earlier runs had edited —
    // silently, and differently depending on ordering.
    const spy = vi.spyOn(preprocess, 'preprocessSide');
    let templateParas: unknown = null;
    let snapshot = '';
    spy.mockImplementation((...args) => {
      const result = preprocessActual(...args);
      if (templateParas === null) {
        templateParas = result.paras;
        snapshot = JSON.stringify(result.paras);
      }
      return result;
    });

    await run(['a', 'b', 'c', 'd'].map(id => variant(id, `by ${id}.`))).index;

    expect(templateParas).not.toBeNull();
    expect(JSON.stringify(templateParas)).toBe(snapshot);
    spy.mockRestore();
  });

  it('gives every variant the same comparison it would get alone', async () => {
    // The real mutation test: an identical variant placed LAST must still
    // report zero changes. If an earlier run had edited the shared template,
    // it would not.
    const variants = [
      variant('changed-1', 'within five Business Days of Completion.'),
      variant('changed-2', 'within ten Business Days of Completion.'),
      doc('identical', TEMPLATE_TEXT),
    ];

    const index = await run(variants, { pool: pool(1) }).index;

    expect(index.rows.find(r => r.id === 'identical')?.changeCount).toBe(0);
    expect(index.rows.find(r => r.id === 'changed-1')?.changeCount).toBeGreaterThan(0);
  });
});

describe('B9 — pooled and in-process runs agree at row level', () => {
  it('produces identical rows however many workers ran them', async () => {
    // `batch.worker.test.ts` proves the two transports agree on one job's
    // output. This is the other half of B9: that the RUNNER, fanning the same
    // variants across a different number of lanes, still reports the same rows.
    const variants = [
      variant('a', 'by Friday.'),
      variant('b', 'within five Business Days of Completion.'),
      doc('identical', TEMPLATE_TEXT),
    ];

    const single = await run(variants, { pool: pool(1) }).index;
    const many = await run(variants, { pool: pool(3) }).index;

    const compact = (i: typeof single) => i.rows.map(r => ({
      id: r.id,
      status: r.status,
      changeCount: r.changeCount,
      byClassification: r.byClassification,
      artefactBytes: r.artefacts[0]?.byteLength ?? 0,
    }));

    expect(compact(many)).toEqual(compact(single));
    expect(many.totals).toEqual(single.totals);
  });
});

describe('duplicate ids are refused rather than silently collapsed', () => {
  it('throws before any pair runs', () => {
    // Two rows keyed the same would overwrite one another: the index would show
    // the right NUMBER of rows with one comparison's result lost.
    expect(() => run([variant('a', 'by Friday.'), variant('a', 'by Monday.')]))
      .toThrow(/duplicate document id/i);
  });
});

describe('B4 — no pipeline result outlives its own pair', () => {
  it('retains a compact row carrying no paragraph records', async () => {
    const index = await run([variant('a', 'by Friday.')]).index;
    const row = index.rows[0];

    expect(row).toBeDefined();
    expect(row).not.toHaveProperty('ops');
    expect(row).not.toHaveProperty('originalParas');
    expect(row).not.toHaveProperty('revisedParas');
    expect(row).not.toHaveProperty('authoredBodyXml');
    // What it does keep is the artefact — a Blob, which is browser-managed
    // storage rather than JS heap.
    expect(row?.artefacts[0]?.kind).toBe('redlineDocx');
    expect(row?.artefacts[0]?.byteLength).toBeGreaterThan(0);
  });

  it('holds no reference to any outcome the pool handed it', async () => {
    // The assertion that can actually fail. Marking every object inside the
    // outcome and then walking the finished index for those marks proves the
    // runner kept none of them — whereas checking a row for an `ops` key only
    // restates the object literal the runner writes.
    const marked = new Set<unknown>();
    const markDeep = (value: unknown, depth = 0): void => {
      if (depth > 6 || value === null || typeof value !== 'object') return;
      if (marked.has(value)) return;
      marked.add(value);
      for (const child of Object.values(value as Record<string, unknown>)) markDeep(child, depth + 1);
    };

    const inner = createBatchPool([createInProcessWorker()]);
    const spy: BatchPool = {
      size: inner.size,
      submit(job, opts) {
        const submitted = inner.submit(job, opts);
        return { id: submitted.id, outcome: submitted.outcome.then(o => { markDeep(o); return o; }) };
      },
      cancel: id => inner.cancel(id),
      cancelAll: () => inner.cancelAll(),
      terminate: () => inner.terminate(),
    };

    const index = await run([variant('a', 'by Friday.'), variant('b', 'by Monday.')], { pool: spy }).index;
    expect(marked.size).toBeGreaterThan(0);

    const found: string[] = [];
    const walk = (value: unknown, path: string, depth = 0): void => {
      if (depth > 8 || value === null || typeof value !== 'object') return;
      // A Blob is browser-managed storage, not JS heap — it is the thing the
      // runner is SUPPOSED to keep, and it is opaque to this walk anyway.
      if (value instanceof Blob) return;
      if (marked.has(value)) { found.push(path); return; }
      for (const [key, child] of Object.entries(value as Record<string, unknown>)) {
        walk(child, `${path}.${key}`, depth + 1);
      }
    };
    walk(index, 'index');

    expect(found).toEqual([]);
  });

  it('keeps comparing but stops retaining once the artefact ceiling is reached', async () => {
    // Spec §4.3 step 6. The ceiling degrades; it never refuses and never
    // silently drops a result — every pair is still compared, still counted,
    // and still obtainable, and the batch says so.
    const variants = [
      variant('a', 'by Friday.'),
      variant('b', 'by Monday.'),
      variant('c', 'by Tuesday.'),
    ];

    const index = await run(variants, { pool: pool(1), artefactBudgetBytes: 1 }).index;

    // Every pair ran and every pair is reported.
    expect(index.rows).toHaveLength(3);
    expect(index.totals.done).toBe(3);
    for (const row of index.rows) expect(row.changeCount).toBeGreaterThan(0);

    // Nothing was retained, every row says so, and the batch discloses it.
    expect(index.rows.every(r => r.artefacts.length === 0)).toBe(true);
    expect(index.rows.every(r => r.artefactOnDemand)).toBe(true);
    expect(index.disclosures).toHaveLength(1);
    expect(index.disclosures[0]).toMatch(/on demand/i);
  });

  it('retains what fits and flips only the rows past the ceiling', async () => {
    // The mixed case is the one that matters: the ceiling must not be
    // all-or-nothing, or the first oversized batch loses every artefact.
    const one = await run([variant('probe', 'by Friday.')], { pool: pool(1) }).index;
    const size = one.rows[0]?.artefacts[0]?.byteLength ?? 0;
    expect(size).toBeGreaterThan(0);

    // A budget generous enough for the first row's estimate and no more. The
    // estimate is the template's own byte length, so price it off that.
    const index = await run(
      [variant('a', 'by Friday.'), variant('b', 'by Monday.'), variant('c', 'by Tuesday.')],
      { pool: pool(1), artefactBudgetBytes: TEMPLATE_BYTES },
    ).index;

    const retained = index.rows.filter(r => r.artefacts.length > 0);
    const onDemand = index.rows.filter(r => r.artefactOnDemand);

    expect(retained.length).toBeGreaterThanOrEqual(1);
    expect(onDemand.length).toBeGreaterThanOrEqual(1);
    expect(retained.length + onDemand.length).toBe(3);
    expect(index.disclosures).toHaveLength(1);
  });

  it('reads each variant only once, and only when it is its turn', async () => {
    const readOrder: string[] = [];
    const lazy = (id: string): BatchDocumentInput => {
      const base = variant(id, `by ${id}.`);
      return { ...base, bytes: () => { readOrder.push(id); return base.bytes(); } };
    };

    await run([lazy('a'), lazy('b'), lazy('c')], { pool: pool(1) }).index;

    expect(readOrder).toEqual(['a', 'b', 'c']);
    expect(new Set(readOrder).size).toBe(readOrder.length);
  });
});

describe('B6 — cancelling keeps what finished', () => {
  it('marks unreached variants cancelled and keeps completed rows with their artefacts', async () => {
    const variants = Array.from({ length: 8 }, (_, i) => variant(`v${i}`, `by day ${i}.`));

    // Cancel deterministically the moment the second pair completes, rather
    // than on a timer — these fixtures are small enough that a timer races the
    // whole batch to completion and proves nothing.
    let completed = 0;
    let batch: ReturnType<typeof run>;
    batch = run(variants, {
      pool: pool(1),
      onRowChanged: row => {
        if (row.status !== 'done') return;
        completed += 1;
        if (completed === 2) batch.cancel();
      },
    });

    const index = await batch.index;

    expect(completed).toBe(2);

    // Discarding sixty finished redlines because the user stopped the
    // sixty-first would be user-hostile; §9.12's "discard partial state"
    // governs the in-flight pair, not the batch.
    expect(index.rows).toHaveLength(8);
    expect(index.totals.cancelled).toBeGreaterThan(0);
    expect(index.totals.done + index.totals.cancelled).toBe(8);
    for (const row of index.rows.filter(r => r.status === 'done')) {
      expect(row.artefacts.length).toBe(1);
    }
  });

  it('reports every pair as cancelled when stopped before anything runs', async () => {
    const variants = ['a', 'b', 'c'].map(id => variant(id, `by ${id}.`));
    const batch = run(variants, { pool: pool(1) });
    batch.cancel();

    const index = await batch.index;
    expect(index.rows.every(r => r.status === 'cancelled')).toBe(true);
    expect(index.totals.pairs).toBe(3);
  });
});

describe('B7 — one bad file does not cost the rest of the batch', () => {
  it('records a failed row with a reason and carries on', async () => {
    const corrupt: BatchDocumentInput = {
      id: 'corrupt',
      label: 'corrupt.docx',
      byteLength: 8,
      bytes: () => Promise.resolve(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]).buffer as ArrayBuffer),
    };

    const index = await run([
      variant('good-1', 'by Monday.'),
      corrupt,
      variant('good-2', 'by Tuesday.'),
    ]).index;

    const failed = index.rows.find(r => r.id === 'corrupt');
    expect(failed?.status).toBe('failed');
    expect(failed?.reason).toBeTruthy();

    expect(index.rows.filter(r => r.status === 'done')).toHaveLength(2);
    expect(index.totals.failed).toBe(1);
    expect(index.totals.done).toBe(2);
  });

  it('marks a non-docx variant unrunnable, with its pairing still reported', async () => {
    const index = await run([
      variant('a', 'by Monday.'),
      doc('scan', TEMPLATE_TEXT, 'execution-copy.pdf'),
    ]).index;

    const unrunnable = index.rows.find(r => r.id === 'scan');
    expect(unrunnable?.status).toBe('unrunnable');
    expect(unrunnable?.reason).toBe(CROSS_FORMAT_REASON);
    // Identified, never silently skipped.
    expect(index.rows).toHaveLength(2);
    expect(index.totals.unrunnable).toBe(1);
  });
});

describe('B10 — an unchanged variant is a result, not an absence', () => {
  it('reports zero changes and still produces an artefact', async () => {
    const index = await run([doc('same', TEMPLATE_TEXT)]).index;
    const row = index.rows[0];

    expect(row?.status).toBe('done');
    expect(row?.changeCount).toBe(0);
    expect(row?.artefacts).toHaveLength(1);
  });
});

describe('caps are enforced before any byte is read', () => {
  it('refuses an oversize batch without reading a single variant', () => {
    let reads = 0;
    const many = Array.from({ length: MAX_BATCH_VARIANTS + 1 }, (_, i) => {
      const base = variant(`v${i}`, 'by Friday.');
      return { ...base, bytes: () => { reads += 1; return base.bytes(); } };
    });

    expect(() => run(many)).toThrow(BatchInputTooLargeError);
    expect(reads).toBe(0);
  });

  it('names the offending document for a per-document cap', () => {
    const huge = { ...variant('huge', 'by Friday.'), byteLength: 999 * 1024 * 1024 };
    expect(() => run([huge])).toThrow(/huge\.docx/);
  });
});

describe('progress and classification', () => {
  it('reports monotonic batch progress ending at 100', async () => {
    const pcts: number[] = [];
    await run(['a', 'b', 'c'].map(id => variant(id, `by ${id}.`)), {
      pool: pool(1),
      onProgress: p => pcts.push(p.pct),
    }).index;

    expect(pcts.at(-1)).toBe(100);
    expect(pcts.every((p, i) => i === 0 || p >= (pcts[i - 1] as number))).toBe(true);
  });

  it('emits each row as it changes state', async () => {
    const seen: string[] = [];
    await run([variant('a', 'by Friday.')], {
      onRowChanged: r => seen.push(`${r.id}:${r.status}`),
    }).index;

    expect(seen).toContain('a:running');
    expect(seen).toContain('a:done');
  });

  it('tallies classifications over the full change set, not a filtered one', async () => {
    const index = await run([variant('a', 'within five Business Days of Completion.')]).index;
    const row = index.rows[0];
    const summed = Object.values(row?.byClassification ?? {}).reduce((n, v) => n + v, 0);

    expect(summed).toBe(row?.changeCount);
  });
});

/**
 * ADR-0042 at the batch seam.
 *
 * The three *comparison* seams (`batch/driver.ts`, `batch/runner.ts`,
 * `batch/pairing/sketch.ts`) had to learn about the numbering context, or a 1:N
 * run would report fewer changes than the single-pair panel reports on the same
 * two files — the same documents, two different answers, with nothing telling
 * the reader which is right. Nothing asserted the wiring, so deleting it broke
 * no test.
 *
 * The shape used here is the one that motivated ADR-0042 in the first place: a
 * firm precedent where `document.xml` carries NO `w:numPr` at all and every
 * clause number is inherited from its `w:pStyle`. The variant's text is
 * identical; only the style differs. Without the numbering context the two
 * paragraphs compare equal on every field the batch reads, and `changeCount` is
 * zero.
 */
describe('ADR-0042 — the numbering context reaches the batch seam', () => {
  function styledDoc(
    id: string, paras: readonly StyledPara[], numberingXml?: string,
  ): BatchDocumentInput {
    const bytes = styleNumberedDocx(paras, numberingXml);
    return {
      id,
      label: `${id}.docx`,
      byteLength: bytes.byteLength,
      bytes: () => Promise.resolve(bytes),
    };
  }

  async function batch(
    template: readonly StyledPara[], variant: readonly StyledPara[],
    schemes?: { template: string; variant: string },
  ) {
    const index = await runOneToManyBatch({
      template: styledDoc('template', template, schemes?.template),
      variants: [styledDoc('v1', variant, schemes?.variant)],
      config: CONFIG,
      author: 'Mutatis Compare',
      date: DATE,
      flattenTemplate: 'acceptAll',
      flattenVariant: 'acceptAll',
      pool: pool(),
    }).index;
    expect(index.rows).toHaveLength(1);
    expect(index.rows[0]!.status).toBe('done');
    return index.rows[0]!;
  }

  it('resolves numbering on BOTH sides, so a pure renumbering reads as one', async () => {
    // A sub-clause inserted at the head: `(a)`→`(b)` and `(b)`→`(c)` for the two
    // that follow. Identity unchanged, so §9.4's answer is `crossRefRenumber`.
    //
    // This is what isolates the wiring rather than merely exercising it. Drop the
    // numbering context from either side and that side's paragraphs resolve to NO
    // numbering at all, so every matched pair differs in identity and the same
    // edit reads as `formatting` — the batch giving a different answer from the
    // single-pair panel on the same two files, which is what ADR-0042 wired these
    // three seams to prevent.
    const row = await batch([HEAD, SUB_A, SUB_B], [HEAD, INSERTED, SUB_A, SUB_B]);

    expect(row.byClassification.crossRefRenumber).toBe(2);
    expect(row.byClassification.formatting).toBe(0);
    expect(row.changeCount).toBe(3); // two shifted labels + the insert
  });

  it('reports a style-inherited re-levelling AND what it drags with it', async () => {
    // The shape that motivated ADR-0042: `document.xml` has no `w:numPr`
    // anywhere, the words are identical, and only the `w:pStyle` differs.
    //
    // Three paragraphs, not two, and that is what makes this a guard rather
    // than a demonstration. Outdenting `SUB_A` also renumbers `SUB_B` from
    // `(b)` to `(a)` — a paragraph whose own style never moved. A `BatchRow`
    // carries counts and no ops, so the re-levelled paragraph alone cannot
    // isolate the wiring: its two `w:pStyle` values differ with or without a
    // numbering context, and a two-paragraph version of this test passes with
    // the context deleted. The bystander is the half that only exists when the
    // numbering resolves.
    const row = await batch([HEAD, SUB_A, SUB_B], [HEAD, ['Clause1', SUB_A[1]], SUB_B]);

    expect(row.changeCount).toBe(2);
    expect(row.byClassification.formatting).toBe(1);      // the outdented clause
    expect(row.byClassification.crossRefRenumber).toBe(1); // the one it dragged
  });

  it('reports nothing when template and variant are the same document', async () => {
    const row = await batch([HEAD, SUB_A, SUB_B], [HEAD, SUB_A, SUB_B]);
    expect(row.changeCount).toBe(0);
  });

  it('reads the VARIANT\'s own list definitions, not the template\'s twice', async () => {
    // The tests above prove a numbering context reaches the variant's
    // `preprocessSide`; they cannot prove it is the variant's own, because the
    // fixture gave both sides the same `numbering.xml`. Passing the template's
    // context for the variant — one token at `runner.ts` — made a list
    // REDEFINITION report nothing at all, with the whole suite green. A
    // `BatchRow` carries counts and no ops, so the count IS the assertion here.
    //
    // Level 0 redefined `decimal` → `upperRoman` under `%1.%2`, so the
    // ancestor's format reaches every descendant: three paragraphs, three
    // `formatting` rows (§8.2’s filter hides them by default; the same section
    // still counts them in the total).
    const paras = [HEAD, SUB_A, SUB_B];
    const row = await batch(paras, paras, {
      template: styleNumberedNumbering({ subText: '%1.%2' }),
      variant: styleNumberedNumbering({ headFmt: 'upperRoman', subText: '%1.%2' }),
    });

    expect(row.changeCount).toBe(3);
    expect(row.byClassification.formatting).toBe(3);
  });
});
