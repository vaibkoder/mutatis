/**
 * The one-to-many batch runner (ADR-0037; spec §5.1, §4.3).
 *
 * Everything structural about this feature lives here, so the three rules it
 * exists to enforce are worth stating plainly:
 *
 *  1. **The template is preprocessed once.** `runComparePipeline` takes
 *     already-built `ParaRecord[]`, which is the seam that makes this possible.
 *     Calling `runPipeline` per variant would re-flatten, re-quarantine and
 *     re-extract the template N times — the exact waste the feature removes.
 *  2. **N results are never retained.** Each `PipelineResult` holds two
 *     document-sized `ParaRecord[]` (every record carries its paragraph's
 *     `rawXml`), so keeping N of them is an out-of-memory crash rather than a
 *     slowdown. A completed pair is turned into a `Blob` artefact — browser-
 *     managed storage the engine may spill to disk — and a compact row, and
 *     everything else is dropped on the spot.
 *  3. **Ingest is bounded, not eager.** Variants are read one at a time, and no
 *     more than `pool.size` of them are alive at once. Reading all N up front
 *     would breach (2) before the first comparison had run.
 *
 * Pure of hosts: no `Office.*`/`Word.*`, no `fetch`. The only browser API is
 * `Blob`, and the only I/O is the caller's own byte-reader callback.
 */
import { PizZipBundle } from '@shared/ooxml/bundle.js';
import type { CompareConfig, FlattenPolicy } from '@shared/config/index.js';
import { preprocessSide, buildNumberingContext } from '../preprocess/index.js';
import { extractParts } from '../io/parts.js';
import { assembleRedlinePackage } from '../output/writeout.js';
import type { ParaRecord } from '../model.js';
import { buildBatchIndex } from './index-model.js';
import { emptyClassificationTally } from './model.js';
import type { ArtefactHandle, BatchIndex, BatchRow } from './model.js';
import {
  artefactBudgetAllows,
  checkDocxByteCap,
  checkTotalByteCap,
  describeDocxCapBreach,
  exceedsDocxByteCap,
  checkBundleCap,
  checkVariantCap,
} from './limits.js';
import { createBatchPoolForHost, type BatchPool } from './worker/index.js';
import { resolveBatchConfig } from './config.js';
import { CROSS_FORMAT_REASON } from './constants.js';

export { CROSS_FORMAT_REASON };

const DOCX_MIME =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

/** One document handed to the runner. Bytes are read lazily, one at a time. */
export interface BatchDocumentInput {
  id: string;
  label: string;
  byteLength: number;
  /** Called at most once, when the runner is ready to spend memory on it. */
  bytes(): Promise<ArrayBuffer>;
}

export interface BatchProgress {
  done: number;
  total: number;
  /** 0–100 across the whole batch, counting in-flight pairs' own progress. */
  pct: number;
}

export interface RunOneToManyInput {
  template: BatchDocumentInput;
  variants: readonly BatchDocumentInput[];
  config: CompareConfig;
  author: string;
  /** ISO-8601, injected so every redline in a batch shares one date. */
  date: string;
  flattenTemplate: FlattenPolicy;
  flattenVariant: FlattenPolicy;
  /** Injected by tests; otherwise the host's pool (or the in-process fallback). */
  pool?: BatchPool;
  /**
   * Cumulative retained-artefact ceiling. Defaults to `MAX_BATCH_ARTEFACT_BYTES`.
   * Past it the runner keeps comparing and keeps reporting, but stops holding
   * `Blob`s — later rows are marked on-demand and a disclosure is added. Never a
   * refusal, and never silent.
   */
  artefactBudgetBytes?: number;
  onRowChanged?: (row: BatchRow) => void;
  onProgress?: (progress: BatchProgress) => void;
}

/**
 * One pair, reduced to what a lane needs.
 *
 * `prepare()` is where the two modes differ and the only place they differ:
 * one-to-many reads and preprocesses the variant against a template that was
 * preprocessed once; many-to-many preprocesses both sides, because a pairing
 * sketch carries hashes and not records (spec §6.7, §4.3). Either way it is
 * called by the lane that runs the pair, so only the pool's worth of working
 * sets exist at a time.
 *
 * `assemble` is returned from `prepare` rather than taken as a parameter so the
 * original bundle stays scoped to the pair that needs it. A batch must never
 * hold M document packages at once.
 */
export interface PreparedPair {
  originalParas: readonly ParaRecord[];
  revisedParas: readonly ParaRecord[];
  quarantinedOriginal: number;
  quarantinedRevised: number;
  /** Byte estimate for the redline this pair will produce. */
  artefactEstimate: number;
  /**
   * Build the redline package. Called ONLY if the artefact budget allows it —
   * which is why it may be async: a mode whose original document is not already
   * in memory reads it here, and past the ceiling that read never happens at
   * all. Doing it in `prepare` instead would read every document in the bundle
   * to produce nothing.
   */
  assemble(authoredBodyXml: string): PizZipBundle | Promise<PizZipBundle>;
}

interface LanePair {
  id: string;
  label: string;
  /** Set when the pair is real but must not be compared. */
  unrunnableReason?: string;
  /** File name for the artefact, before the " (redline)" suffix. */
  artefactStem: string;
  prepare(): Promise<PreparedPair>;
}

export interface BatchRun {
  /** Resolves when every pair has reached a terminal state. Never rejects for a per-pair failure. */
  index: Promise<BatchIndex>;
  /** Stop the batch. Completed pairs and their artefacts are kept. */
  cancel(): void;
}


const ARTEFACT_BUDGET_DISCLOSURE =
  'The retained-artefact limit was reached during this batch. Redlines for the ' +
  'variants marked "on demand" are produced when you download them, not held in memory.';

function isDocx(label: string): boolean {
  return label.toLowerCase().endsWith('.docx');
}

/**
 * Rows are keyed by id, so two variants sharing one would write to a single
 * entry: the index would still show the right NUMBER of rows while one
 * comparison's result had been overwritten by the other's. That is a silent
 * loss of a result, which is the batch-scale shape of the thing ADR-0008
 * forbids — so the runner refuses rather than trusting its caller. (The web
 * driver mints `slot-N` and cannot trip this; the runner is exported API and
 * defends its own invariant.)
 */
function assertDistinctIds(docs: readonly { id: string }[]): void {
  const seen = new Set<string>();
  for (const doc of docs) {
    if (seen.has(doc.id)) {
      throw new Error(
        `Duplicate document id "${doc.id}". Every document in a batch needs its own id.`,
      );
    }
    seen.add(doc.id);
  }
}

function pendingRow(doc: { id: string; label: string }): BatchRow {
  return {
    id: doc.id,
    label: doc.label,
    status: 'queued',
    changeCount: 0,
    byClassification: emptyClassificationTally(),
    quarantinedOriginal: 0,
    quarantinedRevised: 0,
    artefacts: [],
    artefactOnDemand: false,
    durationMs: 0,
  };
}

/** What the shared core needs to run a list of pairs. */
interface BatchCoreOptions {
  pool: BatchPool;
  /** Every pair the batch was asked to run, in order. One row each. */
  ids: readonly { id: string; label: string }[];
  config: CompareConfig;
  author: string;
  date: string;
  artefactBudgetBytes?: number;
  onRowChanged?: (row: BatchRow) => void;
  onProgress?: (progress: BatchProgress) => void;
}

/**
 * The batch machinery, shared by both modes.
 *
 * Everything that makes a batch a batch lives here — bounded concurrency, one
 * result at a time, the artefact budget, cancellation, and a row for every pair
 * whatever became of it. Neither mode gets its own copy, because a second copy
 * is where the two would drift on exactly the properties that are hardest to
 * notice going wrong: a batch that quietly retained N results, or quietly lost
 * a row, looks fine until it is far too large to debug.
 *
 * What the modes supply is only how a pair gets its records and its package
 * (`LanePair.prepare`).
 */
function createBatchCore(opts: BatchCoreOptions) {
  const pool = opts.pool;
  const rows = new Map<string, BatchRow>(
    opts.ids.map(d => [d.id, pendingRow(d)] as const),
  );
  const order = opts.ids.map(d => d.id);
  const inFlightPct = new Map<string, number>();
  const disclosures: string[] = [];

  let cancelled = false;
  let retainedArtefactBytes = 0;
  let done = 0;

  const emitRow = (row: BatchRow): void => {
    rows.set(row.id, row);
    opts.onRowChanged?.(row);
  };

  const emitProgress = (): void => {
    const total = order.length;
    if (total === 0) {
      opts.onProgress?.({ done: 0, total: 0, pct: 100 });
      return;
    }
    let partial = 0;
    for (const pct of inFlightPct.values()) partial += pct / 100;
    opts.onProgress?.({
      done,
      total,
      pct: Math.min(100, Math.round(((done + partial) / total) * 100)),
    });
  };

  const finish = (row: BatchRow): void => {
    inFlightPct.delete(row.id);
    done += 1;
    emitRow(row);
    emitProgress();
  };

  /**
   * The lane machinery, shared by both modes.
   *
   * Everything that makes a batch a batch lives here — bounded concurrency,
   * one result at a time, the artefact budget, cancellation, a row for every
   * pair whatever became of it. Neither mode gets its own copy of it, because a
   * second copy is where the two would drift on exactly the properties that are
   * hardest to notice going wrong.
   */
  const runPairs = async (pairs: readonly LanePair[]): Promise<void> => {
    let next = 0;
    const takeNext = (): LanePair | null => {
      if (cancelled) return null;
      const pair = pairs[next];
      if (pair === undefined) return null;
      next += 1;
      return pair;
    };

    const lane = async (): Promise<void> => {
      for (let pair = takeNext(); pair !== null; pair = takeNext()) {
        const startedAt = Date.now();
        const base = rows.get(pair.id) as BatchRow;

        if (pair.unrunnableReason !== undefined) {
          finish({ ...base, status: 'unrunnable', reason: pair.unrunnableReason });
          continue;
        }

        emitRow({ ...base, status: 'running' });
        inFlightPct.set(pair.id, 0);
        emitProgress();

        try {
          const prepared = await pair.prepare();

          if (cancelled) {
            finish({ ...base, status: 'cancelled' });
            continue;
          }

          const job = pool.submit(
            {
              originalParas: prepared.originalParas as ParaRecord[],
              revisedParas: prepared.revisedParas as ParaRecord[],
              config: opts.config,
              author: opts.author,
              date: opts.date,
              quarantinedOriginal: prepared.quarantinedOriginal,
              quarantinedRevised: prepared.quarantinedRevised,
            },
            {
              onProgress: pct => {
                inFlightPct.set(pair.id, pct);
                emitProgress();
              },
            },
          );

          const outcome = await job.outcome;

          // Decide the budget BEFORE spending the memory rather than after.
          const affordable = artefactBudgetAllows(
            retainedArtefactBytes,
            prepared.artefactEstimate,
            opts.artefactBudgetBytes,
          );

          const artefacts: ArtefactHandle[] = [];
          if (affordable) {
            const pkg = await prepared.assemble(outcome.authoredBodyXml);
            const bytes = new Uint8Array(pkg.toUint8Array());
            artefacts.push({
              kind: 'redlineDocx',
              fileName: `${pair.artefactStem} (redline).docx`,
              byteLength: bytes.byteLength,
              blob: new Blob([bytes], { type: DOCX_MIME }),
            });
            retainedArtefactBytes += bytes.byteLength;
          } else if (!disclosures.includes(ARTEFACT_BUDGET_DISCLOSURE)) {
            disclosures.push(ARTEFACT_BUDGET_DISCLOSURE);
          }

          finish({
            ...base,
            status: 'done',
            // `summary.total`, never a filtered count (ADR-0008).
            changeCount: outcome.summary.total,
            // Copied, not aliased: a row is plain, cloneable, long-lived data,
            // and holding a live reference into the outcome would let a later
            // mutation there change a row that has already been reported.
            byClassification: { ...outcome.summary.byCategory },
            quarantinedOriginal: outcome.quarantinedOriginal,
            quarantinedRevised: outcome.quarantinedRevised,
            artefacts,
            artefactOnDemand: !affordable,
            durationMs: Date.now() - startedAt,
          });
          // `outcome` and everything `prepare` built fall out of scope here.
          // Nothing document-sized survives the iteration.
        } catch (err) {
          // Only an actual abort is a cancellation. Treating every error as one
          // once `cancelled` is set would relabel a pair that failed on a
          // corrupt package microseconds before the user hit stop, losing the
          // reason it failed for.
          if (err instanceof DOMException && err.name === 'AbortError') {
            finish({ ...base, status: 'cancelled', durationMs: Date.now() - startedAt });
            continue;
          }
          // One bad file must not cost the user the other ninety-nine redlines.
          finish({
            ...base,
            status: 'failed',
            reason: err instanceof Error ? err.message : String(err),
            durationMs: Date.now() - startedAt,
          });
        }
      }
    };

    emitProgress();
    await Promise.all(
      Array.from({ length: Math.max(1, Math.min(pool.size, pairs.length || 1)) }, lane),
    );

    // Anything the lanes never reached because the batch was cancelled.
    for (const id of order) {
      const row = rows.get(id) as BatchRow;
      if (row.status === 'queued' || row.status === 'running') {
        rows.set(id, { ...row, status: 'cancelled' });
        opts.onRowChanged?.(rows.get(id) as BatchRow);
      }
    }
  };


  return {
    runPairs,
    cancel: (): void => { cancelled = true; },
    buildIndex: (extra: {
      mode: 'one-to-many' | 'many-to-many';
      templateLabel?: string;
      flatten: { template: FlattenPolicy; variant: FlattenPolicy };
    }): BatchIndex => buildBatchIndex({
      mode: extra.mode,
      ...(extra.templateLabel !== undefined ? { templateLabel: extra.templateLabel } : {}),
      rows: order.map(id => rows.get(id) as BatchRow),
      generatedAt: opts.date,
      // Spec §8 case 7: the flatten choice applies per batch, not per pair, so
      // the index states which one every row in it was produced under.
      flatten: extra.flatten,
      disclosures,
    }),
  };
}

/**
 * Run one template against N variants.
 *
 * Caps are checked before a single byte is read, so an oversize batch costs
 * nothing and — critically — is refused whole rather than quietly trimmed to
 * fit (`limits.ts`).
 */
export function runOneToManyBatch(input: RunOneToManyInput): BatchRun {
  // `compare.batch.*` (spec §7): the caps and the pool size come from the
  // resolved config, so a profile can tighten them without a code change.
  const settings = resolveBatchConfig(input.config);

  checkVariantCap(input.variants.length, settings.maxVariants);
  assertDistinctIds(input.variants);
  for (const v of input.variants) checkDocxByteCap(v.byteLength, v.label);
  checkDocxByteCap(input.template.byteLength, input.template.label);
  checkTotalByteCap(
    input.template.byteLength + input.variants.reduce((n, v) => n + v.byteLength, 0),
  );

  const pool = input.pool ?? createBatchPoolForHost({ poolSize: settings.poolSize });
  const ownsPool = input.pool === undefined;

  const core = createBatchCore({
    pool,
    ids: input.variants.map(v => ({ id: v.id, label: v.label })),
    config: input.config,
    author: input.author,
    date: input.date,
    ...(input.artefactBudgetBytes !== undefined
      ? { artefactBudgetBytes: input.artefactBudgetBytes }
      : {}),
    ...(input.onRowChanged !== undefined ? { onRowChanged: input.onRowChanged } : {}),
    ...(input.onProgress !== undefined ? { onProgress: input.onProgress } : {}),
  });

  const run = async (): Promise<BatchIndex> => {
    try {
      // --- the template, once ----------------------------------------------
      const templateBundle = new PizZipBundle(await input.template.bytes());
      const templateParts = extractParts(templateBundle);
      const template = preprocessSide(
        templateParts.documentXml,
        'document',
        input.flattenTemplate,
        input.config,
        // ADR-0042 — the numbering context, on both sides. See `batch/driver.ts`.
        buildNumberingContext(templateParts.stylesXml, templateParts.numberingXml),
      );
      const templateParas: ParaRecord[] = template.paras;

      await core.runPairs(input.variants.map(doc => ({
        id: doc.id,
        label: doc.label,
        artefactStem: doc.label.replace(/\.docx$/i, ''),
        ...(isDocx(doc.label) ? {} : { unrunnableReason: CROSS_FORMAT_REASON }),
        prepare: async (): Promise<PreparedPair> => {
          const bundle = new PizZipBundle(await doc.bytes());
          const variantParts = extractParts(bundle);
          const variant = preprocessSide(
            variantParts.documentXml,
            'document',
            input.flattenVariant,
            input.config,
            buildNumberingContext(variantParts.stylesXml, variantParts.numberingXml),
          );
          return {
            originalParas: templateParas,
            revisedParas: variant.paras,
            quarantinedOriginal: template.quarantinedCount,
            quarantinedRevised: variant.quarantinedCount,
            // A redline is roughly template-sized.
            artefactEstimate: input.template.byteLength,
            assemble: (xml: string) => assembleRedlinePackage(xml, templateBundle),
          };
        },
      })));

      return core.buildIndex({
        mode: 'one-to-many',
        templateLabel: input.template.label,
        flatten: { template: input.flattenTemplate, variant: input.flattenVariant },
      });
    } finally {
      // Reached even when the template itself is unreadable — otherwise every
      // failed template attempt would strand up to `pool.size` live workers for
      // the life of the page.
      if (ownsPool) pool.terminate();
    }
  };

  return {
    index: run(),
    cancel(): void {
      core.cancel();
      pool.cancelAll();
    },
  };
}

// ---------------------------------------------------------------------------
// Many-to-many (spec §6.7)
// ---------------------------------------------------------------------------

/** One pair's preprocessed input, built on demand. */
export interface PairRecords {
  originalParas: readonly ParaRecord[];
  revisedParas: readonly ParaRecord[];
  quarantinedOriginal: number;
  quarantinedRevised: number;
}

/**
 * One confirmed pair, ready to run.
 *
 * Both of the heavyweight things a pair needs are lazy callbacks rather than
 * held values, for one reason: holding either across a batch is the memory
 * failure §4.3 exists to prevent. The records are built by the lane running the
 * pair (see `records` below); the side-A package is read when that pair's own
 * artefact is assembled, and dropped immediately after.
 */
export interface ConfirmedPairInput {
  id: string;
  /** `"A ▸ B"`. */
  label: string;
  /** Stem for the artefact file name, without the " (redline)" suffix. */
  artefactStem: string;
  /**
   * Both sides' preprocessed records, produced when a lane picks this pair up.
   *
   * A thunk, not two arrays. Passing the arrays meant every confirmed pair's
   * records were live from the moment the batch was assembled — at the §4.4 cap
   * of fifty documents a side that is a hundred document-sized `ParaRecord[]`
   * at once, which spec §4.3 opens by naming as the out-of-memory crash its
   * whole policy exists to prevent. Deferred to the lane, the peak is one
   * pair's working set times the pool size, which is what §4.3 step 4 promises.
   */
  records(): Promise<PairRecords>;
  /** Bytes of the side-A document. Called at most once, at artefact time. */
  originalBytes(): Promise<ArrayBuffer>;
  originalByteLength: number;
  /**
   * Byte length of the side-B document.
   *
   * Its bytes are never read — the records came from the sketch — but its SIZE
   * is still a cap the pair has to satisfy. Without this the per-document cap
   * was enforced on one side of a two-sided comparison, which is not a cap.
   */
  revisedByteLength: number;
  /**
   * Set when the pair was identified but must not be compared — a cross-format
   * pair, or one over a cap. It still gets a row (acceptance M8).
   */
  unrunnableReason?: string;
}

export interface RunManyToManyInput {
  pairs: readonly ConfirmedPairInput[];
  config: CompareConfig;
  author: string;
  date: string;
  flattenA: FlattenPolicy;
  flattenB: FlattenPolicy;
  pool?: BatchPool;
  artefactBudgetBytes?: number;
  onRowChanged?: (row: BatchRow) => void;
  onProgress?: (progress: BatchProgress) => void;
}

/**
 * Run every confirmed pair of a many-to-many batch.
 *
 * Takes CONFIRMED pairs and nothing else. There is deliberately no path from
 * two bundles to a run inside this module: pairing is proposed and confirmed
 * upstream (`pairing/propose.ts`, the grid), and a runner that could pair for
 * itself would be a runner that could run an unconfirmed pairing.
 */
export function runManyToManyBatch(input: RunManyToManyInput): BatchRun {
  assertDistinctIds(input.pairs);
  // The BUNDLE cap, not the variant one. A pairing gives each side-A document
  // at most one counterpart, so the pair count is bounded by one side's
  // document count — and telling someone who handed over two folders that they
  // have too many "variants" names a concept their run does not contain.
  checkBundleCap(input.pairs.length, resolveBatchConfig(input.config).maxBundleDocuments);
  /**
   * An oversize document demotes ITS OWN pair, and nothing else.
   *
   * This used to call `checkDocxByteCap`, which throws — so one document over
   * the cap ended a confirmed run before a single pair had been dispatched, and
   * every other pair's redline was lost. That is right for a batch being
   * started (an oversize input is refused whole rather than quietly trimmed)
   * and wrong here: the pairing grid has ALREADY shown this pair with the cap
   * named, and the user confirmed a run knowing it would not be compared.
   * Throwing at that point turns an accepted, disclosed limitation into a
   * batch-wide failure — the many-to-many form of losing ninety-nine good
   * redlines to one bad file (spec §8 case 2, §6.6).
   *
   * The reason text comes from `limits.ts`, the same source the grid used, so
   * the sentence the user accepted is the sentence the row reports.
   */
  const pairs: readonly ConfirmedPairInput[] = input.pairs.map(pair => {
    if (pair.unrunnableReason !== undefined) return pair;
    // The label is `"A ▸ B"`, so the sentence names the PAIR rather than the
    // document. Reachable only by a direct caller — through the driver the grid
    // always blocks first, with the document's own name — but a message that
    // implies a file called "A ▸ B" would be its own small lie.
    const over = [
      { fileName: `the side-A document of ${pair.label}`, byteLength: pair.originalByteLength },
      { fileName: `the side-B document of ${pair.label}`, byteLength: pair.revisedByteLength },
    ].filter(s => exceedsDocxByteCap(s.byteLength));
    if (over.length === 0) return pair;
    return { ...pair, unrunnableReason: describeDocxCapBreach(over.slice(0, 1)) };
  });

  // The aggregate cap is measured AFTER the demotions, over the pairs that will
  // actually run. Counting a document the previous statement has just decided
  // not to compare would refuse a batch for bytes it was never going to hold —
  // and the aggregate cap's whole justification (spec §6.6) is what the run
  // holds at once.
  checkTotalByteCap(
    pairs
      .filter(p => p.unrunnableReason === undefined)
      .reduce((n, p) => n + p.originalByteLength + p.revisedByteLength, 0),
  );

  const pool = input.pool ?? createBatchPoolForHost({
    poolSize: resolveBatchConfig(input.config).poolSize,
  });
  const ownsPool = input.pool === undefined;

  const core = createBatchCore({
    pool,
    ids: pairs.map(p => ({ id: p.id, label: p.label })),
    config: input.config,
    author: input.author,
    date: input.date,
    ...(input.artefactBudgetBytes !== undefined
      ? { artefactBudgetBytes: input.artefactBudgetBytes }
      : {}),
    ...(input.onRowChanged !== undefined ? { onRowChanged: input.onRowChanged } : {}),
    ...(input.onProgress !== undefined ? { onProgress: input.onProgress } : {}),
  });

  const run = async (): Promise<BatchIndex> => {
    try {
      await core.runPairs(pairs.map(pair => ({
        id: pair.id,
        label: pair.label,
        artefactStem: pair.artefactStem,
        ...(pair.unrunnableReason !== undefined
          ? { unrunnableReason: pair.unrunnableReason }
          : {}),
        // Called by the lane that runs this pair, not up front: preparing every
        // pair eagerly would hold every document's records at once (see
        // `records` on `ConfirmedPairInput`). The side-A package is read inside
        // `assemble`, which runs only when the artefact will actually be kept.
        prepare: async (): Promise<PreparedPair> => ({
          ...await pair.records(),
          artefactEstimate: pair.originalByteLength,
          assemble: async (xml: string) =>
            assembleRedlinePackage(xml, new PizZipBundle(await pair.originalBytes())),
        }),
      })));

      return core.buildIndex({
        mode: 'many-to-many',
        flatten: { template: input.flattenA, variant: input.flattenB },
      });
    } finally {
      if (ownsPool) pool.terminate();
    }
  };

  return {
    index: run(),
    cancel(): void {
      core.cancel();
      pool.cancelAll();
    },
  };
}
