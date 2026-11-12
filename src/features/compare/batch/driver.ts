/**
 * Batch-compare driver (ADR-0037; spec §5.4, §5.5).
 *
 * Instantiable from the first line — there is deliberately no module-level
 * singleton here. The single-pair driver had one and it was the prerequisite
 * that blocked this whole feature (see `compare/driver.ts` and its
 * instance-isolation tests); repeating the mistake in the module whose entire
 * purpose is running many comparisons would be perverse.
 *
 * The driver holds `File` handles, never their bytes. A file's contents are
 * read by the runner when it is that pair's turn (spec §4.3 step 1), which is
 * what keeps a hundred-variant batch from costing a hundred `ArrayBuffer`s
 * before the first comparison starts.
 */
import { signal } from '@preact/signals';

import { getCompareConfig, getActiveProfile } from '@shared/config/index.js';
import type { CompareConfig, FlattenPolicy } from '@shared/config/index.js';
import type { CompareStatus, FlattenChoice } from '../ui-types.js';
import type { BatchIndex, BatchRow } from './model.js';
import type { BatchDriverState, BatchSlot, PairDetail } from './ui-types.js';
import { PizZipBundle } from '@shared/ooxml/bundle.js';
import { preprocessSide, buildNumberingContext } from '../preprocess/index.js';
import { extractParts } from '../io/parts.js';
import { runComparePipeline } from '../pipeline.js';
import { runOneToManyBatch, type BatchDocumentInput, type BatchRun } from './runner.js';
import {
  downloadArtefact,
  downloadRowArchive,
  isDownloadable,
  redlineOf,
} from './artefacts.js';
import { assembleRedlinePackage } from '../output/writeout.js';
import { resolveBatchConfig } from './config.js';
import type { ArtefactHandle } from './model.js';

const DOCX_MIME =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

let slotCounter = 0;
const nextSlotId = (): string => `slot-${(slotCounter += 1)}`;

function describeError(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

export interface CreateBatchDriverOptions {
  /**
   * Cumulative retained-artefact ceiling, passed through to the runner.
   * Exposed so the degradation path is reachable without a half-gigabyte
   * batch — a path that can only be exercised by an unreachable input is a
   * path nobody has actually run.
   */
  artefactBudgetBytes?: number;
}

/**
 * The settings one batch ran under, captured when it started.
 *
 * A re-run — the drill-down, or an on-demand artefact — must reproduce the row
 * it belongs to. Reading the live signals instead would let a redline be
 * assembled under a flatten choice the index explicitly denies (spec §8 case 7:
 * "a batch that silently flattened differently across variants would be
 * incomparable with itself"), and would let a drill-down's change list disagree
 * with the row's own `changeCount`.
 */
interface RunSettings {
  config: CompareConfig;
  flattenTemplate: FlattenPolicy;
  flattenVariant: FlattenPolicy;
  date: string;
  /**
   * The template the batch actually ran against.
   *
   * Pinned for the same reason the flatten choice is, and it is the easier one
   * to lose: choosing a different template after a batch finishes would make
   * every later re-run compare against the NEW template while the rows still
   * carry the old one's counts. The variant files are pinned by
   * `removeVariant`'s refusal; without this the template was not.
   */
  templateFile: File;
}

export function createBatchDriver(options: CreateBatchDriverOptions = {}): BatchDriverState {
  const status = signal<CompareStatus>('idle');
  const progress = signal(0);
  const errorMessage = signal<string | null>(null);
  const template = signal<BatchSlot | null>(null);
  const variants = signal<readonly BatchSlot[]>([]);
  const rows = signal<readonly BatchRow[]>([]);
  const index = signal<BatchIndex | null>(null);
  const flattenTemplate = signal<FlattenChoice>('acceptAll');
  const flattenVariant = signal<FlattenChoice>('acceptAll');

  /** Files by slot id. Not a signal: heavy, non-reactive, and read once each. */
  /** The cap the runner will enforce, so the panel cannot show a different one. */
  const maxVariants = resolveBatchConfig(getCompareConfig(getActiveProfile())).maxVariants;

  const files = new Map<string, File>();
  let templateFile: File | null = null;
  let run: BatchRun | null = null;
  let runSettings: RunSettings | null = null;

  const toInput = (slot: BatchSlot, file: File): BatchDocumentInput => ({
    id: slot.id,
    label: slot.label,
    byteLength: slot.byteLength,
    bytes: () => file.arrayBuffer(),
  });

  const upsertRow = (row: BatchRow): void => {
    const current = rows.value;
    const at = current.findIndex(r => r.id === row.id);
    rows.value = at === -1
      ? [...current, row]
      : [...current.slice(0, at), row, ...current.slice(at + 1)];
  };

  const findRow = (id: string): BatchRow | undefined =>
    (index.value?.rows ?? rows.value).find(r => r.id === id);

  /**
   * Re-run one completed pair from the files still held here.
   *
   * The pipeline ONLY — never `materialise`, which writes back through the host
   * and on web means triggering a download. "Show me what changed" must not
   * save a file. (Spec §5.4, amended: the drill-down is a pipeline re-run of the
   * single pair, not a second single-pair driver, for exactly this reason.)
   *
   * Also the source of on-demand artefacts: past the retained-artefact ceiling
   * a row keeps no `Blob`, and this is how its redline is still obtainable.
   */
  const comparePair = async (
    id: string,
  ): Promise<{ result: Awaited<ReturnType<typeof runComparePipeline>>; templateBundle: PizZipBundle } | null> => {
    const row = findRow(id);
    const file = files.get(id);
    if (!row || row.status !== 'done' || !file) return null;

    // The batch's own settings and the batch's own template, not whatever the
    // signals and the file input hold now.
    const live = templateFile;
    if (runSettings === null && live === null) return null;
    const settings: RunSettings = runSettings ?? {
      config: getCompareConfig(getActiveProfile()),
      flattenTemplate: flattenTemplate.value,
      flattenVariant: flattenVariant.value,
      date: index.value?.generatedAt ?? new Date().toISOString(),
      templateFile: live as File,
    };

    const templateBundle = new PizZipBundle(await settings.templateFile.arrayBuffer());
    const side = async (
      bundle: PizZipBundle,
      policy: FlattenPolicy,
    ): Promise<ReturnType<typeof preprocessSide>> => {
      const parts = extractParts(bundle);
      // ADR-0042 — same context as the single-pair path, so a batch redline
      // does not silently omit a change the panel would report.
      return preprocessSide(
        parts.documentXml,
        'document',
        policy,
        settings.config,
        buildNumberingContext(parts.stylesXml, parts.numberingXml),
      );
    };

    const original = side(templateBundle, settings.flattenTemplate);
    const revised = side(new PizZipBundle(await file.arrayBuffer()), settings.flattenVariant);

    const result = await runComparePipeline({
      originalParas: (await original).paras,
      revisedParas: (await revised).paras,
      config: settings.config,
      author: settings.config.author,
      date: settings.date,
      quarantinedOriginal: (await original).quarantinedCount,
      quarantinedRevised: (await revised).quarantinedCount,
    });

    return { result, templateBundle };
  };

  /** Produce a row's redline on demand, by re-running that one pair. */
  const produceArtefact = async (row: BatchRow): Promise<ArtefactHandle | null> => {
    const rerun = await comparePair(row.id);
    if (!rerun) return null;
    const pkg = assembleRedlinePackage(rerun.result.authoredBodyXml, rerun.templateBundle);
    const bytes = new Uint8Array(pkg.toUint8Array());
    return {
      kind: 'redlineDocx',
      fileName: `${row.label.replace(/\.docx$/i, '')} (redline).docx`,
      byteLength: bytes.byteLength,
      blob: new Blob([bytes], { type: DOCX_MIME }),
    };
  };

  return {
    status,
    progress,
    errorMessage,
    template,
    variants,
    rows,
    index,
    flattenTemplate,
    flattenVariant,
    maxVariants,

    setTemplate(file: File): void {
      templateFile = file;
      template.value = { id: 'template', label: file.name, byteLength: file.size };
    },

    addVariants(incoming: readonly File[]): void {
      const added: BatchSlot[] = [];
      for (const file of incoming) {
        const slot: BatchSlot = { id: nextSlotId(), label: file.name, byteLength: file.size };
        files.set(slot.id, file);
        added.push(slot);
      }
      variants.value = [...variants.value, ...added];
    },

    removeVariant(id: string): void {
      // A row that has already been produced still needs its file: the
      // drill-down and any on-demand artefact are re-runs, not replays. Dropping
      // the file would leave a row whose "Open changes" and "Download redline"
      // buttons silently do nothing — an offered action that yields no result
      // and no message, which is worse than not offering it.
      if (rows.value.some(r => r.id === id)) {
        errorMessage.value =
          'That version has already been compared. Clear the batch to start a new one.';
        return;
      }
      files.delete(id);
      variants.value = variants.value.filter(v => v.id !== id);
    },

    clearVariants(): void {
      files.clear();
      variants.value = [];
      rows.value = [];
      index.value = null;
      runSettings = null;
      errorMessage.value = null;
      status.value = 'idle';
      progress.value = 0;
    },

    async runBatch(): Promise<void> {
      if (status.value === 'running') return;

      const templateSlot = template.value;
      if (templateSlot === null || templateFile === null) {
        errorMessage.value = 'Choose a template document before running the batch.';
        status.value = 'error';
        return;
      }
      if (variants.value.length === 0) {
        errorMessage.value = 'Add at least one variant to compare against the template.';
        status.value = 'error';
        return;
      }

      status.value = 'running';
      progress.value = 0;
      errorMessage.value = null;
      rows.value = [];
      index.value = null;

      try {
        const profile = getActiveProfile();
        const config = getCompareConfig(profile);
        runSettings = {
          config,
          flattenTemplate: flattenTemplate.value,
          flattenVariant: flattenVariant.value,
          date: new Date().toISOString(),
          templateFile,
        };

        run = runOneToManyBatch({
          template: toInput(templateSlot, templateFile),
          variants: variants.value.map(slot => toInput(slot, files.get(slot.id) as File)),
          config,
          author: config.author,
          date: runSettings.date,
          flattenTemplate: runSettings.flattenTemplate,
          flattenVariant: runSettings.flattenVariant,
          ...(options.artefactBudgetBytes !== undefined
            ? { artefactBudgetBytes: options.artefactBudgetBytes }
            : {}),
          onRowChanged: upsertRow,
          onProgress: p => { progress.value = p.pct; },
        });

        const finished = await run.index;
        index.value = finished;
        rows.value = finished.rows;
        // A cancelled batch is not a finished one. Driving the bar to 100 %
        // over a run that stopped at 60 of 100 contradicts the rows' own
        // Cancelled chips, and that contradiction resolves in a reader's head
        // as "it completed" — the wrong half. The bar shows how far it actually
        // got; the panel says so in words (spec §5.5). `status` stays within
        // the shared `CompareStatus` union — the run did end — and cancellation
        // is read off the totals, which cannot disagree with the rows.
        const stopped = finished.totals.cancelled > 0;
        progress.value = stopped
          ? Math.round((finished.totals.done / Math.max(finished.totals.pairs, 1)) * 100)
          : 100;
        status.value = 'done';
      } catch (err: unknown) {
        // A per-pair failure never reaches here — the runner turns it into a
        // failed row. This is a whole-batch refusal: a cap, or a template that
        // is not a readable package.
        errorMessage.value = describeError(err);
        status.value = 'error';
      } finally {
        run = null;
      }
    },

    cancelBatch(): void {
      run?.cancel();
    },

    async openPair(id: string): Promise<PairDetail | null> {
      const row = findRow(id);
      const rerun = await comparePair(id);
      if (!row || !rerun) return null;
      return { rowId: id, label: row.label, ops: rerun.result.ops, summary: rerun.result.summary };
    },

    async downloadRow(id: string): Promise<void> {
      const row = findRow(id);
      if (!row) return;
      if (!isDownloadable(row)) {
        // A row that failed, was cancelled or has not run yet has no redline
        // to produce, and "that redline could not be produced" would
        // misdescribe it: the packaging did not fail, the comparison never
        // completed. The row's own status and reason say which.
        errorMessage.value =
          `There is no redline for ${row.label} — that comparison did not complete.`;
        return;
      }
      try {
        // `downloadArtefact` returns quietly when there is nothing to hand
        // over, so the producer is called here to find out whether there was —
        // otherwise an offered button could do nothing and say nothing, which
        // is the failure this whole path was reopened for.
        const produced = redlineOf(row) ?? (await produceArtefact(row));
        if (produced === null) {
          errorMessage.value =
            'That redline could not be produced. Run the batch again to recover it.';
          return;
        }
        await downloadArtefact({ ...row, artefacts: [produced] });
        // A download that worked must not leave the previous one's failure on
        // screen beside it. Nothing else clears this until the batch is.
        errorMessage.value = null;
      } catch (err: unknown) {
        errorMessage.value = `That redline could not be produced: ${describeError(err)}`;
      }
    },

    async downloadAll(): Promise<void> {
      const all = index.value?.rows ?? rows.value;
      await downloadRowArchive(all, template.value?.label ?? 'batch', produceArtefact);
    },
  };
}
