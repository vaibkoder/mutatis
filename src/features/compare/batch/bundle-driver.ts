/**
 * The many-to-many (bundle-against-bundle) driver (ADR-0037; spec §6).
 *
 * Sits between the pairing grid and the 17a batch runner, and its whole job is
 * to keep those two apart until a human has said so. `runConfirmed()` refuses
 * while `confirmed` is false — not because the UI would not offer the button,
 * but because the UI is not where an invariant belongs. A mis-pairing produces
 * a complete, well-formed, entirely worthless redline that a reader has no cue
 * to distrust, so the gate is enforced where it cannot be routed around.
 *
 * A separate driver from `createBatchDriver` rather than an extension of it.
 * The two share the runner, the pool, the caps, the index model and the zip
 * writer — everything that is genuinely common — but their INPUT states have
 * nothing in common: one template and N variants versus two bundles, a score
 * matrix, a proposal and a confirmation. Folding both into one object would
 * give every field a "…but only in the other mode" caveat, which is how a gate
 * like `confirmed` ends up quietly inapplicable on one path.
 *
 * Instantiable, like every driver written since ADR-0037's prerequisite
 * refactor. No module-level singleton.
 */
import { computed, signal } from '@preact/signals';
import type { Signal } from '@preact/signals';

import { getCompareConfig, getActiveProfile } from '@shared/config/index.js';
import type { CompareConfig, FlattenPolicy, JurisdictionProfile } from '@shared/config/index.js';
import type { CompareStatus } from '../ui-types.js';
import type { BatchIndex, BatchRow } from './model.js';
import type { BatchSlot, PairDetail } from './ui-types.js';
import { checkBundleCap, exceedsDocxByteCap } from './limits.js';
import { resolveBatchConfig, resolvePairingSettings } from './config.js';
import {
  runManyToManyBatch,
  type BatchRun,
  type ConfirmedPairInput,
  type PairRecords,
} from './runner.js';
import {
  buildPairingSketch,
  filenameOnlySketch,
  preprocessBundleDocument,
  sourceKindOf,
  unreadableSketch,
  type BundleSide,
  type DocumentRecords,
  type PairingSketch,
} from './pairing/sketch.js';
import { proposePairing, repointPairing } from './pairing/propose.js';
import type { PairingProposal } from './pairing/model.js';
import { scorePair, rankCandidates, type PairScore } from './pairing/score.js';
import { PizZipBundle } from '@shared/ooxml/bundle.js';
import { runComparePipeline } from '../pipeline.js';
import { assembleRedlinePackage } from '../output/writeout.js';
import {
  downloadArtefact,
  downloadRowArchive,
  isDownloadable,
  redlineOf,
} from './artefacts.js';
import type { ArtefactHandle } from './model.js';

const DOCX_MIME =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
import type { PairingGridState } from './pairing/ui-types.js';

export interface BundleDriverState extends PairingGridState {
  status: Signal<CompareStatus>;
  progress: Signal<number>;
  errorMessage: Signal<string | null>;
  sideA: Signal<readonly BatchSlot[]>;
  sideB: Signal<readonly BatchSlot[]>;
  rows: Signal<readonly BatchRow[]>;
  index: Signal<BatchIndex | null>;
  flattenA: Signal<FlattenPolicy>;
  flattenB: Signal<FlattenPolicy>;
  /** The per-side document cap this driver will enforce. */
  maxBundleDocuments: number;

  addDocuments(side: BundleSide, files: readonly File[]): void;
  removeDocument(side: BundleSide, id: string): void;
  clearBundles(): void;

  /** Read both bundles, sketch every document, and propose a pairing. */
  proposePairs(): Promise<void>;

  /** Run every confirmed, runnable pair. Refuses while unconfirmed. */
  runConfirmed(): Promise<void>;
  cancelRun(): void;

  /**
   * Open one pair's full change list by re-running that single pair
   * (spec §2.1 capability 8).
   *
   * The batch drops every `PipelineResult` the moment its pair completes — two
   * document-sized record arrays per pair is what breaks first at N — so the
   * ops are not kept, they are recovered. A pipeline run only: packaging a
   * redline here would trigger a download just to answer "what changed?".
   */
  openPair(id: string): Promise<PairDetail | null>;

  /** Download one pair's redline, re-running it if it was not retained. */
  downloadRow(id: string): Promise<void>;
  /** Download every obtainable redline as one archive. */
  downloadAll(): Promise<void>;
}

export interface CreateBundleDriverOptions {
  artefactBudgetBytes?: number;
}

let slotCounter = 0;
const nextSlotId = (side: BundleSide): string => `${side}-${(slotCounter += 1)}`;

function describeError(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

/**
 * The settings one batch ran under, captured when it started.
 *
 * A re-run — the drill-down, or an on-demand redline — must reproduce the row
 * it belongs to. The flatten signals are public and writable, and `index.value`
 * is null for the whole duration of a run, so reading either live would let a
 * redline be assembled under a policy the index denies (spec §8 case 7) or
 * stamped with a different date from every other redline in its own batch.
 * The one-to-many driver pins the same values for the same reason.
 */
interface RunSettings {
  flattenA: FlattenPolicy;
  flattenB: FlattenPolicy;
  date: string;
}

export function createBundleDriver(
  options: CreateBundleDriverOptions = {},
): BundleDriverState {
  const status = signal<CompareStatus>('idle');
  const progress = signal(0);
  const errorMessage = signal<string | null>(null);
  const sideA = signal<readonly BatchSlot[]>([]);
  const sideB = signal<readonly BatchSlot[]>([]);
  const rows = signal<readonly BatchRow[]>([]);
  const index = signal<BatchIndex | null>(null);
  const proposal = signal<PairingProposal | null>(null);
  const confirmed = signal(false);
  const busy = signal(false);
  /**
   * Derived, never assigned: the grid must not be able to disagree with the
   * runner about whether a run is in flight, and a second signal kept in step
   * by hand is exactly how it would.
   */
  const running = computed(() => status.value === 'running');
  const flattenA = signal<FlattenPolicy>('acceptAll');
  const flattenB = signal<FlattenPolicy>('acceptAll');

  const profile: JurisdictionProfile = getActiveProfile();
  const config: CompareConfig = getCompareConfig(profile);
  const maxBundleDocuments = resolveBatchConfig(config).maxBundleDocuments;

  /**
   * The settings the current index was produced under, captured when the run
   * started. `null` until one has. See `readRecords` for why the flatten policy
   * is here; `date` is here for the same reason, and because `index.value` is
   * null for the whole duration of a run — so a redline downloaded mid-batch
   * would otherwise carry a different date from every other one in it.
   */
  let runSettings: RunSettings | null = null;

  /** Files and sketches by slot id. Not signals: heavy and non-reactive. */
  const files = new Map<string, File>();
  const sketches = new Map<string, PairingSketch>();
  /** The flatten policy each sketch was built under, so a change can rebuild it. */
  const sketchFlatten = new Map<string, FlattenPolicy>();
  let run: BatchRun | null = null;

  const sketchesFor = (side: BundleSide): PairingSketch[] =>
    (side === 'a' ? sideA.value : sideB.value)
      .map(s => sketches.get(s.id))
      .filter((s): s is PairingSketch => s !== undefined);

  const proposeInput = () => ({
    sideA: sketchesFor('a'),
    sideB: sketchesFor('b'),
    generatedAt: proposal.value?.generatedAt ?? new Date().toISOString(),
    config,
  });

  /**
   * Discard a proposal whenever the documents behind it change.
   *
   * A confirmation is a statement about a specific set of documents. Carrying
   * it across an edit would let a run proceed on a pairing the user confirmed
   * for a different bundle — a confirmation gate that can go stale is not one.
   */
  /**
   * Refuse an edit that would change the pairing a run is executing.
   *
   * Every one of these paths ends in `invalidate()` → `discardResults()`, which
   * sets `status` back to `'idle'`. Mid-run that does not just clear the screen:
   * it makes `running` — which is derived from `status` — say false while a run
   * is still writing rows, which in turn disarms `runConfirmed`'s own re-entry
   * guard and lets a SECOND run start while the first is still going, both
   * interleaving into the same `rows`. A review reached exactly that by calling
   * `addDocuments` from a `rows` subscriber. The run's continuation then writes
   * the old pairing's rows and index back over the cleared state, leaving a
   * complete, downloadable set of redlines under a grid saying "add documents
   * to both sides".
   *
   * So the rule is the one ADR-0037 states: a run consumes the pairing it was
   * given, and the way out of a run is to cancel it.
   */
  const refuseWhileWorking = (): boolean => {
    if (running.value) {
      errorMessage.value =
        'Comparisons are running against this pairing. Cancel the run before changing it.';
      return true;
    }
    // `busy` is the sketch-building window. It is short, but `proposePairs`
    // resolves its sketches against the CURRENT bundle lists, so a document
    // added while it runs would be absent from the grid it produces — and "every
    // document appears exactly once" (§6.6) is the one property the grid cannot
    // be wrong about.
    if (busy.value) {
      errorMessage.value = 'The bundles are still being read. Wait for the pairing to be proposed.';
      return true;
    }
    return false;
  };

  const invalidate = (): void => {
    proposal.value = null;
    confirmed.value = false;
    discardResults();
  };

  /**
   * Drop the results of a previous run.
   *
   * Rows and an index describe one specific pairing. When that pairing changes
   * — a document added or removed, a row re-pointed — the rows below the grid
   * are answers to a question nobody is asking any more.
   */
  function discardResults(): void {
    rows.value = [];
    index.value = null;
    // The settings belong to the results. Kept past them, a later re-run would
    // reproduce a row that no longer exists.
    runSettings = null;
    status.value = 'idle';
    progress.value = 0;
  }

  /**
   * Both sides' records for one pair, built now and released when the caller
   * lets go of them.
   *
   * Sketches carry hashes, not records (`pairing/sketch.ts`), so this is the
   * only place a bundle document's `ParaRecord[]` comes from — and it is
   * reached once per pair, from the lane running it.
   *
   * `settings` is passed rather than read: a re-run must reproduce the row it
   * belongs to, and the flatten signals are public and writable, so reading
   * them here would let a redline be assembled under a policy the index
   * explicitly denies (§8 case 7) and a drill-down's change list disagree with
   * its own row's `changeCount`. Taking it as a parameter means there is no
   * fallback to a live value for a caller to reach by accident. That is what keeps the
   * peak at one pair's working set times the pool size rather than at the whole
   * bundle's (spec §4.3).
   */
  const readRecords = async (
    sk: PairingSketch,
    settings: RunSettings,
  ): Promise<DocumentRecords> => {
    const file = files.get(sk.docId) as File;
    const flatten = sk.side === 'a' ? settings.flattenA : settings.flattenB;
    return preprocessBundleDocument(await file.arrayBuffer(), config, flatten);
  };

  const recordsFor = async (
    a: PairingSketch,
    b: PairingSketch,
    settings: RunSettings,
  ): Promise<PairRecords> => {
    const original = await readRecords(a, settings);
    const revised = await readRecords(b, settings);
    return {
      originalParas: original.paras,
      revisedParas: revised.paras,
      quarantinedOriginal: original.quarantinedCount,
      quarantinedRevised: revised.quarantinedCount,
    };
  };

  /**
   * Re-run one confirmed pair's PIPELINE, and nothing else.
   *
   * Deliberately separate from packaging: the drill-down needs the ops and the
   * summary, and must not touch `assembleRedlinePackage` — on web, materialising
   * a redline means triggering a download, and "show me what changed" must not
   * save a file. The same wall the one-to-many driver keeps (spec §5.4).
   */
  const comparePair = async (
    rowId: string,
  ): Promise<{
    result: Awaited<ReturnType<typeof runComparePipeline>>;
    file: File;
    label: string;
  } | null> => {
    const row = proposal.value?.rows.find(r => r.id === rowId);
    // The gate. `runConfirmed` is not the only public method that starts a
    // comparison — `openPair` is the other, and it was outside it, so a caller
    // holding a proposal could get a complete, plausible change list for a
    // pairing no human had read. That artefact is the whole thing ADR-0008
    // applied to pairing exists to prevent (spec §6.1: "nothing is compared
    // before the user confirms the pairing").
    //
    // The condition is `runSettings`, not `confirmed`: a re-run only means
    // anything for a row a confirmed run actually produced, and `runSettings`
    // is set only by `runConfirmed` and cleared with the results. Reading
    // `confirmed` instead would also refuse a download of an on-demand redline
    // after the user re-opened the grid — results that are still on screen and
    // still correct.
    if (runSettings === null) return null;
    const settings = runSettings;
    if (!row || row.a === null || row.b === null || !row.runnable) return null;
    const a = sketches.get(row.a);
    const b = sketches.get(row.b);
    const file = a === undefined ? undefined : files.get(a.docId);
    if (a === undefined || b === undefined || file === undefined) return null;

    const original = await readRecords(a, settings);
    const revised = await readRecords(b, settings);
    const result = await runComparePipeline({
      originalParas: original.paras,
      revisedParas: revised.paras,
      quarantinedOriginal: original.quarantinedCount,
      quarantinedRevised: revised.quarantinedCount,
      config,
      author: config.author,
      date: settings.date,
    });

    return { result, file, label: row.label };
  };

  /**
   * Re-run one confirmed pair and package its redline.
   *
   * One pipeline run over records built for this call, plus one package read.
   * It is what makes the artefact ceiling a disclosed degradation rather than a
   * hundred results the user cannot reach (spec §4.3 step 6).
   */
  const produceArtefact = async (rowId: string): Promise<ArtefactHandle | null> => {
    const rerun = await comparePair(rowId);
    if (rerun === null) return null;
    const { result, file } = rerun;

    const pkg = assembleRedlinePackage(
      result.authoredBodyXml,
      new PizZipBundle(await file.arrayBuffer()),
    );
    const bytes = new Uint8Array(pkg.toUint8Array());
    return {
      kind: 'redlineDocx',
      fileName: `${file.name.replace(/\.docx$/i, '')} (redline).docx`,
      byteLength: bytes.byteLength,
      blob: new Blob([bytes], { type: DOCX_MIME }),
    };
  };

  return {
    status,
    progress,
    errorMessage,
    sideA,
    sideB,
    rows,
    index,
    proposal,
    confirmed,
    busy,
    running,
    flattenA,
    flattenB,
    maxBundleDocuments,

    addDocuments(side: BundleSide, incoming: readonly File[]): void {
      if (refuseWhileWorking()) return;
      const target = side === 'a' ? sideA : sideB;
      const added: BatchSlot[] = [];
      for (const file of incoming) {
        const slot: BatchSlot = { id: nextSlotId(side), label: file.name, byteLength: file.size };
        files.set(slot.id, file);
        added.push(slot);
      }
      target.value = [...target.value, ...added];
      invalidate();
    },

    removeDocument(side: BundleSide, id: string): void {
      if (refuseWhileWorking()) return;
      const target = side === 'a' ? sideA : sideB;
      files.delete(id);
      sketches.delete(id);
      sketchFlatten.delete(id);
      target.value = target.value.filter(s => s.id !== id);
      invalidate();
    },

    clearBundles(): void {
      if (refuseWhileWorking()) return;
      files.clear();
      sketches.clear();
      sketchFlatten.clear();
      sideA.value = [];
      sideB.value = [];
      rows.value = [];
      index.value = null;
      errorMessage.value = null;
      status.value = 'idle';
      progress.value = 0;
      invalidate();
    },

    async proposePairs(): Promise<void> {
      if (refuseWhileWorking()) return;
      errorMessage.value = null;
      // A previous run's rows and index describe the pairing that produced
      // them. A fresh proposal is a new question, so they go — the same rule
      // every other path into `invalidate()` follows, and the one reason a
      // result may ever sit under a grid is that the grid is the one it came
      // from.
      discardResults();

      if (sideA.value.length === 0 || sideB.value.length === 0) {
        errorMessage.value = 'Add documents to both sides before proposing a pairing.';
        return;
      }

      busy.value = true;
      try {
        checkBundleCap(sideA.value.length, maxBundleDocuments);
        checkBundleCap(sideB.value.length, maxBundleDocuments);

        // One document at a time. The sketch keeps only the hashes, but
        // building it needs the whole document's records for the length of the
        // call — reading a whole bundle at once would peak at every document's
        // bytes AND every document's records simultaneously.
        for (const side of ['a', 'b'] as const) {
          const flatten = side === 'a' ? flattenA.value : flattenB.value;
          for (const slot of side === 'a' ? sideA.value : sideB.value) {
            // Reuse a sketch only if it was built under the flatten policy in
            // force now. Skipping on presence alone let the two drift: a
            // flatten change after a proposal would leave records preprocessed
            // under the old policy while the index stated the new one — an
            // index claiming a policy its redlines were not produced under,
            // which is the exact thing spec §8 case 7 puts flatten on the index
            // to prevent.
            if (sketches.has(slot.id) && sketchFlatten.get(slot.id) === flatten) continue;
            const file = files.get(slot.id) as File;
            sketchFlatten.set(slot.id, flatten);
            const identity = {
              docId: slot.id,
              side,
              fileName: slot.label,
              byteLength: file.size,
            };

            // Two documents whose bytes are never read, for the same reason:
            // neither is going to be compared, so spending every one of them —
            // and then holding the records for the length of the call — is work
            // spent on a comparison already ruled out. Spec §4.4 asks the guard
            // to fail before any partial work; §4.3 is the peak it protects.
            // Over the per-document cap: shown with the cap named (§6.6).
            // Not a Word document: shown with its own reason (§2.1, M8).
            if (exceedsDocxByteCap(file.size) || sourceKindOf(slot.label) !== 'docx') {
              sketches.set(slot.id, filenameOnlySketch(identity));
              continue;
            }

            // Per document, not per bundle — and the READ is inside it, not
            // only the parse. A file named `.docx` that is not a Word package
            // throws out of the OOXML reader; a file that has moved, been
            // truncated or lost permission since the picker ran rejects out of
            // `arrayBuffer()`. One `try` around the whole loop turned either
            // into a null proposal: no grid at all, and every other document in
            // both bundles gone with it. Spec §8 case 2 — one bad file must not
            // lose ninety-nine good redlines — and the failure boundary is one
            // document, which means everything done to that document.
            try {
              sketches.set(slot.id, buildPairingSketch({
                docId: slot.id,
                side,
                fileName: slot.label,
                bytes: await file.arrayBuffer(),
                config,
                profile,
                flatten,
              }));
            } catch (err: unknown) {
              sketches.set(slot.id, unreadableSketch(
                identity,
                `${slot.label} could not be read, so it was left out of the pairing. `
                  + `(${describeError(err)})`,
              ));
            }
          }
        }

        proposal.value = proposePairing({
          sideA: sketchesFor('a'),
          sideB: sketchesFor('b'),
          generatedAt: new Date().toISOString(),
          config,
        });
        confirmed.value = false;
      } catch (err: unknown) {
        errorMessage.value = describeError(err);
        proposal.value = null;
      } finally {
        busy.value = false;
      }
    },

    candidatesFor(rowId: string): readonly PairScore[] {
      const row = proposal.value?.rows.find(r => r.id === rowId);
      if (!row || row.a === null) return [];
      const a = sketches.get(row.a);
      // An unreadable document has nothing to score with and cannot be paired
      // (`pairing/propose.ts` keeps it out of the assignment), so it offers no
      // candidates and is offered as none.
      if (a === undefined || a.unreadable !== undefined) return [];
      const settings = resolvePairingSettings(config);
      return rankCandidates(
        sketchesFor('b')
          .filter(b => b.unreadable === undefined)
          .map(b => scorePair(a, b, settings)),
        settings,
      );
    },

    repoint(rowId: string, b: string | null): void {
      const current = proposal.value;
      // `confirmed` is the ordinary guard; `running` is the one a mid-run
      // `unconfirm()` would otherwise walk straight past.
      if (current === null || confirmed.value) return;
      if (refuseWhileWorking()) return;
      // Neither end of a re-point may be a document that could not be read.
      // The grid does not offer them, but `repoint` is public API, and the
      // failure is silent: the assignment never sees such a document, so the
      // choice is discarded — after the row holding the real counterpart has
      // already been rebuilt away, leaving a correct pairing destroyed and a
      // reason ("no counterpart in the other bundle") that is untrue.
      const row = current.rows.find(r => r.id === rowId);
      if (row?.a !== null && row?.a !== undefined && sketches.get(row.a)?.unreadable !== undefined) {
        return;
      }
      if (b !== null && sketches.get(b)?.unreadable !== undefined) return;
      // Asking for the pairing a row already has is not a change. Identity on
      // the rebuilt proposal does not catch it — `repointPairing` rebuilds
      // regardless — and the cost of missing it is a batch's redlines.
      if (row !== undefined && (row.b ?? null) === b) return;

      const next = repointPairing(proposeInput(), current, rowId, b);
      proposal.value = next;
      // Results from a previous run described a DIFFERENT pairing. Left on
      // screen they would sit under a grid that no longer agrees with them,
      // and a reader has no way to tell which of the two is current — so the
      // row list goes when the pairing it described does.
      //
      // Only when it actually does: `repointPairing` returns the proposal
      // untouched when there is nothing to re-point, and throwing away a
      // batch's redlines for a request that changed nothing is a loss with no
      // corresponding gain.
      if (next !== current) discardResults();
      // Re-pointing after confirming is impossible (guarded above); re-pointing
      // before it leaves the gate exactly where it was — shut.
    },

    confirm(): void {
      if (proposal.value !== null) confirmed.value = true;
    },

    unconfirm(): void {
      // A run consumes the pairing it was given. Re-opening the grid while one
      // is in flight leaves finished rows and a written index describing a
      // pairing the grid above them no longer shows — and `runConfirmed`'s own
      // continuation then writes the OLD pairing back over the cleared state.
      // Cancel first; the way out of a run is the cancel button.
      if (refuseWhileWorking()) return;
      confirmed.value = false;
    },

    async runConfirmed(): Promise<void> {
      // Re-entry first, and BEFORE the flatten-drift check below — that check
      // calls `invalidate()`, which resets `status`, which is what `running` is
      // derived from. Placed after it, a second `runConfirmed()` during a run
      // cleared the proposal, made `running` say false while the first run was
      // still writing rows, and re-opened every editing path the helper guards.
      // Every writer of `status` has to be behind the same door.
      //
      // `busy` too. During the sketch-building window `confirmed` is still true
      // and `proposal` still holds the PREVIOUS one, so a run started then would
      // be un-confirmed mid-flight by the propose that follows it — a complete,
      // downloadable batch under a grid that says nothing was confirmed. Today
      // the window is not observable (a re-propose that reuses every sketch
      // never awaits), which is an implementation detail of the reuse check
      // rather than a property, so the guard is here rather than in a comment.
      if (running.value || busy.value) return;

      const current = proposal.value;
      if (current === null) return;

      // The gate, enforced here rather than only in the UI. A disabled button
      // is a courtesy; this is the invariant.
      if (!confirmed.value) {
        errorMessage.value = 'Confirm the pairing before running any comparisons.';
        return;
      }

      // A confirmation is a statement about specific documents UNDER a specific
      // flatten policy. `proposePairs` rebuilds a sketch whose policy changed,
      // but nothing stopped the policy changing after the proposal — which
      // would run records preprocessed under the old one while the index stated
      // the new one, the exact contradiction §8 case 7 puts flatten on the
      // index to prevent. Treated like any other change to the inputs: the
      // proposal goes, and the user proposes again.
      const drifted = [...sketches.keys()].some(id => sketchFlatten.get(id)
        !== (sideA.value.some(slot => slot.id === id) ? flattenA.value : flattenB.value));
      if (drifted) {
        invalidate();
        errorMessage.value =
          'The tracked-changes setting changed after this pairing was proposed. '
          + 'Propose the pairing again so the comparisons match it.';
        return;
      }

      status.value = 'running';
      progress.value = 0;
      errorMessage.value = null;
      rows.value = [];
      index.value = null;

      try {
        // Captured before the pairs are built, so each pair's `records()` thunk
        // closes over the settings THIS run is using rather than reading them
        // again when the lane calls it.
        const settings: RunSettings = {
          flattenA: flattenA.value,
          flattenB: flattenB.value,
          date: new Date().toISOString(),
        };
        runSettings = settings;

        const pairs: ConfirmedPairInput[] = current.rows
          .filter(r => r.a !== null && r.b !== null)
          .map(r => {
            const a = sketches.get(r.a as string) as PairingSketch;
            const b = sketches.get(r.b as string) as PairingSketch;
            const file = files.get(a.docId) as File;
            return {
              id: r.id,
              label: r.label,
              artefactStem: a.fileName.replace(/\.docx$/i, ''),
              records: () => recordsFor(a, b, settings),
              originalBytes: () => file.arrayBuffer(),
              originalByteLength: a.byteLength,
              revisedByteLength: b.byteLength,
              ...(r.runnable ? {} : { unrunnableReason: r.reason as string }),
            };
          });

        run = runManyToManyBatch({
          pairs,
          config,
          author: config.author,
          date: settings.date,
          flattenA: settings.flattenA,
          flattenB: settings.flattenB,
          ...(options.artefactBudgetBytes !== undefined
            ? { artefactBudgetBytes: options.artefactBudgetBytes }
            : {}),
          onRowChanged: row => {
            const at = rows.value.findIndex(r => r.id === row.id);
            rows.value = at === -1
              ? [...rows.value, row]
              : [...rows.value.slice(0, at), row, ...rows.value.slice(at + 1)];
          },
          onProgress: p => { progress.value = p.pct; },
        });

        const finished = await run.index;
        index.value = finished;
        rows.value = finished.rows;
        // A cancelled batch shows how far it actually got (spec §5.5).
        progress.value = finished.totals.cancelled > 0
          ? Math.round((finished.totals.done / Math.max(finished.totals.pairs, 1)) * 100)
          : 100;
        status.value = 'done';
      } catch (err: unknown) {
        errorMessage.value = describeError(err);
        status.value = 'error';
      } finally {
        run = null;
      }
    },

    cancelRun(): void {
      run?.cancel();
    },

    async openPair(id: string): Promise<PairDetail | null> {
      // The pipeline only — never `produceArtefact`, which packages a redline
      // and on web means triggering a download. Showing what changed must not
      // save a file (spec §5.4).
      const rerun = await comparePair(id);
      if (rerun === null) return null;
      return {
        rowId: id,
        label: rerun.label,
        ops: rerun.result.ops,
        summary: rerun.result.summary,
      };
    },

    async downloadRow(id: string): Promise<void> {
      const row = (index.value?.rows ?? rows.value).find(r => r.id === id);
      if (!row) return;
      if (!isDownloadable(row)) {
        // Without this, an id for a row the GRID declared unrunnable would be
        // re-run here: `produceArtefact` reads the sketches, and a cross-format
        // pair's side-B sketch has no records, so it would package a
        // delete-everything "redline" for a comparison ADR-0017 gates. The
        // panel never offers the button, but `downloadRow` is public API and
        // an exported hazard is still a hazard.
        errorMessage.value =
          `There is no redline for ${row.label} — that comparison did not run.`;
        return;
      }
      try {
        const produced = redlineOf(row) ?? (await produceArtefact(row.id));
        if (produced === null) {
          errorMessage.value =
            'That redline could not be produced. Run the comparisons again to recover it.';
          return;
        }
        await downloadArtefact({ ...row, artefacts: [produced] });
        errorMessage.value = null;
      } catch (err: unknown) {
        errorMessage.value = `That redline could not be produced: ${describeError(err)}`;
      }
    },

    async downloadAll(): Promise<void> {
      const all = index.value?.rows ?? rows.value;
      await downloadRowArchive(all, 'bundle', row => produceArtefact(row.id));
    },
  };
}
