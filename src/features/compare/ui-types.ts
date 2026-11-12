/**
 * Feature 08 — Compare: typed driver interface + UI signal shapes (C6).
 *
 * The ComparePanel imports ONLY from this file, model.ts, and Preact. The real
 * driver (Wave 2 / I1) implements `CompareDriverState`; until then the panel
 * runs against the stub returned by `makeStubDriverState()`.
 *
 * ADR-0008: the filter signals here change ONLY navigation-list visibility.
 * They never alter `ops` and never reach the output/author layer.
 *
 * Spec: spec/features/08-compare.md §8.
 */
import { signal, type Signal } from '@preact/signals';
import type { ChangeClass, ChangeSummary, DiffOp, ParaRecord } from './model.js';
import type { ImagePresenceReport } from './io/pdf/images.js';
import type { DocxBundle } from '@shared/ooxml/index.js';
import type { NumberingScheme } from '@shared/numbering/index.js';
import type { SearchScope } from '@shared/host/index.js';

// ---------------------------------------------------------------------------
// RenderCore (feature 16, ADR-0024, Phase 2) — the render source surfaced by
// the driver so the Compare→RenderCore adapter can paint a paginated preview
// and a vector PDF from the SAME `authoredBodyXml` that backs the editable
// redline (the two stay in lock-step). Carrying this changes NOTHING about the
// existing materialise / openRedline / host-doc redline paths — it is captured
// alongside them, read-only.
// ---------------------------------------------------------------------------

/** Everything RenderCore needs to render the redline, surfaced by the driver. */
export interface RenderSource {
  /** Merged `<w:body>` with native revision markup (the Compare authoredBodyXml). */
  authoredBodyXml: string;
  /** `word/styles.xml` of the source package (full style cascade). */
  stylesXml?: string;
  /** `word/footnotes.xml` of the source package, if present. */
  footnotesXml?: string;
  /** Parsed numbering scheme (real auto-numbering labels). */
  numbering?: NumberingScheme;
  /** Emitted w:id → change classes (styling annotations only; never filters). */
  classByRevId?: ReadonlyMap<string, readonly ChangeClass[]>;
  /** The source package — for embedded `.odttf` font extraction (B2). */
  bundle: DocxBundle;
}

/** One header, footer or note the change list shows rows from (ADR-0047). */
export interface PartView {
  originalParas: ParaRecord[];
  revisedParas: ParaRecord[];
  location: {
    /** The row's badge, e.g. "Footer" or "Footnote 3 (added)". */
    label: string;
    /** The badge's tooltip: what "go to change" does there. */
    hint: string;
    /** The story the host searches. */
    scope: SearchScope;
  };
}

/** Per-side flatten policy for pre-existing tracked changes (§5.0, §8.1). */
export type FlattenChoice = 'acceptAll' | 'rejectAll';

/** Run status of a compare operation. */
export type CompareStatus = 'idle' | 'running' | 'done' | 'error';

/** UI state the driver exposes to the ComparePanel. */
export interface CompareDriverState {
  /** Lifecycle of the current run. */
  status: Signal<CompareStatus>;
  /** Progress 0–100 during a run (§9.12). */
  progress: Signal<number>;
  /** Error message when status === 'error'. */
  errorMessage: Signal<string | null>;
  /** The full ordered op list (after classify + cascade). null before first run. */
  ops: Signal<DiffOp[] | null>;
  /** The change summary. null before first run. */
  summary: Signal<ChangeSummary | null>;

  /**
   * Per-side ParaRecords from the last run — let the change-set AI export
   * (feature 11) resolve whole-paragraph before/after text for pure
   * insert/delete/move ops (which carry no wordDiffs). Mirrors the text/PDF
   * drivers. Empty until the first run.
   */
  originalParas: Signal<ParaRecord[]>;
  revisedParas: Signal<ParaRecord[]>;

  /**
   * The headers, footers and notes with changes, by part id (ADR-0047). A part
   * op's `partId` keys its part here: its ranges index THESE records, not the
   * body's, and its location labels its row. Empty until the first run.
   */
  parts: Signal<ReadonlyMap<string, PartView>>;

  /**
   * The render source for the rendered-redline snapshot (RenderCore, feature 16).
   * null before the first successful run. Read-only; surfacing it leaves the
   * editable-redline (materialise) path untouched.
   */
  renderSource: Signal<RenderSource | null>;

  // --- Filter toggles (presentation only — never alter ops or the redline) ---
  /** Hide formatting-only changes from the list. Default true. */
  hideFormattingOnly: Signal<boolean>;
  /** Collapse renumbering cascades in the list. Default true. */
  collapseRenumbering: Signal<boolean>;

  // --- Per-side flatten policy (shown only when that side has revisions) ---
  originalHasRevisions: Signal<boolean>;
  revisedHasRevisions: Signal<boolean>;
  flattenOriginal: Signal<FlattenChoice>;
  flattenRevised: Signal<FlattenChoice>;

  // --- Slot / source selection ---
  /** Whether the Original slot has a usable source selected. */
  originalReady: Signal<boolean>;
  /** Whether the Revised slot has a usable source selected. */
  revisedReady: Signal<boolean>;

  // --- Actions (implemented by I1; no-ops in the stub) ---
  /** Pick the active document for a side. */
  useCurrentDocument: (side: 'original' | 'revised') => void;
  /** Provide an uploaded .docx for a side (File from <input type=file>). */
  uploadDocx: (side: 'original' | 'revised', file: File) => void;
  /** Start a compare run. The driver validates inputs and runs the pipeline. */
  runCompare: () => void;
  /** Cancel a running compare. Discards partial state. */
  cancelCompare: () => void;
  /** Navigate to a change in the document (via the host's findAndReveal in I1). */
  goToChange: (opId: string) => void;
}

// ---------------------------------------------------------------------------
// Feature 12 (Compare — additional sources) — Text-vs-text mode (ADR-0018)
// ---------------------------------------------------------------------------

/** The display-only filter signals shared by both compare modes (ADR-0008). */
export interface FilterSignals {
  hideFormattingOnly: Signal<boolean>;
  collapseRenumbering: Signal<boolean>;
}

/**
 * UI state the text-vs-text driver exposes to the panel (feature 12, ADR-0018).
 *
 * Mirrors CompareDriverState's result surface (status/progress/ops/summary +
 * the three filter signals) but its inputs are two pasted plain-text blobs, and
 * its output is an eagerly-generated redline that is **held in memory and opened
 * only on `openRedline()`** — never materialised during `runCompare` (TT5).
 */
export interface TextCompareDriverState extends FilterSignals {
  status: Signal<CompareStatus>;
  progress: Signal<number>;
  errorMessage: Signal<string | null>;
  ops: Signal<DiffOp[] | null>;
  summary: Signal<ChangeSummary | null>;

  /** Original (base) pasted text. */
  originalText: Signal<string>;
  /** Modified (compared) pasted text. */
  revisedText: Signal<string>;
  /** True once a redline has been generated and is ready to open (eager-generate). */
  redlineReady: Signal<boolean>;
  /**
   * Per-side ParaRecords from the last run — let the in-pane inline redline
   * (ADR-0019) resolve plain text for pure insert/delete/move ops. Mirrors the
   * PDF driver. Empty until the first run.
   */
  originalParas: Signal<ParaRecord[]>;
  revisedParas: Signal<ParaRecord[]>;

  /**
   * The render source for the rendered-redline snapshot + inline preview
   * (RenderCore, feature 16). Derived from the held redline; null before a run.
   */
  renderSource: Signal<RenderSource | null>;

  /** Build both sides, run the pipeline, and HOLD the redline (no materialise). */
  runCompare: () => void;
  /** Cancel a running compare. Discards partial state. */
  cancelCompare: () => void;
  /** Materialise the held redline as a new standalone document — on user click only. */
  openRedline: () => void;
  /** Navigate to a change. No-op in text mode (the redline opens as a new doc). */
  goToChange: (opId: string) => void;
}

// ---------------------------------------------------------------------------
// Feature 12 — PDF-vs-PDF mode
// ---------------------------------------------------------------------------

/** The per-run image-presence disclosure, one report per side (D3). */
export interface PdfImageNotice {
  original: ImagePresenceReport;
  revised: ImagePresenceReport;
}

/**
 * UI state the PDF-vs-PDF driver exposes to the panel (feature 12).
 *
 * Mirrors the result surface (status/progress/ops/summary + the three display
 * filters) but its inputs are two uploaded PDFs and there is NO redline output
 * (PDFs have no tracked-changes target) — a change is shown in-pane via the
 * side-by-side view. It additionally surfaces the image-presence blind-spot
 * notice (D3) and holds the reconstructed ParaRecords for the detail view.
 */
export interface PdfCompareDriverState extends FilterSignals {
  status: Signal<CompareStatus>;
  progress: Signal<number>;
  errorMessage: Signal<string | null>;
  ops: Signal<DiffOp[] | null>;
  summary: Signal<ChangeSummary | null>;

  /** Reconstructed Original/Revised paragraphs — resolve text for the detail view. */
  originalParas: Signal<ParaRecord[]>;
  revisedParas: Signal<ParaRecord[]>;
  /** Per-run image-presence disclosure (null before first successful run). */
  imageNotice: Signal<PdfImageNotice | null>;

  /** Whether each slot has a usable PDF selected. */
  originalReady: Signal<boolean>;
  revisedReady: Signal<boolean>;

  /** Provide an uploaded .pdf for a side (File from <input type=file>). */
  uploadPdf: (side: 'original' | 'revised', file: File) => void;
  /** Build both sides, run the PDF pipeline, hold results. */
  runCompare: () => void;
  /** Cancel a running compare. Discards partial state. */
  cancelCompare: () => void;
  /** Select a change to view (no host-document navigation in PDF mode). */
  goToChange: (opId: string) => void;
}

/** Stub PDF-vs-PDF driver — inert signals and no-op actions (filters default ON). */
export function makeStubPdfDriverState(): PdfCompareDriverState {
  return {
    status: signal<CompareStatus>('idle'),
    progress: signal(0),
    errorMessage: signal<string | null>(null),
    ops: signal<DiffOp[] | null>(null),
    summary: signal<ChangeSummary | null>(null),

    hideFormattingOnly: signal(true),
    collapseRenumbering: signal(true),

    originalParas: signal<ParaRecord[]>([]),
    revisedParas: signal<ParaRecord[]>([]),
    imageNotice: signal<PdfImageNotice | null>(null),

    originalReady: signal(false),
    revisedReady: signal(false),

    uploadPdf: () => { /* no-op stub */ },
    runCompare: () => { /* no-op stub */ },
    cancelCompare: () => { /* no-op stub */ },
    goToChange: () => { /* no-op stub */ },
  };
}

/** Stub text-vs-text driver — inert signals and no-op actions (filters default ON). */
export function makeStubTextDriverState(): TextCompareDriverState {
  return {
    status: signal<CompareStatus>('idle'),
    progress: signal(0),
    errorMessage: signal<string | null>(null),
    ops: signal<DiffOp[] | null>(null),
    summary: signal<ChangeSummary | null>(null),

    hideFormattingOnly: signal(true),
    collapseRenumbering: signal(true),

    originalText: signal(''),
    revisedText: signal(''),
    redlineReady: signal(false),
    originalParas: signal<ParaRecord[]>([]),
    revisedParas: signal<ParaRecord[]>([]),
    renderSource: signal<RenderSource | null>(null),

    runCompare: () => { /* no-op stub */ },
    cancelCompare: () => { /* no-op stub */ },
    openRedline: () => { /* no-op stub */ },
    goToChange: () => { /* no-op stub */ },
  };
}

/**
 * A stub driver state for testing / storybook — inert signals and no-op actions.
 * Defaults match the spec: all three filters default ON.
 */
export function makeStubDriverState(): CompareDriverState {
  return {
    status: signal<CompareStatus>('idle'),
    progress: signal(0),
    errorMessage: signal<string | null>(null),
    ops: signal<DiffOp[] | null>(null),
    summary: signal<ChangeSummary | null>(null),
    originalParas: signal<ParaRecord[]>([]),
    revisedParas: signal<ParaRecord[]>([]),
    parts: signal<ReadonlyMap<string, PartView>>(new Map()),
    renderSource: signal<RenderSource | null>(null),

    hideFormattingOnly: signal(true),
    collapseRenumbering: signal(true),

    originalHasRevisions: signal(false),
    revisedHasRevisions: signal(false),
    flattenOriginal: signal<FlattenChoice>('acceptAll'),
    flattenRevised: signal<FlattenChoice>('acceptAll'),

    originalReady: signal(false),
    revisedReady: signal(false),

    useCurrentDocument: () => { /* no-op stub */ },
    uploadDocx: () => { /* no-op stub */ },
    runCompare: () => { /* no-op stub */ },
    cancelCompare: () => { /* no-op stub */ },
    goToChange: () => { /* no-op stub */ },
  };
}
