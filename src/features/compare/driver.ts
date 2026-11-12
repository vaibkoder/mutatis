/**
 * Feature 08 — Compare: driver (I1).
 *
 * Module-level Preact signals + actions implementing CompareDriverState.
 * Wires: load (S2) → pipeline (I1) → parts-diff (I2) → materialise (C2).
 * Signals pattern mirrors repair/driver.ts.
 *
 * Spec: spec/features/08-compare.md §8.
 */
import { signal } from '@preact/signals';
import type {
  CompareDriverState,
  FlattenChoice,
  CompareStatus,
  TextCompareDriverState,
  PdfCompareDriverState,
  PdfImageNotice,
  PartView,
  RenderSource,
} from './ui-types.js';
import type { ChangeClass, DiffOp, ChangeSummary, ParaRecord, WordDiff } from './model.js';
import type { CompareDirection, CompareParts } from './io/index.js';
import type { AuxiliaryParts } from './output/index.js';
import { loadBothSides, extractParts } from './io/index.js';
import type { DocxBundle } from '@shared/ooxml/index.js';
import { PizZipBundle } from '@shared/ooxml/bundle.js';
import { parseNumberingScheme } from '@shared/numbering/index.js';
import { getActiveHost } from '@shared/host/index.js';
import { getActiveProfile, getCompareConfig } from '@shared/config/index.js';
import { runPipeline, runPipelinePdf, runComparePipeline } from './pipeline.js';
import { synthesiseDocx } from './io/text/synthesise-docx.js';
import { extractTextRecords } from './io/text/extract-text-records.js';
import { diffAndAuthorParts, mergePartsIntoSummary } from './parts-diff.js';
import { diffStyles } from './styles-diff.js';
import { materialise, relationshipIdAllocator } from './output/index.js';
import { planHeaderFooters } from './output/package-parts.js';
import { tableAnchorDiffs } from './diff/table-summary.js';
import { extractPdfParaRecords } from './io/pdf/parts.js';
import { openPdfDocument, loadPdfImageOps } from './io/pdf/pdfjs-loader.js';
import { detectImagePresence } from './io/pdf/images.js';
import { fuzzyAnchorPdf } from './pdf/fuzzy-anchor.js';

// ---------------------------------------------------------------------------
// Error description — Office JS surfaces the real cause on `OfficeExtension.Error.
// debugInfo` (code / message / errorLocation), NOT on `.message` (which is the
// generic "Sorry, something went wrong…" string). Pull it out so a host writeBack
// failure is diagnosable, and log the full error for the console.
// ---------------------------------------------------------------------------

function describeError(err: unknown): string {
  if (!(err instanceof Error)) return String(err);
  const di = (err as { debugInfo?: { code?: string; message?: string; errorLocation?: string } }).debugInfo;
  if (di) {
    console.error('[Compare] Office JS error:', err, di);
    const detail = [di.code, di.message, di.errorLocation ? `at ${di.errorLocation}` : '']
      .filter(Boolean).join(' — ');
    if (detail) return `${err.message} (${detail})`;
  }
  return err.message;
}

// ---------------------------------------------------------------------------
// Internal source tracking (separate from CompareSource to allow lazy File reads)
// ---------------------------------------------------------------------------

type SideSource =
  | { kind: 'activeDocument' }
  | { kind: 'docxFile'; file: File };

// ---------------------------------------------------------------------------
// RenderCore (feature 16, Phase 2) — assemble the read-only render source.
// Reads parts from the ALREADY-LOADED source bundle; does not touch the
// materialise / openRedline paths. Numbering parse failures degrade to no
// labels (never throw) so a render never blocks on numbering.xml quirks.
// ---------------------------------------------------------------------------

function buildRenderSource(
  authoredBodyXml: string,
  bundle: DocxBundle,
  classByRevId: ReadonlyMap<string, readonly ChangeClass[]>,
  authoredFootnotesXml?: string,
): RenderSource {
  let numbering: RenderSource['numbering'];
  try {
    numbering = parseNumberingScheme(bundle);
  } catch {
    numbering = undefined; // best-effort; absent numbering → non-list paragraphs
  }
  const stylesXml = bundle.part('word/styles.xml');
  // ADR-0047: the redline's own notes part when Compare authored one — the
  // body's references point into it, not into the original's.
  const footnotesXml = authoredFootnotesXml ?? bundle.part('word/footnotes.xml');
  // Spread optional keys conditionally — exactOptionalPropertyTypes forbids
  // assigning an explicit `undefined` to an optional property.
  return {
    authoredBodyXml,
    bundle,
    classByRevId,
    ...(stylesXml !== undefined ? { stylesXml } : {}),
    ...(footnotesXml !== undefined ? { footnotesXml } : {}),
    ...(numbering !== undefined ? { numbering } : {}),
  };
}

/** The footnotes part the redline carries, when Compare authored one (ADR-0047). */
function authoredFootnotes(aux: AuxiliaryParts, original: CompareParts): string | undefined {
  const path = original.notesPaths?.footnotes ?? null;
  if (path !== null) return aux.replace?.get(path);
  return aux.add?.find((p) => p.relType === 'footnotes')?.xml;
}

async function resolveSideSource(src: SideSource): Promise<CompareDirection['original']> {
  if (src.kind === 'activeDocument') return { kind: 'activeDocument' };
  const buffer = await src.file.arrayBuffer();
  return { kind: 'docxArrayBuffer', buffer };
}

/** Word characters in a candidate anchor — how distinctive it can possibly be. */
function anchorWeight(text: string): number {
  return (text.match(/[\p{L}\p{N}]/gu) ?? []).length;
}

/**
 * The text to hand `findAndReveal` so that "go to change" lands on the change.
 *
 * The host search takes the **first** match in the document
 * (`office-host.ts findAndReveal` selects `items[0]`), so the anchor has to be
 * distinctive or navigation goes somewhere arbitrary. Two things make the
 * obvious choice — "the first changed span" — wrong:
 *
 *  - a merge/split op's first mark spans a paragraph break (ADR-0032), which
 *    arrives here as a newline, and the host search cannot cross one;
 *  - that first mark is very often bare punctuation. For a paragraph join it is
 *    `"."`, which selects the **first full stop in the document**.
 *
 * So: prefer the first changed line carrying enough word characters to identify
 * a place, and otherwise fall back to the longest UNCHANGED line of the op —
 * which at least lands the reader in the right clause — rather than to
 * punctuation.
 */
export function searchAnchorFor(wordDiffs: readonly WordDiff[]): string {
  const lines = (predicate: (d: WordDiff) => boolean): string[] =>
    wordDiffs
      .filter(predicate)
      .flatMap(d => d.text.split('\n'))
      .map(line => line.trim())
      .filter(line => line !== '');

  const MIN_WEIGHT = 4;
  const changed = lines(d => d.op !== 'eq');
  const distinctive = changed.find(line => anchorWeight(line) >= MIN_WEIGHT);
  if (distinctive !== undefined) return distinctive;

  // Nothing distinctive changed — a punctuation-only edit, or a paragraph join
  // whose only mark is the lost full stop. Anchor on the op's own unchanged
  // text; the longest line is the likeliest to be unique in the document.
  const unchanged = lines(d => d.op === 'eq').sort((a, b) => anchorWeight(b) - anchorWeight(a));
  if (unchanged[0] !== undefined && anchorWeight(unchanged[0]) >= MIN_WEIGHT) return unchanged[0];

  // Genuinely nothing to aim at. Better to do nothing than to select an
  // arbitrary character somewhere else in the document.
  return '';
}

// ---------------------------------------------------------------------------
// Exported driver state object (injected into ComparePanel in I3)
// ---------------------------------------------------------------------------

export function createCompareDriver(): CompareDriverState {
  let _originalSource: SideSource | null = null;
  let _revisedSource:  SideSource | null = null;
  let _abortSignal: { aborted: boolean } | null = null;

  const compareStatus       = signal<CompareStatus>('idle');
  const compareProgress     = signal<number>(0);
  const compareError        = signal<string | null>(null);
  const compareOps          = signal<DiffOp[] | null>(null);
  const compareSummary      = signal<ChangeSummary | null>(null);
  const compareOriginalParas = signal<ParaRecord[]>([]);
  const compareRevisedParas  = signal<ParaRecord[]>([]);
  const compareParts         = signal<ReadonlyMap<string, PartView>>(new Map());
  const compareRenderSource  = signal<RenderSource | null>(null);
  const hideFormattingOnly  = signal<boolean>(true);
  const collapseRenumbering = signal<boolean>(true);
  const originalHasRevisions = signal<boolean>(false);
  const revisedHasRevisions  = signal<boolean>(false);
  const flattenOriginal     = signal<FlattenChoice>('acceptAll');
  const flattenRevised      = signal<FlattenChoice>('acceptAll');
  const originalReady       = signal<boolean>(false);
  const revisedReady        = signal<boolean>(false);

  function useCurrentDocument(side: 'original' | 'revised'): void {
    if (side === 'original') {
      _originalSource = { kind: 'activeDocument' };
      originalReady.value = true;
    } else {
      _revisedSource = { kind: 'activeDocument' };
      revisedReady.value = true;
    }
  }

  function uploadDocx(side: 'original' | 'revised', file: File): void {
    if (side === 'original') {
      _originalSource = { kind: 'docxFile', file };
      originalReady.value = true;
    } else {
      _revisedSource = { kind: 'docxFile', file };
      revisedReady.value = true;
    }
    void detectRevisions(side, file);
  }

  async function detectRevisions(side: 'original' | 'revised', file: File): Promise<void> {
    try {
      const buf = await file.arrayBuffer();
      const bundle = new PizZipBundle(buf);
      const docXml = bundle.part('word/document.xml') ?? '';
      const hasRevisions = /<w:ins\b|<w:del\b|<w:moveFrom\b|<w:moveTo\b/.test(docXml);
      if (side === 'original') originalHasRevisions.value = hasRevisions;
      else revisedHasRevisions.value = hasRevisions;
    } catch {
      // Best-effort — no errors surfaced for revision detection
    }
  }

  function cancelCompare(): void {
    if (_abortSignal) _abortSignal.aborted = true;
    compareStatus.value   = 'idle';
    compareProgress.value = 0;
  }

  async function runCompare(): Promise<void> {
    if (compareStatus.value === 'running') return;
    if (!_originalSource || !_revisedSource) {
      compareError.value  = 'Select or upload both documents before running Compare.';
      compareStatus.value = 'error';
      return;
    }

    compareStatus.value   = 'running';
    compareProgress.value = 0;
    compareError.value    = null;
    compareRenderSource.value = null;
    _abortSignal = { aborted: false };
    const abort = _abortSignal;

    try {
      const profile = getActiveProfile();
      const config  = getCompareConfig(profile);
      const date    = new Date().toISOString();
      const author  = config.author;

      // Resolve file sources: File → ArrayBuffer (consumed on load)
      const direction: CompareDirection = {
        original: await resolveSideSource(_originalSource),
        revised:  await resolveSideSource(_revisedSource),
      };

      const { originalBundle, revisedBundle } = await loadBothSides(direction);
      if (abort.aborted) return;

      const originalParts = extractParts(originalBundle);
      const extractedRevised = extractParts(revisedBundle);
      // ADR-0047 — where each header and footer goes in the redline package,
      // and the revised body's section references pointed there BEFORE the
      // pipeline writes them (they carry the revised package's relationship ids).
      const newRelationshipId = relationshipIdAllocator(originalBundle.part('word/_rels/document.xml.rels'));
      const plan = planHeaderFooters(
        originalParts, extractedRevised, config, originalBundle.listParts(), newRelationshipId,
      );
      const revisedParts = { ...extractedRevised, documentXml: plan.revisedDocumentXml };

      // Main body pipeline (reports 0–90 %)
      const result = await runPipeline({
        originalParts,
        revisedParts,
        config,
        author,
        date,
        onProgress: pct => { compareProgress.value = Math.round(pct * 0.9); },
        abortSignal: abort,
      });
      if (abort.aborted) return;

      // Parts (I2, ADR-0047): headers/footers/footnotes compared AND authored
      // into the redline's own parts, the body's note references retargeted;
      // and the styles.xml diff. Both results reach the summary.
      const partsDiff = await diffAndAuthorParts({
        originalParts, revisedParts, bodyXml: result.authoredBodyXml, config, author, date,
        newRelationshipId, plan, bundles: { original: originalBundle, revised: revisedBundle },
      });
      if (abort.aborted) return;
      const stylesDiff = diffStyles(
        originalParts.stylesXml, revisedParts.stylesXml, author, date,
      );

      compareProgress.value = 95;

      // RenderCore render source (read-only; captured from the SAME body that
      // materialise writes, so rendered + editable redlines stay in lock-step).
      // Set before materialise so an Office writeBack failure cannot strand it.
      compareRenderSource.value = buildRenderSource(
        partsDiff.bodyXml, originalBundle, result.classByRevId, authoredFootnotes(partsDiff.aux, originalParts),
      );

      // Materialise as a new standalone document (default; §9.13)
      await materialise(partsDiff.bodyXml, originalBundle, {
        mode: 'newDocument',
        author,
        date,
      }, partsDiff.aux);
      if (abort.aborted) return;

      // Each header, footer and note change is a row of its own (ADR-0047),
      // after the body's. Its ranges index its OWN part's records, so the part
      // goes into `compareParts` under the ops' `partId`: every reader of a
      // range (paragraph text, the table anchor, the export) resolves through
      // `recordsFor`, never the body's arrays.
      compareParts.value   = new Map(partsDiff.parts.map(set => [set.partId, set]));
      compareOps.value     = [...result.ops, ...partsDiff.ops.filter(op => op.type !== 'equal')];
      compareSummary.value = mergePartsIntoSummary(
        { ...result.summary, styleDefinitionNote: stylesDiff.note },
        partsDiff,
      );
      compareOriginalParas.value = result.originalParas;
      compareRevisedParas.value  = result.revisedParas;
      compareStatus.value  = 'done';
      compareProgress.value = 100;
    } catch (err: unknown) {
      if (err instanceof DOMException && err.name === 'AbortError') {
        compareStatus.value   = 'idle';
        compareProgress.value = 0;
        return;
      }
      compareError.value  = describeError(err);
      compareStatus.value = 'error';
    }
  }

  /** The records an op's ranges index: its own part's, or the body's. */
  function recordsFor(op: DiffOp): { originalParas: ParaRecord[]; revisedParas: ParaRecord[] } {
    return compareParts.value.get(op.partId)
      ?? { originalParas: compareOriginalParas.value, revisedParas: compareRevisedParas.value };
  }

  /** The matched paragraph's own text, revised side preferred (it is the live doc). */
  function paraTextFor(op: DiffOp): string[] {
    const { originalParas, revisedParas } = recordsFor(op);
    const rev = op.revRange ? revisedParas[op.revRange.start] : undefined;
    const orig = op.origRange ? originalParas[op.origRange.start] : undefined;
    const text = rev?.normText ?? orig?.normText ?? '';
    return text === '' ? [] : [text];
  }

  function goToChange(opId: string): void {
    const ops = compareOps.value;
    if (ops === null) return;
    const op = ops.find(o => o.id === opId);
    if (!op) return;

    // A property op (ADR-0042) has no word diffs, so there is no changed text to
    // aim at — but the paragraph is right there and unchanged. Anchor on its own
    // words, or the row is counted, rendered, and silently does nothing when
    // clicked. `searchAnchorFor` picks the most distinctive line either way.
    // A Word table op (ADR-0046) aims inside one cell: Word's search cannot
    // cross a cell boundary, so the table's own text would never be found.
    const searchText =
      op.tableDiff !== undefined
        ? searchAnchorFor(tableAnchorDiffs(op, recordsFor(op).originalParas, recordsFor(op).revisedParas))
        : op.type === 'property'
          ? searchAnchorFor(paraTextFor(op).map(text => ({ op: 'eq' as const, text })))
          : searchAnchorFor(op.wordDiffs ?? []);
    if (!searchText) return;

    // Navigate via the active host's free-text search (office: body.search+select),
    // inside the header, footer or notes a part op lives in (ADR-0047).
    // On the web target there is no live host doc to navigate, so findAndReveal is
    // omitted (undefined) → a clean no-op.
    void getActiveHost().findAndReveal?.(searchText, compareParts.value.get(op.partId)?.location.scope);
  }

  return {
    status:               compareStatus,
    progress:             compareProgress,
    errorMessage:         compareError,
    ops:                  compareOps,
    summary:              compareSummary,
    originalParas:        compareOriginalParas,
    revisedParas:         compareRevisedParas,
    parts:                compareParts,
    renderSource:         compareRenderSource,
    hideFormattingOnly,
    collapseRenumbering,
    originalHasRevisions,
    revisedHasRevisions,
    flattenOriginal,
    flattenRevised,
    originalReady,
    revisedReady,
    useCurrentDocument,
    uploadDocx,
    runCompare,
    cancelCompare,
    goToChange,
  };
}

/**
 * The process-wide Word-vs-Word driver the add-in shell and the web Compare
 * view are wired to. It is one instance of {@link createCompareDriver}, not a
 * special case — batch work (ADR-0037) creates its own instances, and the
 * per-signal exports below are aliases onto this one so existing consumers and
 * tests are unaffected.
 */
export const compareDriverState: CompareDriverState = createCompareDriver();

export const compareStatus        = compareDriverState.status;
export const compareProgress      = compareDriverState.progress;
export const compareError         = compareDriverState.errorMessage;
export const compareOps           = compareDriverState.ops;
export const compareSummary       = compareDriverState.summary;
export const compareOriginalParas = compareDriverState.originalParas;
export const compareRevisedParas  = compareDriverState.revisedParas;
export const compareRenderSource  = compareDriverState.renderSource;
export const hideFormattingOnly   = compareDriverState.hideFormattingOnly;
export const collapseRenumbering  = compareDriverState.collapseRenumbering;
export const originalHasRevisions = compareDriverState.originalHasRevisions;
export const revisedHasRevisions  = compareDriverState.revisedHasRevisions;
export const flattenOriginal      = compareDriverState.flattenOriginal;
export const flattenRevised       = compareDriverState.flattenRevised;
export const originalReady        = compareDriverState.originalReady;
export const revisedReady         = compareDriverState.revisedReady;

// ===========================================================================
// Feature 12 (Compare — additional sources): Text-vs-text driver (ADR-0018, T3)
// ===========================================================================
//
// Paste two plain-text blobs → run the EXISTING pipeline on synthesised .docx
// bundles → hold the authored redline in memory. The redline is generated
// eagerly on Compare but materialised (opened) ONLY on the user's click via
// openRedline() (TT5). Fully client-side: synthesise + pipeline are pure; the
// only Office bridge is materialise → createDocument().open() (ADR-0001/0003).

/**
 * Per-side character cap for a pasted blob. A named constant (not a magic
 * literal) per spec §7; oversize input raises a typed limit error (TT7) before
 * anything is synthesised, so nothing partial is generated.
 */
export const MAX_TEXT_PASTE_CHARS = 500_000;

/** Friendly message shown when both boxes are empty/whitespace-only (TT6). */
const EMPTY_INPUT_MESSAGE =
  'Paste text into both the Original and Modified boxes to compare.';

/** Typed error raised when a pasted blob exceeds {@link MAX_TEXT_PASTE_CHARS}. */
export class TextPasteTooLargeError extends Error {
  constructor(
    readonly side: 'original' | 'revised',
    readonly limit: number,
    readonly actual: number,
  ) {
    super(
      `The ${side} text is too large to compare (${actual.toLocaleString()} characters; ` +
        `limit ${limit.toLocaleString()}). Please compare a smaller passage.`,
    );
    this.name = 'TextPasteTooLargeError';
  }
}

/** Throw a typed limit error if a side exceeds the cap. Pure — exported for tests. */
export function checkPasteSize(
  side: 'original' | 'revised',
  text: string,
  limit: number = MAX_TEXT_PASTE_CHARS,
): void {
  if (text.length > limit) throw new TextPasteTooLargeError(side, limit, text.length);
}

// Held redline — the authored body + the Original synthesised bundle that
// materialise() clones. NOT a signal: it is heavy and non-reactive; the panel
// learns of its availability via textRedlineReady. Cleared on every new run.
interface HeldRedline {
  authoredBodyXml: string;
  originalBundle: DocxBundle;
  classByRevId: ReadonlyMap<string, readonly ChangeClass[]>;
  author: string;
  date: string;
}
export function createTextCompareDriver(): TextCompareDriverState {
  let _heldRedline: HeldRedline | null = null;
  let _textAbort: { aborted: boolean } | null = null;

  const textStatus       = signal<CompareStatus>('idle');
  const textProgress     = signal<number>(0);
  const textError        = signal<string | null>(null);
  const textOps          = signal<DiffOp[] | null>(null);
  const textSummary      = signal<ChangeSummary | null>(null);
  const textOriginal     = signal<string>('');
  const textRevised      = signal<string>('');
  const textRedlineReady = signal<boolean>(false);
  const textOriginalParas = signal<ParaRecord[]>([]);
  const textRevisedParas  = signal<ParaRecord[]>([]);
  const textRenderSource  = signal<RenderSource | null>(null);
  const textHideFormattingOnly  = signal<boolean>(true);
  const textCollapseRenumbering = signal<boolean>(true);

  async function runTextCompare(): Promise<void> {
    if (textStatus.value === 'running') return;

    const original = textOriginal.value;
    const revised  = textRevised.value;

    textStatus.value   = 'running';
    textProgress.value = 0;
    textError.value    = null;
    textRedlineReady.value = false;
    textRenderSource.value = null;
    _heldRedline = null;
    _textAbort = { aborted: false };
    const abort = _textAbort;

    try {
      // Size cap first — nothing is synthesised if either side is oversize (TT7).
      checkPasteSize('original', original);
      checkPasteSize('revised', revised);

      // Both sides empty/whitespace-only → clean no-op with a message (TT6).
      if (original.trim() === '' && revised.trim() === '') {
        textError.value  = EMPTY_INPUT_MESSAGE;
        textStatus.value = 'error';
        textProgress.value = 0;
        return;
      }

      const profile = getActiveProfile();
      const config  = getCompareConfig(profile);
      const date    = new Date().toISOString();
      const author  = config.author;

      // ADR-0027: build records directly from the paste — paragraphs AND Markdown
      // tables (each table is one alignment unit) — then run the shared pipeline
      // core. The Original bundle is still synthesised: materialise()/
      // buildRenderSource() clone it as the redline package scaffold (its body is
      // replaced by the authored redline, so it carries only the package parts).
      const originalBundle = synthesiseDocx(original);
      if (abort.aborted) return;

      const originalParas = extractTextRecords(original);
      const revisedParas  = extractTextRecords(revised);

      const result = await runComparePipeline({
        originalParas,
        revisedParas,
        config,
        author,
        date,
        onProgress: pct => { textProgress.value = Math.round(pct); },
        abortSignal: abort,
      });
      if (abort.aborted) return;

      // Eager-generate, click-to-open: HOLD the redline; do NOT materialise (TT5).
      _heldRedline = {
        authoredBodyXml: result.authoredBodyXml, originalBundle,
        classByRevId: result.classByRevId, author, date,
      };
      // Surface the render source for the paginated preview + PDF snapshot.
      textRenderSource.value = buildRenderSource(
        result.authoredBodyXml, originalBundle, result.classByRevId,
      );

      textOps.value         = result.ops;
      textSummary.value     = result.summary;
      // Per-side ParaRecords for the in-pane inline preview (ADR-0019).
      textOriginalParas.value = result.originalParas;
      textRevisedParas.value  = result.revisedParas;
      textRedlineReady.value = true;
      textStatus.value      = 'done';
      textProgress.value    = 100;
    } catch (err: unknown) {
      if (err instanceof DOMException && err.name === 'AbortError') {
        textStatus.value   = 'idle';
        textProgress.value = 0;
        return;
      }
      textError.value  = describeError(err);
      textStatus.value = 'error';
    }
  }

  function cancelTextCompare(): void {
    if (_textAbort) _textAbort.aborted = true;
    textStatus.value   = 'idle';
    textProgress.value = 0;
  }

  /**
   * Materialise the held redline as a NEW standalone document — invoked only on
   * the user's explicit click (never auto-opened). Reuses output/writeout's
   * materialise unchanged (createDocument().open()).
   */
  async function openRedline(): Promise<void> {
    const held = _heldRedline;
    if (held === null || !textRedlineReady.value) return;
    await materialise(held.authoredBodyXml, held.originalBundle, {
      mode: 'newDocument',
      author: held.author,
      date: held.date,
    });
  }

  /** Text mode has no host document to navigate; the redline opens as a new doc. */
  function textGoToChange(_opId: string): void {
    /* no-op in text mode */
  }

  return {
    status:              textStatus,
    progress:            textProgress,
    errorMessage:        textError,
    ops:                 textOps,
    summary:             textSummary,
    hideFormattingOnly:  textHideFormattingOnly,
    collapseRenumbering: textCollapseRenumbering,
    originalText:        textOriginal,
    revisedText:         textRevised,
    redlineReady:        textRedlineReady,
    originalParas:       textOriginalParas,
    revisedParas:        textRevisedParas,
    renderSource:        textRenderSource,
    runCompare:          runTextCompare,
    cancelCompare:       cancelTextCompare,
    openRedline,
    goToChange:          textGoToChange,
  };
}

/** The process-wide Text-vs-text driver. One instance of the factory above. */
export const textCompareDriverState: TextCompareDriverState = createTextCompareDriver();

export const textStatus              = textCompareDriverState.status;
export const textProgress            = textCompareDriverState.progress;
export const textError               = textCompareDriverState.errorMessage;
export const textOps                 = textCompareDriverState.ops;
export const textSummary             = textCompareDriverState.summary;
export const textOriginal            = textCompareDriverState.originalText;
export const textRevised             = textCompareDriverState.revisedText;
export const textRedlineReady        = textCompareDriverState.redlineReady;
export const textOriginalParas       = textCompareDriverState.originalParas;
export const textRevisedParas        = textCompareDriverState.revisedParas;
export const textRenderSource        = textCompareDriverState.renderSource;
export const textHideFormattingOnly  = textCompareDriverState.hideFormattingOnly;
export const textCollapseRenumbering = textCompareDriverState.collapseRenumbering;

// ===========================================================================
// Feature 12 (Compare — additional sources): PDF-vs-PDF driver
// ===========================================================================
//
// Upload two PDFs → extract+reconstruct ParaRecords per side (A4) → runPipelinePdf
// (B2, with the B3 fuzzy fallback) → hold ops/summary + the reconstructed paras
// for the in-pane side-by-side view (C1). There is NO redline output — PDFs have
// no tracked-changes target. A scanned/image-only PDF raises NoExtractableTextError
// (P7), surfaced verbatim as a clear unsupported message. A render-free image
// presence scan (D3) produces the blind-spot disclosure. Fully client-side: the
// only external surface is the local, confidentiality-hardened PDF.js opener.

/** Open helper used as the injected PdfOpener (confidentiality-hardened loader). */
const browserPdfOpener = (data: Uint8Array) => openPdfDocument(data);

export function createPdfCompareDriver(): PdfCompareDriverState {
  let _pdfOriginalFile: File | null = null;
  let _pdfRevisedFile: File | null = null;
  let _pdfAbort: { aborted: boolean } | null = null;

  const pdfStatus        = signal<CompareStatus>('idle');
  const pdfProgress      = signal<number>(0);
  const pdfError         = signal<string | null>(null);
  const pdfOps           = signal<DiffOp[] | null>(null);
  const pdfSummary       = signal<ChangeSummary | null>(null);
  const pdfOriginalParas = signal<ParaRecord[]>([]);
  const pdfRevisedParas  = signal<ParaRecord[]>([]);
  const pdfImageNotice   = signal<PdfImageNotice | null>(null);
  const pdfOriginalReady = signal<boolean>(false);
  const pdfRevisedReady  = signal<boolean>(false);
  const pdfHideFormattingOnly  = signal<boolean>(true);
  const pdfCollapseRenumbering = signal<boolean>(true);
  const pdfSelectedOpId  = signal<string | null>(null);

  function uploadPdf(side: 'original' | 'revised', file: File): void {
    if (side === 'original') {
      _pdfOriginalFile = file;
      pdfOriginalReady.value = true;
    } else {
      _pdfRevisedFile = file;
      pdfRevisedReady.value = true;
    }
  }

  async function runPdfCompare(): Promise<void> {
    if (pdfStatus.value === 'running') return;
    if (!_pdfOriginalFile || !_pdfRevisedFile) {
      pdfError.value  = 'Upload both PDFs before running Compare.';
      pdfStatus.value = 'error';
      return;
    }

    pdfStatus.value    = 'running';
    pdfProgress.value  = 0;
    pdfError.value     = null;
    pdfImageNotice.value = null;
    pdfSelectedOpId.value = null;
    _pdfAbort = { aborted: false };
    const abort = _pdfAbort;

    try {
      const profile = getActiveProfile();
      const config  = getCompareConfig(profile);
      const date    = new Date().toISOString();
      const author  = config.author;

      // Read each file ONCE; PDF.js detaches the buffer it is handed, so every
      // consumer (extract, image-scan) gets its own copy via slice().
      const origMaster = new Uint8Array(await _pdfOriginalFile.arrayBuffer());
      const revMaster  = new Uint8Array(await _pdfRevisedFile.arrayBuffer());
      if (abort.aborted) return;

      const original = await extractPdfParaRecords(origMaster.slice(), browserPdfOpener, config, abort);
      pdfProgress.value = 25;
      if (abort.aborted) return;
      const revised = await extractPdfParaRecords(revMaster.slice(), browserPdfOpener, config, abort);
      pdfProgress.value = 40;
      if (abort.aborted) return;

      const result = await runPipelinePdf({
        originalParas: original.paras,
        revisedParas:  revised.paras,
        config,
        author,
        date,
        fuzzyAnchor: fuzzyAnchorPdf,
        onProgress: pct => { pdfProgress.value = 40 + Math.round(pct * 0.45); },
        abortSignal: abort,
      });
      if (abort.aborted) return;

      // D3 — render-free image-presence disclosure (best-effort; never blocks).
      const ops = await loadPdfImageOps();
      const [origImages, revImages] = await Promise.all([
        detectImagePresence(origMaster.slice(), browserPdfOpener, ops, undefined, abort),
        detectImagePresence(revMaster.slice(), browserPdfOpener, ops, undefined, abort),
      ]);
      if (abort.aborted) return;

      pdfOriginalParas.value = result.originalParas;
      pdfRevisedParas.value  = result.revisedParas;
      pdfImageNotice.value   = { original: origImages, revised: revImages };
      pdfOps.value     = result.ops;
      pdfSummary.value = result.summary;
      pdfStatus.value  = 'done';
      pdfProgress.value = 100;
    } catch (err: unknown) {
      if (err instanceof DOMException && err.name === 'AbortError') {
        pdfStatus.value   = 'idle';
        pdfProgress.value = 0;
        return;
      }
      // NoExtractableTextError (P7) and input-cap errors carry friendly messages.
      pdfError.value  = describeError(err);
      pdfStatus.value = 'error';
    }
  }

  function cancelPdfCompare(): void {
    if (_pdfAbort) _pdfAbort.aborted = true;
    pdfStatus.value   = 'idle';
    pdfProgress.value = 0;
  }

  /** PDF mode has no host document; selecting a change drives the in-pane detail. */
  function pdfGoToChange(opId: string): void {
    pdfSelectedOpId.value = opId;
  }

  return {
    status:              pdfStatus,
    progress:            pdfProgress,
    errorMessage:        pdfError,
    ops:                 pdfOps,
    summary:             pdfSummary,
    hideFormattingOnly:  pdfHideFormattingOnly,
    collapseRenumbering: pdfCollapseRenumbering,
    originalParas:       pdfOriginalParas,
    revisedParas:        pdfRevisedParas,
    imageNotice:         pdfImageNotice,
    originalReady:       pdfOriginalReady,
    revisedReady:        pdfRevisedReady,
    uploadPdf,
    runCompare:          runPdfCompare,
    cancelCompare:       cancelPdfCompare,
    goToChange:          pdfGoToChange,
  };
}

/** The process-wide PDF-vs-PDF driver. One instance of the factory above. */
export const pdfCompareDriverState: PdfCompareDriverState = createPdfCompareDriver();

export const pdfStatus              = pdfCompareDriverState.status;
export const pdfProgress            = pdfCompareDriverState.progress;
export const pdfError               = pdfCompareDriverState.errorMessage;
export const pdfOps                 = pdfCompareDriverState.ops;
export const pdfSummary             = pdfCompareDriverState.summary;
export const pdfOriginalParas       = pdfCompareDriverState.originalParas;
export const pdfRevisedParas        = pdfCompareDriverState.revisedParas;
export const pdfImageNotice         = pdfCompareDriverState.imageNotice;
export const pdfOriginalReady       = pdfCompareDriverState.originalReady;
export const pdfRevisedReady        = pdfCompareDriverState.revisedReady;
export const pdfHideFormattingOnly  = pdfCompareDriverState.hideFormattingOnly;
export const pdfCollapseRenumbering = pdfCompareDriverState.collapseRenumbering;
