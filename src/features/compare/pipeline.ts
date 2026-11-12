/**
 * Feature 08 — Compare: pure pipeline orchestration (I1).
 *
 * Wires the 5-stage pipeline:
 *   preprocess → structural → word-diff (lazy) → moves → classify → cascade → author.
 *
 * Pure: no Office JS, no Date.now(). Author/date injected by the driver.
 * Yields to the event loop between word-diff chunks so the UI stays responsive (§9.12).
 * Spec: spec/features/08-compare.md §5.
 */
import type { DiffOp, ParaRecord, ChangeSummary, ChangeClass, CascadeAnnotation } from './model.js';
import type { CompareParts } from './io/index.js';
import type { CompareConfig } from '@shared/config/index.js';
import { preprocessSide, buildNumberingContext } from './preprocess/index.js';
import {
  structuralDiff, gapPairingFor, wordDiff, detectMoves, demoteDissimilarReplaces,
  detectMergesAndSplits, detectPropertyChanges, rangeDiffText, PARA_SEAM,
  insertBeforeDelete,
} from './diff/index.js';
import { analyseWordTables } from './diff/word-table.js';
import { classifyAll, applySgUkRules, computeCascades } from './classify/index.js';
import { authorBody, makeAuthorContext } from './output/index.js';

export interface PipelineInput {
  originalParts: CompareParts;
  revisedParts: CompareParts;
  config: CompareConfig;
  author: string;
  date: string;
  onProgress?: (pct: number) => void;
  abortSignal?: { aborted: boolean };
}

export interface PipelineResult {
  ops: DiffOp[];
  originalParas: ParaRecord[];
  revisedParas: ParaRecord[];
  summary: ChangeSummary;
  authoredBodyXml: string;
  /**
   * Emitted w:id → change classes for the revisions in `authoredBodyXml`
   * (RenderCore feature 16, Phase 2). Side-channel only: capturing it leaves
   * `authoredBodyXml` byte-identical. Lets the renderer attach styling
   * annotations keyed by the native w:id; it never filters (ADR-0008).
   */
  classByRevId: ReadonlyMap<string, readonly ChangeClass[]>;
  quarantinedOriginal: number;
  quarantinedRevised: number;
}

const WORD_DIFF_CHUNK = 25;

function checkAbort(signal?: { aborted: boolean }): void {
  if (signal?.aborted) throw new DOMException('Compare cancelled', 'AbortError');
}

export async function runPipeline(input: PipelineInput): Promise<PipelineResult> {
  const { originalParts, revisedParts, config, author, date, onProgress, abortSignal } = input;
  const report = (pct: number) => onProgress?.(Math.round(pct));

  // Stage 1 — Pre-process both sides (0–20 %)
  report(0);
  // ADR-0042 — styles + numbering for each side, so a paragraph's EFFECTIVE
  // numbering (which in a firm template lives on the style, not the paragraph)
  // is resolvable. Only this Word-vs-Word entry point has the parts to build
  // them; every other caller of `preprocessSide` passes nothing and behaves
  // exactly as before.
  const origNumbering = buildNumberingContext(
    originalParts.stylesXml, originalParts.numberingXml,
  );
  const revNumbering = buildNumberingContext(
    revisedParts.stylesXml, revisedParts.numberingXml,
  );
  const origResult = preprocessSide(
    originalParts.documentXml, 'document', config.flattenOriginal, config, origNumbering,
  );
  const revResult = preprocessSide(
    revisedParts.documentXml, 'document', config.flattenRevised, config, revNumbering,
  );
  checkAbort(abortSignal);
  report(20);

  // Stages 2–7 run on the pre-built records (shared core, mapped onto 20–100 %).
  return runComparePipeline({
    originalParas: origResult.paras,
    revisedParas: revResult.paras,
    config,
    author,
    date,
    quarantinedOriginal: origResult.quarantinedCount,
    quarantinedRevised: revResult.quarantinedCount,
    onProgress: (pct) => report(20 + pct * 0.8),
    abortSignal,
  });
}

/** Input for {@link runComparePipeline} — stages 2–7 over pre-built records. */
export interface ComparePipelineCoreInput {
  originalParas: ParaRecord[];
  revisedParas: ParaRecord[];
  config: CompareConfig;
  author: string;
  date: string;
  /** Quarantined-comment counts surfaced in the summary (0 for synthesised sources). */
  quarantinedOriginal?: number;
  quarantinedRevised?: number;
  onProgress?: ((pct: number) => void) | undefined;
  abortSignal?: { aborted: boolean } | undefined;
}

/** Input for {@link analyseRecords} — the analysis stages over pre-built records. */
export interface AnalyseInput {
  originalParas: ParaRecord[];
  revisedParas: ParaRecord[];
  config: CompareConfig;
  author: string;
  date: string;
  abortSignal?: { aborted: boolean } | undefined;
  /** Progress on the 0–100 scale `runComparePipeline` reports (25 … 88 here). */
  report?: ((pct: number) => void) | undefined;
  /**
   * The records are one Word table cell's content (ADR-0046), whose last
   * paragraph can never carry a mark revision; Stage 2.5 keeps a pair that
   * includes it inline instead of demoting it.
   */
  inCell?: boolean;
}

/**
 * Stages 2–6.5 of the Compare pipeline: alignment, merge/split, demotion,
 * property changes, Word tables, word diff, moves, classification, cascades and
 * mark order. Everything `runComparePipeline` does before authoring, and the
 * ONE place those stages are listed — the Word table stage (ADR-0046) runs a
 * cell's content back through this same function, so a cell is compared by
 * exactly the rules the body is. Returns the ops and the cascades.
 */
export async function analyseRecords(
  input: AnalyseInput,
): Promise<{ ops: DiffOp[]; cascades: CascadeAnnotation[] }> {
  const { originalParas, revisedParas, config, author, date, abortSignal } = input;
  const report = input.report ?? ((): void => undefined);

  // Stage 2 — Level-1 structural diff. A gap is paired by similarity, not by
  // position (ADR-0039): positional pairs that Stages 2.25 and 2.5 would render
  // inline are locked (by running those two passes), and the rest is aligned
  // through Stage 2.5's Dice threshold so no pair chosen here is demoted there.
  // The PDF pipeline below keeps positional pairing (ADR-0017).
  const pairing = gapPairingFor(config, originalParas, revisedParas);
  let ops: DiffOp[] = structuralDiff(originalParas, revisedParas, author, date, pairing);
  // Stage 2.25 — paragraph merge / split (ADR-0032): re-pair
  // replace+delete / replace+insert into ONE range-spanning `replace` so a join
  // reads as the small punctuation edit it is instead of striking a paragraph
  // and re-inserting the same words. Runs BEFORE the threshold so the threshold
  // scores the merge against the joined original text. A record headed for a
  // Level-3 move is absorbed only where positional pairing absorbed it too
  // (ADR-0039 §1(b)).
  ops = detectMergesAndSplits(ops, originalParas, revisedParas, config, pairing.mayAbsorb);
  // Stage 2.5 — change-complexity threshold (ADR-0026): demote a too-dissimilar
  // `replace` pair to a clean whole-paragraph delete + insert so a wholesale
  // rewrite reads cleanly instead of as token-by-token "diff salad".
  ops = demoteDissimilarReplaces(ops, originalParas, revisedParas, config, { keepLastParagraph: input.inCell === true });
  // Stage 2.75 — non-text property changes (ADR-0042): a matched pair with
  // IDENTICAL text whose style or resolved numbering differs becomes a
  // `property` op instead of an invisible `equal`. Runs last of the structural
  // stages so it only ever sees pairs the alignment already settled on, and it
  // re-pairs nothing — the authored XML is byte-identical either way.
  ops = detectPropertyChanges(ops, originalParas, revisedParas);
  // Stage 2.8 — Word tables (ADR-0046): every op pairing two Word table units
  // gets its row and cell alignment, each matched cell's content analysed by
  // this very function. Before the word diff and classification, which read
  // the result; table units are never word-diffed themselves (ADR-0027 guard).
  ops = await analyseWordTables(ops, originalParas, revisedParas, config, (orig, rev) =>
    analyseRecords({ originalParas: orig, revisedParas: rev, config, author, date, abortSignal, inCell: true }).then((r) => r.ops),
  );
  checkAbort(abortSignal);
  report(25);

  // Stage 3 — Level-2 word diff on replace pairs (lazy + yielding).
  await applyWordDiffs(ops, originalParas, revisedParas, config, abortSignal, (done, total) =>
    report(25 + Math.round((done / Math.max(total, 1)) * 35)),
  );
  report(60);

  // Stage 4 — Level-3 verbatim move detection.
  ops = detectMoves(ops, originalParas, revisedParas);
  checkAbort(abortSignal);
  report(70);

  // Stage 5 — Classify: taxonomy (B1) → SG/UK rules (B2).
  ops = classifyAll(ops, originalParas, revisedParas, config);
  ops = ops.map(op => applySgUkRules(op, originalParas, revisedParas, config));
  checkAbort(abortSignal);
  report(80);

  // Stage 6 — Renumbering-cascade collapse (B3).
  const { ops: cascadedOps, cascades } = computeCascades(ops, config);
  ops = cascadedOps;
  checkAbort(abortSignal);
  report(88);

  // Stage 6.5 — Mark order (ADR-0043): the insertion reads before the deletion
  // it replaces. Here, not in `word.ts` (the classifiers would eat the new order
  // for nothing) and not in `author.ts` (the side-by-side panel and the exports
  // would then disagree with the redline) — so every consumer downstream of this
  // line sees ONE order. This one call also serves Text-vs-text and the batch
  // driver, both of which enter through this function.
  //
  // BEFORE `spellOutSeams` below, which is what still lets the pass see a
  // `PARA_SEAM` and refuse to move a paragraph mark (ADR-0032).
  applyMarkOrder(ops);

  return { ops, cascades };
}

/**
 * Stages 2–7 of the Compare pipeline over already-built `ParaRecord[]`. Shared by
 * the Word-vs-Word path (`runPipeline`, after preprocess) and the Text-vs-text
 * Markdown path (ADR-0027, which builds records via `io/text/extract-text-records`
 * so a table is one alignment unit). The stages and their output are byte-identical
 * to the prior inline implementation for records without `.table`; the table-aware
 * branches in the word-diff, move, classify and author stages fire only when a
 * record carries `.table`, which Word-vs-Word never produces. Reports 0–100 %.
 */
export async function runComparePipeline(input: ComparePipelineCoreInput): Promise<PipelineResult> {
  const { originalParas, revisedParas, config, author, date, onProgress, abortSignal } = input;
  const quarantinedOriginal = input.quarantinedOriginal ?? 0;
  const quarantinedRevised = input.quarantinedRevised ?? 0;
  const report = (pct: number) => onProgress?.(Math.round(pct));

  report(0);

  const { ops, cascades } = await analyseRecords({
    originalParas, revisedParas, config, author, date, abortSignal, report,
  });

  // Stage 7 — Author the redline body XML.
  // captureClasses builds the w:id → change-class side channel for RenderCore;
  // it does not change `authoredBodyXml` (the editable redline) by one byte.
  const ctx = makeAuthorContext(author, date, 1, { captureClasses: true });
  const authoredBodyXml = authorBody(ops, originalParas, revisedParas, ctx);
  spellOutSeams(ops);
  report(98);

  const summary = buildSummary(ops, cascades, quarantinedOriginal, quarantinedRevised);
  report(100);

  return {
    ops,
    originalParas,
    revisedParas,
    summary,
    authoredBodyXml,
    classByRevId: ctx.classByRevId ?? new Map(),
    quarantinedOriginal,
    quarantinedRevised,
  };
}

/**
 * Stage 3 word diff on replace ops, chunked + yielding (shared by the Word and
 * PDF pipelines). Mutates `op.wordDiffs` in place. `onChunk(done, total)` reports
 * progress after each yielded chunk.
 */
async function applyWordDiffs(
  ops: DiffOp[],
  originalParas: ParaRecord[],
  revisedParas: ParaRecord[],
  config: CompareConfig,
  abortSignal: { aborted: boolean } | undefined,
  onChunk?: (done: number, total: number) => void,
): Promise<void> {
  const replaceOps = ops.filter((op) => op.type === 'replace');
  for (let i = 0; i < replaceOps.length; i += WORD_DIFF_CHUNK) {
    checkAbort(abortSignal);
    const chunk = replaceOps.slice(i, i + WORD_DIFF_CHUNK);
    for (const op of chunk) {
      const origPara = op.origRange ? originalParas[op.origRange.start] : undefined;
      const revPara = op.revRange ? revisedParas[op.revRange.start] : undefined;
      // ADR-0027: skip word diff for table-unit records (text-vs-text Markdown) —
      // their normText is a table flattening; the table author handles cell diffs.
      // Inert for Word-vs-Word (no table-unit records).
      if (origPara !== undefined && revPara !== undefined && !origPara.table && !revPara.table) {
        // `rangeDiffText` is the record's own normText for an ordinary one-to-one
        // pair, and the range joined by PARA_SEAM markers for a merge/split — so
        // the paragraph mark is diffed like any other token.
        op.wordDiffs = wordDiff(
          rangeDiffText(originalParas, op.origRange),
          rangeDiffText(revisedParas, op.revRange),
          config,
        );
      }
    }
    await Promise.resolve(); // yield to event loop between chunks
    onChunk?.(Math.min(i + WORD_DIFF_CHUNK, replaceOps.length), replaceOps.length);
  }
}

/**
 * Stage 6.5 — re-order each op's word diff so an insertion precedes the deletion
 * it replaces (§5.5, ADR-0043).
 *
 * `insertBeforeDelete` is **not idempotent** on two changed regions with no
 * equality between them, so it must be applied exactly once per op. That is why
 * this is one line in each pipeline and not a helper anything else may call.
 */
function applyMarkOrder(ops: DiffOp[]): void {
  for (const op of ops) {
    if (op.wordDiffs === undefined) continue;
    op.wordDiffs = insertBeforeDelete(op.wordDiffs);
  }
}

/**
 * Replace every {@link PARA_SEAM} in `op.wordDiffs` with a plain newline, once
 * the author has consumed it.
 *
 * The sentinel is an **authoring device**: it carries the paragraph mark through
 * the Level-2 diff so `authorBody` can emit `<w:pPr><w:rPr><w:del/></w:rPr>` on
 * it (ADR-0032). It is a Private-Use-Area code point, and every OTHER consumer
 * of `wordDiffs` is user- or LLM-visible — the AI-export change set, the
 * exported compare report, the change-list row preview, and `goToChange`'s
 * host-search anchor, which can never match document text while it holds one.
 * So the seam stops at the author, spelled as the paragraph break it stands for.
 *
 * `authoredBodyXml` is already built at this point and is not touched.
 */
function spellOutSeams(ops: DiffOp[]): void {
  for (const op of ops) {
    if (op.wordDiffs === undefined) continue;
    // Only an op the merge/split pass built can legitimately hold a seam: it is
    // the join between two records, so the op must CLAIM more than one. A
    // one-to-one op carrying U+E000 is carrying the DOCUMENT'S own character
    // (`detectMergesAndSplits` declines to re-pair such a record precisely so it
    // stays one-to-one), and rewriting that to a newline would corrupt the text
    // in every derived view. Word-level text is never re-authored here, so the
    // rewrite has to be provably confined to text this pipeline synthesised.
    if (!spansMultipleRecords(op)) continue;
    if (!op.wordDiffs.some((wd) => wd.text.includes(PARA_SEAM))) continue;
    op.wordDiffs = op.wordDiffs.map((wd) =>
      wd.text.includes(PARA_SEAM) ? { ...wd, text: wd.text.split(PARA_SEAM).join('\n') } : wd,
    );
  }
}

/** Does this op claim more than one paragraph record on either side? */
function spansMultipleRecords(op: DiffOp): boolean {
  const spans = (r: { start: number; end: number } | null | undefined): boolean =>
    r !== null && r !== undefined && r.end > r.start;
  return spans(op.origRange) || spans(op.revRange);
}

// ---------------------------------------------------------------------------
// PDF pipeline (feature 12, ADR-0017) — pre-built ParaRecord[], no Stage 1
// (preprocessSide needs documentXml) and no Stage 7 (no redline to author).
// ---------------------------------------------------------------------------

export interface PdfPipelineInput {
  /** Already-reconstructed paragraph records (from io/pdf/extractPdfParaRecords). */
  originalParas: ParaRecord[];
  revisedParas: ParaRecord[];
  config: CompareConfig;
  author: string;
  date: string;
  /** Optional fuzzy-anchor pass (B3) applied to the structural ops before word diff. */
  fuzzyAnchor?: (ops: DiffOp[], original: ParaRecord[], revised: ParaRecord[]) => DiffOp[];
  onProgress?: (pct: number) => void;
  abortSignal?: { aborted: boolean };
}

/**
 * PDF result — no `authoredBodyXml` (PDFs have no tracked-changes target). The
 * shared UI consumes `ops`/`summary`/paras, identical to the Word result.
 */
export interface PdfPipelineResult {
  ops: DiffOp[];
  originalParas: ParaRecord[];
  revisedParas: ParaRecord[];
  summary: ChangeSummary;
}

export async function runPipelinePdf(input: PdfPipelineInput): Promise<PdfPipelineResult> {
  const { originalParas, revisedParas, config, author, date, fuzzyAnchor, onProgress, abortSignal } =
    input;
  const report = (pct: number): void => onProgress?.(Math.round(pct));

  report(0);

  // Stage 2 — Level-1 structural diff (0–35 %).
  let ops: DiffOp[] = structuralDiff(originalParas, revisedParas, author, date);
  checkAbort(abortSignal);
  report(35);

  // Stage 2.5 — Fuzzy-anchor fallback (B3): re-pair spurious delete+insert that
  // arose only from reconstruction boundary noise. Never suppresses (ADR-0008).
  if (fuzzyAnchor !== undefined) {
    ops = fuzzyAnchor(ops, originalParas, revisedParas);
    checkAbort(abortSignal);
  }
  report(40);

  // Stage 3 — Level-2 word diff on replace pairs (40–70 %).
  await applyWordDiffs(ops, originalParas, revisedParas, config, abortSignal, (done, total) =>
    report(40 + Math.round((done / Math.max(total, 1)) * 30)),
  );
  report(70);

  // Stage 4 — Level-3 verbatim move detection (70–75 %).
  ops = detectMoves(ops, originalParas, revisedParas);
  checkAbort(abortSignal);
  report(75);

  // Stage 5 — Classify (B1) → SG/UK rules (B2) (75–85 %).
  ops = classifyAll(ops, originalParas, revisedParas, config);
  ops = ops.map((op) => applySgUkRules(op, originalParas, revisedParas, config));
  checkAbort(abortSignal);
  report(85);

  // Stage 6 — Renumbering-cascade collapse (85–95 %).
  const { ops: cascadedOps, cascades } = computeCascades(ops, config);
  ops = cascadedOps;

  // Stage 6.5 — Mark order (ADR-0043), as in `runComparePipeline`. The PDF path
  // authors no redline, but its `wordDiffs` drive the side-by-side panel, the
  // candidate-differences report and the AI export, which must read in the same
  // order as the Word path's.
  applyMarkOrder(ops);
  report(95);

  // NO Stage 7 — no redline authored for PDF. Quarantined counts are 0 (PDFs
  // carry no tracked changes / comments to quarantine).
  const summary = buildSummary(ops, cascades, 0, 0);
  report(100);

  return { ops, originalParas, revisedParas, summary };
}

function buildSummary(
  ops: DiffOp[],
  cascades: CascadeAnnotation[],
  quarantinedOriginal: number,
  quarantinedRevised:  number,
): ChangeSummary {
  const nonEqual = ops.filter(op => op.type !== 'equal');
  const total = nonEqual.length;

  const byCategory: Record<ChangeClass, number> = {
    structural: 0,
    obligationLanguage: 0,
    numericalValue: 0,
    definedTerm: 0,
    crossRefRenumber: 0,
    substantiveText: 0,
    whitespacePunct: 0,
    formatting: 0,
  };
  for (const op of nonEqual) {
    byCategory[op.classification] = (byCategory[op.classification] ?? 0) + 1;
  }

  const movedParagraphs = Math.ceil(ops.filter(op => op.type === 'move').length / 2);
  const spellingVariantCount = nonEqual.filter(op => op.spellingVariant !== undefined).length;

  return {
    total,
    byCategory,
    renumberingCascades: cascades,
    movedParagraphs,
    spellingVariantCount,
    quarantinedCommentsOriginal: quarantinedOriginal,
    quarantinedCommentsRevised:  quarantinedRevised,
    // Defaults. The auxiliary parts are not this function's to diff — it works
    // on records, and the parts are files. The Word-vs-Word driver fills these
    // in from `diffPartsIfEnabled`/`diffStyles`; the Text and PDF paths have no
    // such parts and leave them at zero.
    partChanges: { headers: 0, footers: 0, footnotes: 0 },
    styleDefinitionNote: null,
    filtersActive: {
      hideFormattingOnly: true,
      collapseRenumbering: true,
    },
  };
}
