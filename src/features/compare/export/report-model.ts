/**
 * Pure Compare export report-model builder.
 *
 * Projects a completed comparison (the FULL `ops` + summary + per-side
 * ParaRecords) into a flat, framework-agnostic {@link CompareReport} that a
 * render target (the print view, A4) turns into a downloadable artefact.
 *
 * THE ADR-0008 WALL LIVES HERE. This builder reads `ops` directly and ignores
 * EVERY filter signal (`hideFormattingOnly` / `collapseRenumbering`). The exported artefact always carries the full
 * classified change set — identical whether the on-screen filters are ON or
 * OFF. Filters are navigation only and never reach this module.
 *
 * Pure: no Preact, no Office JS, no Date.now() (the generated date is injected).
 */
import type { ChangeClass, ChangeSummary, DiffOp, OpType, ParaRecord, PropertyChange, WordDiff } from '../model.js';
import { textWithObjects } from '../diff/object-edits.js';

/**
 * The read-only artefact framing (ADR-0020): the PDF-modes **candidate-differences**
 * report (explicitly NOT a redline — a PDF has no tracked-changes target). The
 * Word/Text rendered-redline snapshot is now produced programmatically by
 * RenderCore (feature 16, ADR-0024), not by this model.
 */
export type CompareReportKind = 'candidate-differences';

/** The compare mode that produced the report. */
export type CompareReportMode = 'pdf' | 'text' | 'word';

/** One change row in the report — a flat projection of a non-equal DiffOp. */
export interface ReportRow {
  /** Stable op id (for keys / navigation parity with the panel). */
  id: string;
  /** Structural op type. */
  type: OpType;
  /** Primary (priority-resolved) classification. */
  classification: ChangeClass;
  /** Short classification tag (e.g. 'NUM'), mirrors the on-screen ChangeList. */
  classTag: string;
  /** Best-effort location citation (part + paragraph index). */
  location: string;
  /** Original-side text (empty for pure inserts). */
  original: string;
  /** Revised-side text (empty for pure deletes). */
  revised: string;
  /** Low-confidence (fuzzy-anchored / approximate table cell) — badged, never hidden. */
  lowConfidence: boolean;
  /** Move target citation (only for type === 'move'). */
  moveTarget?: string;
}

/** The complete, render-target-agnostic export model. */
export interface CompareReport {
  kind: CompareReportKind;
  /** Human mode label, e.g. 'PDF-vs-PDF comparison'. */
  modeLabel: string;
  /** Artefact title, e.g. 'Candidate differences report'. */
  title: string;
  /** Provenance line — generated locally, with the injected date. */
  generatedNote: string;
  /** Blind-spot disclosure lines (verbatim from the shared safety-copy). */
  disclosures: string[];
  /** Per-side image-presence notice lines (PDF modes only). */
  imageNotice?: string[];
  /**
   * Total classified change count == `rows.length`: one row per non-equal BODY
   * op, filter-independent.
   *
   * This is `summary.total` for every path that diffs only the body. It is less
   * than `summary.total` when the auxiliary parts changed, because
   * `mergePartsIntoSummary` counts header/footer/footnote ops into the summary
   * (§8.2, §9.8) and those ops are not rows here — a part op's range indexes its
   * own part's paragraphs, and this model reads only the body's records. Only
   * the PDF modes build this report, and they have no parts; the Word-vs-Word
   * change list and change-set export resolve a part op through its own records
   * (ADR-0047, `CompareDriverState.parts`).
   */
  changeCount: number;
  /** Every classified change — the full set, filter-independent (ADR-0008). */
  rows: ReportRow[];
}

/** Inputs to {@link buildCompareReport}. NO filter signals — by design (ADR-0008). */
export interface BuildCompareReportInput {
  mode: CompareReportMode;
  /** The FULL ordered op list (after classify + cascade). */
  ops: readonly DiffOp[];
  /**
   * The change summary.
   *
   * **Currently read by nothing in this module** — `changeCount` below is
   * `rows.length`, and no cross-check against `summary.total` exists despite
   * what this comment used to say. Kept required because every production
   * caller already passes it and callers may come to depend on it; flagged
   * here rather than removed, which is an interface change for its owner.
   */
  summary: ChangeSummary;
  /** Original-side ParaRecords — resolve text for pure deletes/moves. */
  originalParas: readonly ParaRecord[];
  /** Revised-side ParaRecords — resolve text for pure inserts/moves. */
  revisedParas: readonly ParaRecord[];
  /** Blind-spot disclosure lines (from the shared safety-copy). */
  disclosures: string[];
  /** Per-side image-presence lines (PDF modes). */
  imageNotice?: string[];
  /** Pre-formatted generation date (injected — never Date.now() in a pure fn). */
  generatedAt: string;
}

const CLASS_TAG: Record<ChangeClass, string> = {
  structural: 'STRUCT',
  obligationLanguage: 'OBLIG',
  numericalValue: 'NUM',
  definedTerm: 'TERM',
  crossRefRenumber: 'XREF',
  substantiveText: 'TEXT',
  whitespacePunct: 'WS',
  formatting: 'FMT',
};

const MODE_LABEL: Record<CompareReportMode, string> = {
  pdf: 'PDF-vs-PDF comparison',
  text: 'Text-vs-text comparison',
  word: 'Word-vs-Word comparison',
};

/**
 * Build the export report model from a completed comparison.
 *
 * ADR-0008: reads `input.ops` directly; takes no filter flags; produces one row
 * per non-equal op, regardless of any on-screen filtering. `rows.length` is
 * therefore the full BODY change count — equal to `summary.total` unless an
 * auxiliary part changed too; see `CompareReport.changeCount`.
 */
export function buildCompareReport(input: BuildCompareReportInput): CompareReport {
  // The full classified change set — every non-equal op, no filter applied.
  const rows = input.ops
    .filter(op => op.type !== 'equal')
    .map(op => toRow(op, input.originalParas, input.revisedParas));

  const report: CompareReport = {
    kind: 'candidate-differences',
    modeLabel: MODE_LABEL[input.mode],
    title: 'Candidate differences report',
    generatedNote: `Generated locally on ${input.generatedAt} — nothing left your machine.`,
    disclosures: input.disclosures,
    changeCount: rows.length,
    rows,
  };
  if (input.imageNotice && input.imageNotice.length > 0) report.imageNotice = input.imageNotice;
  return report;
}

function toRow(op: DiffOp, originalParas: readonly ParaRecord[], revisedParas: readonly ParaRecord[]): ReportRow {
  const row: ReportRow = {
    id: op.id,
    type: op.type,
    classification: op.classification,
    classTag: CLASS_TAG[op.classification],
    location: locationFor(op),
    original: origText(op, originalParas),
    revised: revText(op, revisedParas),
    lowConfidence: op.lowConfidence === true || op.fuzzyAnchored === true,
  };
  if (op.type === 'move' && op.moveTargetId) row.moveTarget = op.moveTargetId;
  return row;
}

/** Best-effort citation: the part id plus the leading paragraph index. */
function locationFor(op: DiffOp): string {
  const range = op.origRange ?? op.revRange;
  if (range === null) return op.partId;
  return `${op.partId} ¶${range.start + 1}`;
}

function joinSide(wordDiffs: WordDiff[] | undefined, keep: ReadonlyArray<WordDiff['op']>): string | null {
  if (!wordDiffs) return null;
  return wordDiffs
    .filter(d => keep.includes(d.op))
    .map(d => d.text)
    .join('');
}

/**
 * What a `property` op (ADR-0042) changed, as a before/after pair. Its two
 * paragraphs are word-for-word identical, so falling through to the paragraph
 * text would print a row whose `original` and `revised` are the same string —
 * the exact shape the AI export refuses on the grounds that it reads as a
 * mistake. The report prints the property instead.
 *
 * Which is not automatic: the identity can change while the rendered label
 * stands still (ADR-0042 § "Where the key still exceeds the label"), and then
 * the property is the same string on both sides too. The revised side says so
 * rather than repeating itself.
 */
function labelOf(v: PropertyChange['from']): string {
  if (v.numberLabel === null) return v.ilvl === null ? '(unnumbered)' : `level ${v.ilvl + 1}`;
  return v.numberLabel === '' ? '(no visible number)' : v.numberLabel;
}

function propertySides(
  op: DiffOp,
  origPara: ParaRecord | undefined,
  revPara: ParaRecord | undefined,
): { original: string; revised: string } | null {
  const pc = op.propertyChange;
  if (op.type !== 'property' || !pc) return null;
  const sides = pc.numberLabel || pc.numberingIdentity || pc.styleId
    ? numberingAndStyleSides(pc)
    : { original: '', revised: '' };
  // ADR-0045 § Amendment — a footnote, picture or object added, removed or
  // moved with the words untouched: the paragraph's text, each object written
  // in at its place, so the two cells differ exactly where the document does.
  if (pc.objects === true) {
    const join = (a: string, b: string): string => (a === '' ? b : `${a} · ${b}`);
    return {
      original: join(sides.original, textWithObjects(origPara?.normText ?? '', pc.from.objects ?? [])),
      revised: join(sides.revised, textWithObjects(revPara?.normText ?? '', pc.to.objects ?? [])),
    };
  }
  return sides;
}

function numberingAndStyleSides(pc: PropertyChange): { original: string; revised: string } {
  const side = (v: PropertyChange['from']): string => {
    const parts: string[] = [];
    if (pc.numberLabel || pc.numberingIdentity) {
      // The level when no label resolved, so the row is not "(unnumbered)" on
      // both sides for a paragraph that genuinely changed level — and a WORD for
      // the label that resolved to the empty string, which is a different fact
      // and prints as nothing at all. A level at `numFmt="none"` (or with an
      // empty `lvlText`) renders `''`, and the row then had a blank original
      // cell beside a revised cell reading " (list definition changed)" — the
      // exact shape this function exists to prevent, reached from the other end.
      parts.push(labelOf(v));
    }
    if (pc.styleId) parts.push(`style: ${v.styleId ?? '(none)'}`);
    return parts.join(' · ');
  };
  let original = side(pc.from);
  let revised = side(pc.to);
  // Equal sides have two different causes and the row must not conflate them: a
  // RE-LEVELLING onto a level that happens to render the same label, or a list
  // DEFINITION change under a paragraph whose number did not move. `ilvl` is
  // carried on both sides to tell them apart.
  if (original === revised && pc.from.ilvl !== pc.to.ilvl) {
    const lvl = (v: PropertyChange['from']): string =>
      v.ilvl === null ? 'no level' : `level ${v.ilvl + 1}`;
    original = `${original} · ${lvl(pc.from)}`;
    revised = `${revised} · ${lvl(pc.to)}`;
  }
  return {
    original,
    revised: revised === original ? `${revised} (list definition changed)` : revised,
  };
}

/** Original-side text: word-diff (del+eq) for replaces, whole paragraph otherwise. */
function origText(op: DiffOp, paras: readonly ParaRecord[]): string {
  const property = propertySides(op, op.origRange ? paras[op.origRange.start] : undefined, undefined);
  if (property) return property.original;
  const fromDiffs = joinSide(op.wordDiffs, ['del', 'eq']);
  if (fromDiffs !== null) return fromDiffs;
  if (op.origRange === null) return '';
  return paras[op.origRange.start]?.normText ?? '';
}

/** Revised-side text: word-diff (ins+eq) for replaces, whole paragraph otherwise. */
function revText(op: DiffOp, paras: readonly ParaRecord[]): string {
  const property = propertySides(op, undefined, op.revRange ? paras[op.revRange.start] : undefined);
  if (property) return property.revised;
  const fromDiffs = joinSide(op.wordDiffs, ['ins', 'eq']);
  if (fromDiffs !== null) return fromDiffs;
  if (op.revRange === null) return '';
  return paras[op.revRange.start]?.normText ?? '';
}
