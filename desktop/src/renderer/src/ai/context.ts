import { serialiseChangeSet } from '@features/ai-export/serialise/change-set.js';
import type { ChangeEntry, ChangeSetExport } from '@features/ai-export/types.js';
import type { ChangeClass, ChangeSummary, DiffOp, ParaRecord } from '@features/compare/model.js';
import { matchKey, type StructureOutline } from '../structure/outline.js';

/**
 * The AI context payload for one comparison.
 *
 * This is what replaces RAG. Both documents are already in hand and already
 * diffed; retrieval over a corpus we do not have would answer a question nobody
 * asked. The model is handed the engine's own change set — before/after text
 * pairs, never OOXML markup, never strikethrough notation — plus the counts and
 * an explicit list of what this payload does NOT cover.
 *
 * Built in the WORKER, beside the pipeline that produced the ops, because that
 * is the only place `originalParas`/`revisedParas` exist. Carrying tens of
 * thousands of ParaRecords back across the boundary so the UI thread could do
 * the same work would be pure waste.
 *
 * Every change anchors to its `DiffOp.id`. There are no character offsets
 * anywhere in the engine, so an id is the only stable way for a comment to
 * point at a change — and it is the same id the redline uses.
 */

/** Characters of change text after which the payload discloses and stops. */
export const CONTEXT_CHAR_BUDGET = 120_000;

export interface RedlineContext {
  readonly generated: string;
  readonly jurisdiction: 'UK' | 'SG';
  readonly baseline: string;
  readonly revised: string;
  readonly totals: {
    readonly total: number;
    readonly byCategory: Record<string, number>;
    readonly moved: number;
    readonly spellingVariants: number;
  };
  readonly changes: readonly ChangeEntry[];
  /**
   * Changes the engine found but this payload does not carry, with the reason.
   * NEVER a silent truncation — the model is told the set is partial, so it can
   * say so rather than answering as though it had seen everything.
   */
  readonly omitted: number;
  /** What this payload cannot tell you. Always populated. */
  readonly blindSpots: readonly string[];
}

export interface ContextInput {
  readonly ops: readonly DiffOp[];
  readonly summary: ChangeSummary;
  readonly originalParas: readonly ParaRecord[];
  readonly revisedParas: readonly ParaRecord[];
  readonly baseline: string;
  readonly revised: string;
  readonly jurisdiction: 'UK' | 'SG';
  readonly quarantinedOriginal: number;
  readonly quarantinedRevised: number;
  /** ISO-8601, injected so a payload is reproducible in tests. */
  readonly generated: string;
  readonly charBudget?: number;
  /**
   * The structure analysis of each document, when it ran. Gives citations with
   * their schedule, annex and part ("Schedule 2 para 3(a)"); without it, the
   * citer falls back to each paragraph's own numbering prefix.
   */
  readonly structure?: { readonly original: StructureOutline | null; readonly revised: StructureOutline | null };
}

export function buildRedlineContext(input: ContextInput): RedlineContext {
  const cite = makeCiter(input.originalParas, input.revisedParas, input.structure);

  const full: ChangeSetExport = serialiseChangeSet(input.ops, {
    baseline: input.baseline,
    revised: input.revised,
    clauseFor: cite.clauseFor,
    origParaText: (op) => paraText(input.originalParas, op.origRange),
    revParaText: (op) => paraText(input.revisedParas, op.revRange),
    moveTarget: (op) => {
      const target = input.ops.find((candidate) => candidate.id === op.moveTargetId);
      return target ? cite.clauseFor(target) : '';
    },
  });

  const { kept, omitted } = withinBudget(full.changes, input.charBudget ?? CONTEXT_CHAR_BUDGET);

  return {
    generated: input.generated,
    jurisdiction: input.jurisdiction,
    baseline: input.baseline,
    revised: input.revised,
    totals: {
      total: input.summary.total,
      byCategory: countsOf(input.summary.byCategory),
      moved: input.summary.movedParagraphs,
      spellingVariants: input.summary.spellingVariantCount,
    },
    changes: kept,
    omitted,
    blindSpots: blindSpotsFor(input, cite.resolved, cite.fromStructure, omitted, kept.length),
  };
}

/** Drop the zero rows — a category with no changes is noise in a prompt. */
function countsOf(byCategory: Record<ChangeClass, number>): Record<string, number> {
  const out: Record<string, number> = {};
  for (const [key, value] of Object.entries(byCategory)) {
    if (value > 0) out[key] = value;
  }
  return out;
}

function paraText(
  paras: readonly ParaRecord[],
  range: { start: number; end: number } | null,
): string {
  if (!range) return '';
  const parts: string[] = [];
  for (let i = range.start; i <= range.end; i += 1) {
    const record = paras[i];
    if (record) parts.push(record.normText);
  }
  return parts.join(' ').trim();
}

/**
 * Clause citations for each change.
 *
 * The engine's own `makeCiter` wants a Section Index built from the Doc Model,
 * and the ONLY Doc Model loader upstream reads the live Word document through
 * Office JS. (An upstream gap, not a bug.)
 *
 * So the citation comes from our own structure analysis of each document
 * (structure/), which names schedules, annexes and parts and composes limbs
 * ("Schedule 2 Part A para 3(a)"). The engine's paragraphs are matched to the
 * analyser's by their text. Where there is no analysis, or a paragraph cannot
 * be matched, it falls back to the nearest numbering prefix — a bare "4.2"
 * with no region — and the blind-spot list says which of the two was used.
 */
function makeCiter(
  original: readonly ParaRecord[],
  revised: readonly ParaRecord[],
  structure: ContextInput['structure'],
) {
  let resolved = 0;
  let fromStructure = 0;
  const originalMap = structure?.original ? alignToOutline(original, structure.original) : null;
  const revisedMap = structure?.revised ? alignToOutline(revised, structure.revised) : null;

  const nearest = (paras: readonly ParaRecord[], from: number): string => {
    for (let i = Math.min(from, paras.length - 1); i >= 0; i -= 1) {
      const prefix = paras[i]?.numberingPrefix?.trim();
      if (prefix) return prefix.replace(/[.\s]+$/, '');
    }
    return '';
  };
  const citeIn = (
    paras: readonly ParaRecord[],
    map: readonly (string | null)[] | null,
    from: number,
  ): string => {
    const structured = map?.[from];
    if (structured) {
      fromStructure += 1;
      return structured;
    }
    return nearest(paras, from);
  };

  return {
    get resolved() {
      return resolved;
    },
    get fromStructure() {
      return fromStructure;
    },
    clauseFor(op: DiffOp): string {
      const label =
        op.revRange !== null
          ? citeIn(revised, revisedMap, op.revRange.start)
          : op.origRange !== null
            ? citeIn(original, originalMap, op.origRange.start)
            : '';
      if (label) resolved += 1;
      return label;
    },
  };
}

/**
 * The structure citation for each engine paragraph, found by walking both
 * paragraph lists in order and matching their text. Body paragraphs only: the
 * analyser does not read headers, footers or footnotes. An unmatched paragraph
 * takes the citation of the last matched one before it.
 */
export function alignToOutline(paras: readonly ParaRecord[], outline: StructureOutline): (string | null)[] {
  const keys = outline.paragraphKeys;
  const out: (string | null)[] = paras.map(() => null);
  let j = 0;
  let last: string | null = null;
  for (let i = 0; i < paras.length; i += 1) {
    const p = paras[i]!;
    if (p.isHeaderFooter || p.isFootnote) continue;
    const key = matchKey(p.normText);
    if (key !== '') {
      for (let m = j; m < keys.length && m < j + 60; m += 1) {
        if (keys[m] === key || (key.length >= 40 && keys[m]!.startsWith(key.slice(0, 40)))) {
          last = outline.paragraphCitations[m] ?? null;
          j = m + 1;
          break;
        }
      }
    }
    out[i] = last;
  }
  return out;
}

function withinBudget(
  changes: readonly ChangeEntry[],
  budget: number,
): { kept: ChangeEntry[]; omitted: number } {
  const kept: ChangeEntry[] = [];
  let used = 0;
  for (const change of changes) {
    used += (change.original?.length ?? 0) + (change.revised?.length ?? 0);
    if (used > budget && kept.length > 0) {
      return { kept, omitted: changes.length - kept.length };
    }
    kept.push(change);
  }
  return { kept, omitted: 0 };
}

function blindSpotsFor(
  input: ContextInput,
  resolvedCitations: number,
  structuredCitations: number,
  omitted: number,
  carried: number,
): string[] {
  const spots: string[] = [];

  if (structuredCitations > 0) {
    spots.push(
      'Clause citations come from a structure analysis of each document, which names ' +
        'Schedules, Annexes, Appendices and Parts ("Schedule 2 para 3(a)"). It is a heuristic ' +
        'reading of the layout, not the document\'s own cross-reference table: treat a citation ' +
        'as a pointer and say so if it looks wrong.' +
        (structuredCitations < resolvedCitations
          ? ' Some citations fell back to the paragraph\'s bare numbering prefix, which does not ' +
            'name the schedule a clause sits in.'
          : ''),
    );
  } else {
    spots.push(
      'Clause labels come from each paragraph\'s own numbering prefix. Schedules, Annexes ' +
        'and Appendices are NOT named, so a label like "4.2" may belong to a schedule rather ' +
        'than the body. Do not assert where a clause sits.',
    );
  }
  if (resolvedCitations === 0 && input.ops.length > 0) {
    spots.push('No clause numbering was found at all, so no change carries a citation.');
  }

  spots.push(
    'No defined-terms glossary, cross-reference check, factual-consistency check or party-role ' +
      'check has run. Do not claim a term is undefined or a reference is broken.',
  );

  if (input.quarantinedOriginal > 0 || input.quarantinedRevised > 0) {
    spots.push(
      `${input.quarantinedOriginal + input.quarantinedRevised} comment(s) were quarantined out ` +
        'before comparison and are not represented here.',
    );
  }

  if (omitted > 0) {
    spots.push(
      // NOT `total - omitted`. That subtraction assumed every counted change is
      // one this payload could carry, which stopped being true when the engine
      // gained `property` ops and part counts — both inflate `total` without
      // ever reaching the change list. Count what actually went in.
      `This payload carries ${carried} of the ${input.summary.total} changes the engine ` +
        `counted; ${omitted} were omitted to fit the context budget. The redline beside you ` +
        'is complete — say so rather than implying these are all the changes.',
    );
  }

  // ---------------------------------------------------------------------
  // Changes the engine COUNTS but this payload cannot carry.
  //
  // `summary.total` is the full count, and the change list is before/after
  // TEXT. Two kinds of change therefore never reach it. Unstated, the model
  // sees a total it cannot reconcile against the entries it was handed, and
  // the gap is exactly the sort of thing it will paper over with a guess.
  // ---------------------------------------------------------------------

  const propertyOps = input.ops.filter((op) => op.type === 'property').length;
  if (propertyOps > 0) {
    spots.push(
      `${propertyOps} change(s) alter a paragraph's style, numbering or level while leaving its ` +
        'words identical. They are counted in the totals and they are drawn in the redline, but ' +
        'the change list below cannot represent them: its entries are before and after TEXT, ' +
        'and here the text is the same on both sides. Do not call those paragraphs unchanged.',
    );
  }

  const { headers, footers, footnotes } = input.summary.partChanges;
  const outsideBody = headers + footers + footnotes;
  if (outsideBody > 0) {
    const where = [
      headers > 0 ? `${headers} in headers` : null,
      footers > 0 ? `${footers} in footers` : null,
      footnotes > 0 ? `${footnotes} in footnotes` : null,
    ].filter((bit): bit is string => bit !== null);
    spots.push(
      `${outsideBody} change(s) sit outside the document body (${where.join(', ')}). They are ` +
        'counted in the totals but are not in the change list below. ' +
        (headers + footers > 0
          ? 'Header and footer changes are not drawn in the redline either — it renders the ' +
            'body and the footnotes, nothing else. '
          : '') +
        'Do not assert what they say.',
    );
  }

  if (input.summary.styleDefinitionNote !== null) {
    spots.push(
      `The two documents' style definitions differ: ${input.summary.styleDefinitionNote}. That ` +
        'is a change to the style table itself rather than to any one paragraph, so no entry ' +
        'in the change list below carries it.',
    );
  }

  // Deliberately narrower than it used to read. "Formatting-only changes are
  // included as changes" was true until the engine gained `property` ops, and
  // is now false for exactly those — which the disclosure above covers. The
  // second sentence is rule 4 and still holds without qualification.
  spots.push(
    'Whitespace and punctuation changes are counted in the totals and included as changes. ' +
      'Nothing has been filtered out of the redline.',
  );

  return spots;
}
