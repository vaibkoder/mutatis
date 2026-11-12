/**
 * Envelope + briefing-pack assembly (feature 11 — spec §4.1, §4.7).
 *
 * Wraps the serialiser outputs in the provenance envelope, populates
 * `checks_run` + `blind_spots` (F6/F8), and — for the composite — assembles the
 * mechanical preamble summary and open-items list. No data structure of its own.
 */
import type { SectionIndex } from '@shared/section-index/index.js';
import type { DefinitionsMap, UsageIndex } from '@features/defined-terms/index.js';
import type { DiffOp } from '@features/compare/model.js';
import type {
  DateRecord, AmountPairRecord, PlaceholderHit,
} from '@features/factual-consistency/index.js';
import type { RoleRecord, PartyAnchor } from '@features/party-role/index.js';
import type { AiExportEnvelope, CheckId, OpenItem } from '../types.js';
import { makeCiter } from '../cite.js';
import { buildBlindSpots } from './blind-spots.js';
import { serialiseDefinedTerms } from '../serialise/defined-terms.js';
import { serialiseSectionOutline } from '../serialise/section-outline.js';
import { serialiseChangeSet } from '../serialise/change-set.js';
import { serialiseDealFacts } from '../serialise/deal-facts.js';
import { serialiseCastList } from '../serialise/cast-list.js';

/** Document-level context shared by every export. */
export interface ExportContext {
  /** ISO-8601 timestamp, stamped by the caller. */
  generated: string;
  document: string;
  jurisdiction: 'SG' | 'EW' | 'unknown';
  /** Section Index, or null when Cross-Reference has not run. */
  index: SectionIndex | null;
  /** All paragraph ids in document order (for citation). */
  orderedParaIds: readonly string[];
  /** Full text of a paragraph by id. */
  paraText: (paragraphId: string) => string;
}

export interface DefinedTermsData { map: DefinitionsMap; usage: UsageIndex; }
export interface ChangeSetData {
  ops: readonly DiffOp[];
  baseline: string;
  revised: string;
  origParaText?: (op: DiffOp) => string;
  revParaText?: (op: DiffOp) => string;
  moveTarget?: (op: DiffOp) => string;
  /** The op's locator; default {@link partLabel} of its part (Compare's header/footer/note labels, ADR-0047). */
  clauseFor?: (op: DiffOp) => string;
}
export interface DealFactsData {
  dates: readonly DateRecord[];
  amounts: readonly AmountPairRecord[];
  placeholders: readonly PlaceholderHit[];
}
export interface CastListData { roles: readonly RoleRecord[]; anchors: readonly PartyAnchor[]; }

/** All checker data the briefing pack may draw on (each optional — present when run). */
export interface BriefingData {
  definedTerms?: DefinedTermsData;
  sectionOutline?: SectionIndex;
  dealFacts?: DealFactsData;
  castList?: CastListData;
  /** Undefined-but-used terms (from the defined-terms findings) — open-items source. */
  undefinedTerms?: ReadonlyArray<{ raw: string; clause: string }>;
}

function baseEnvelope(ctx: ExportContext, checksRun: CheckId[]): AiExportEnvelope {
  return {
    _mutatis: '1.0',
    generated: ctx.generated,
    document: ctx.document,
    jurisdiction: ctx.jurisdiction,
    checks_run: checksRun,
    blind_spots: buildBlindSpots(checksRun),
  };
}

/** Friendly label for a Compare part id (e.g. "document" → "Body"). */
export function partLabel(partId: string): string {
  if (partId === 'document') return 'Body';
  if (partId === 'footnotes' || partId.startsWith('footnote')) return 'Footnotes';
  if (partId === 'endnotes' || partId.startsWith('endnote')) return 'Endnotes';
  if (partId.startsWith('header')) return 'Header';
  if (partId.startsWith('footer')) return 'Footer';
  if (partId.startsWith('tableCell')) return 'Table';
  if (partId === 'styles') return 'Styles';
  return partId;
}

// ---------------------------------------------------------------------------
// Single-section export (#1–#5)
// ---------------------------------------------------------------------------

export function buildSectionEnvelope(
  checkId: CheckId,
  data: DefinedTermsData | SectionIndex | ChangeSetData | DealFactsData | CastListData,
  ctx: ExportContext,
): AiExportEnvelope {
  const cite = makeCiter(ctx.index, ctx.orderedParaIds);
  const env = baseEnvelope(ctx, [checkId]);

  switch (checkId) {
    case 'defined-terms': {
      const d = data as DefinedTermsData;
      env.defined_terms = serialiseDefinedTerms(d.map, {
        usage: d.usage, index: ctx.index, cite, paraText: ctx.paraText,
      });
      break;
    }
    case 'cross-reference': {
      env.section_outline = serialiseSectionOutline(data as SectionIndex, { paraText: ctx.paraText });
      break;
    }
    case 'compare': {
      const d = data as ChangeSetData;
      env.change_set = serialiseChangeSet(d.ops, {
        baseline: d.baseline, revised: d.revised,
        // Compare runs without a Section Index, so `clause` degrades to the part
        // label ("Body"/"Footnotes") rather than a clause number — the best
        // locator available on this path, not a citation bug.
        clauseFor: d.clauseFor ?? (op => partLabel(op.partId)),
        ...(d.origParaText ? { origParaText: d.origParaText } : {}),
        ...(d.revParaText ? { revParaText: d.revParaText } : {}),
        ...(d.moveTarget ? { moveTarget: d.moveTarget } : {}),
      });
      break;
    }
    case 'factual-consistency': {
      const d = data as DealFactsData;
      env.deal_facts = serialiseDealFacts(d, { cite });
      break;
    }
    case 'party-role': {
      const d = data as CastListData;
      env.cast_list = serialiseCastList(d, { cite });
      break;
    }
  }
  return env;
}

// ---------------------------------------------------------------------------
// Composite briefing pack (#14) — #1 + #2 + #4 + #5 + preamble
// ---------------------------------------------------------------------------

export function buildBriefingEnvelope(data: BriefingData, ctx: ExportContext): AiExportEnvelope {
  const cite = makeCiter(ctx.index, ctx.orderedParaIds);
  const checksRun: CheckId[] = [];

  if (data.definedTerms) checksRun.push('defined-terms');
  if (data.sectionOutline) checksRun.push('cross-reference');
  if (data.dealFacts) checksRun.push('factual-consistency');
  if (data.castList) checksRun.push('party-role');

  const env = baseEnvelope(ctx, checksRun);

  if (data.definedTerms) {
    env.defined_terms = serialiseDefinedTerms(data.definedTerms.map, {
      usage: data.definedTerms.usage, index: ctx.index, cite, paraText: ctx.paraText,
    });
  }
  if (data.sectionOutline) {
    env.section_outline = serialiseSectionOutline(data.sectionOutline, { paraText: ctx.paraText });
  }
  if (data.dealFacts) env.deal_facts = serialiseDealFacts(data.dealFacts, { cite });
  if (data.castList) env.cast_list = serialiseCastList(data.castList, { cite });

  const openItems = buildOpenItems(data, cite);
  env.preamble = {
    summary: buildSummary(env, ctx.jurisdiction, openItems.length),
    open_items: openItems,
  };
  return env;
}

/** Mechanically-assembled one-sentence summary (not AI-generated, spec §4.7). */
function buildSummary(
  env: AiExportEnvelope,
  jurisdiction: ExportContext['jurisdiction'],
  nOpen: number,
): string {
  const juris = jurisdiction === 'SG' ? 'Singapore'
    : jurisdiction === 'EW' ? 'England & Wales' : '';

  const roles = env.cast_list?.parties.map(p => p.role) ?? [];
  const between = roles.length ? ` between ${joinList(roles)}` : '';

  const execDate = env.deal_facts?.dates.find(d => d.role === 'execution')?.raw;
  const dated = execDate ? `, dated ${execDate}` : '';

  const nTerms = env.defined_terms?.glossary.length ?? 0;
  const nClauses = countClauses(env);
  const nDates = env.deal_facts?.dates.length ?? 0;

  const lead = juris ? `A ${juris} agreement` : 'An agreement';
  return `${lead}${between}${dated}, containing ${nTerms} defined term${plural(nTerms)}, ` +
    `${nClauses} clause${plural(nClauses)}, ${nDates} deal date${plural(nDates)}, and ` +
    `${nOpen} open item${plural(nOpen)}.`;
}

function buildOpenItems(data: BriefingData, cite: ReturnType<typeof makeCiter>): OpenItem[] {
  const items: OpenItem[] = [];
  for (const p of data.dealFacts?.placeholders ?? []) {
    items.push({ type: 'placeholder', raw: p.rawText, clause: cite(p.anchor.paragraphId) });
  }
  for (const t of data.undefinedTerms ?? []) {
    items.push({ type: 'undefined-term', raw: t.raw, clause: t.clause });
  }
  return items;
}

function countClauses(env: AiExportEnvelope): number {
  let n = 0;
  const walk = (nodes: { children?: unknown[] }[] | undefined): void => {
    for (const node of nodes ?? []) {
      n += 1;
      walk(node.children as { children?: unknown[] }[] | undefined);
    }
  };
  walk(env.section_outline?.clauses);
  return n;
}

function plural(n: number): string { return n === 1 ? '' : 's'; }

function joinList(items: string[]): string {
  if (items.length <= 1) return items.join('');
  if (items.length === 2) return `${items[0]} and ${items[1]}`;
  return `${items.slice(0, -1).join(', ')} and ${items[items.length - 1]}`;
}
