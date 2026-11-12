/**
 * Resolver for Interactive Reference Lookup (feature 04).
 *
 * Maps a LookupClassification → LookupResult using only:
 *  - The shared Section Index (no new section-parsing)
 *  - The in-memory TermsDictionary (no new definition-detection)
 *  - The DocModel paragraph snapshot (for preview text and proximity)
 *
 * No Office JS calls — this is a pure in-memory operation.
 *
 * Key responsibilities (M1/M2/M4/M6):
 *  - Cross-ref resolution via SectionIndex.resolve(key)
 *  - Compound schedule-paragraph: container-scoped listString paragraph lookup (M2)
 *  - One-hop definition-by-cross-ref (M4)
 *  - Proximity heuristic for multi-hit ambiguity (M6, spec §8.6)
 *
 * ESCALATE — touches the shared Section Index contract (container-scoped
 * `listString` paragraph resolution) and the one-hop forwardRef chain.
 */

import type { SectionIndex, SectionNode } from '@shared/section-index/types.js';
import type { ParaNode } from '@shared/doc-model/index.js';
import type { JurisdictionProfile } from '@shared/config/index.js';
import { getLookupConfig } from '@shared/config/index.js';
import type {
  LookupClassification,
  LookupResult,
  LookupResolved,
  AmbiguousCandidate,
  TermsDictionary,
  DefinedTerm,
  RefDescriptor,
} from './types.js';
import { classify } from './classify.js';

// ── Key construction (Section Index key grammar) ─────────────────────────────

/** Build the Section Index key string from a RefDescriptor. */
function descriptorToKey(desc: RefDescriptor): string {
  switch (desc.kind) {
    case 'clause':      return `clause:${desc.primary}`;
    case 'schedule':    return `schedule:${desc.primary}`;
    case 'recital':     return `recital:${desc.primary}`;
    case 'annexure':    return `annex:${desc.primary}`;
    case 'schedule-paragraph':
      return `schedule:${desc.primary}/para:${desc.secondary ?? '1'}`;
    case 'para-only':
      // Bare "paragraph N" — key construction is context-dependent (handled below).
      return `clause:${desc.primary}`;
  }
}

// ── Preview text extraction ───────────────────────────────────────────────────

function previewForNode(
  node: SectionNode,
  allParas: ReadonlyArray<ParaNode>,
  maxChars: number,
): string {
  const para = allParas.find(p => p.id === node.paraId);
  if (!para?.text) return '';
  return para.text.length > maxChars
    ? para.text.slice(0, maxChars) + '…'
    : para.text;
}

// ── Container helpers (proximity heuristic, spec §8.6) ───────────────────────

/**
 * Determine the "root container" of a section node: the schedule key it belongs
 * to (e.g. "schedule:2"), or "body" if it is in the main body.
 */
function rootContainer(node: SectionNode): string {
  // Schedule paragraph: "schedule:2/para:3" → container is "schedule:2"
  const slashIdx = node.key.indexOf('/para:');
  if (slashIdx !== -1) return node.key.slice(0, slashIdx);

  // Node is itself a top-level schedule.
  if (node.kind === 'schedule' && !node.parentKey) return node.key;

  // Node has a parent schedule.
  if (node.parentKey?.startsWith('schedule:')) return node.parentKey;

  return 'body';
}

/**
 * Find the section node that the cursor is inside (the nearest preceding section
 * in document order). Used to determine the cursor's container for proximity.
 */
function cursorSection(
  cursorParaId: string,
  allParas: ReadonlyArray<ParaNode>,
  index: SectionIndex,
): SectionNode | undefined {
  const cursorPara = allParas.find(p => p.id === cursorParaId);
  if (!cursorPara) return undefined;

  let last: SectionNode | undefined;
  for (const node of index.ordered) {
    const np = allParas.find(p => p.id === node.paraId);
    if (!np) continue;
    if (np.index > cursorPara.index) break;
    last = node;
  }
  return last;
}

// ── Single-node result builder ────────────────────────────────────────────────

function makeResolved(
  node: SectionNode,
  allParas: ReadonlyArray<ParaNode>,
  maxChars: number,
): LookupResolved {
  return {
    type: 'resolved',
    label: node.displayLabel,
    previewText: previewForNode(node, allParas, maxChars),
    node,
  };
}

// ── Ambiguity resolver (proximity heuristic) ─────────────────────────────────

/**
 * Given multiple candidate SectionNodes, apply the spec §8.6 proximity heuristic:
 *  1. Same container as cursor → preferred
 *  2. Nearest by paragraph index
 *  3. Otherwise → ambiguous pick-list, best-guess index 0
 */
function resolveAmbiguousNodes(
  nodes: SectionNode[],
  allParas: ReadonlyArray<ParaNode>,
  maxChars: number,
  cursorParaId: string | undefined,
  index: SectionIndex,
): LookupResult {
  const candidates: AmbiguousCandidate[] = nodes.map(n => ({
    kind: 'section' as const,
    label: n.displayLabel,
    containerBadge: containerBadge(n),
    previewText: previewForNode(n, allParas, maxChars),
    node: n,
  }));

  if (candidates.length === 0) return { type: 'no-definition', phrase: '' };
  if (candidates.length === 1) return makeResolved(nodes[0]!, allParas, maxChars);

  // Find best guess by proximity.
  let bestGuessIndex = 0;
  if (cursorParaId) {
    const cursorNode = cursorSection(cursorParaId, allParas, index);
    const cursorContainer = cursorNode ? rootContainer(cursorNode) : 'body';
    const cursorIdx = allParas.findIndex(p => p.id === cursorParaId);

    // 1. Prefer same container.
    const sameContainer = nodes
      .map((n, i) => ({ n, i }))
      .filter(({ n }) => rootContainer(n) === cursorContainer);
    if (sameContainer.length === 1) {
      return makeResolved(sameContainer[0]!.n, allParas, maxChars);
    }
    if (sameContainer.length > 1) {
      // Within same container, prefer nearest.
      bestGuessIndex = nearestByIndex(sameContainer.map(x => x.n), allParas, cursorIdx);
      return {
        type: 'ambiguous',
        candidates: reorder(candidates, sameContainer[bestGuessIndex]?.i ?? 0),
        bestGuessIndex: 0,
      };
    }

    // 2. Different containers → nearest by paragraph index.
    bestGuessIndex = nearestByIndex(nodes, allParas, cursorIdx);
  }

  return {
    type: 'ambiguous',
    candidates: reorder(candidates, bestGuessIndex),
    bestGuessIndex: 0,
  };
}

/** The display badge for a section node ("Body" / "Schedule 2"). */
function containerBadge(node: SectionNode): string {
  const c = rootContainer(node);
  if (c === 'body') return 'Body';
  // "schedule:2" → "Schedule 2"
  const num = c.replace(/^schedule:/, '');
  return `Schedule ${num}`;
}

/** Return the index of the node whose para-index is closest to the cursor. */
function nearestByIndex(
  nodes: SectionNode[],
  allParas: ReadonlyArray<ParaNode>,
  cursorIdx: number,
): number {
  let best = 0;
  let bestDist = Infinity;
  for (let i = 0; i < nodes.length; i++) {
    const np = allParas.findIndex(p => p.id === nodes[i]!.paraId);
    const dist = Math.abs(np - cursorIdx);
    if (dist < bestDist) { bestDist = dist; best = i; }
  }
  return best;
}

/** Move `candidates[idx]` to index 0 (best guess expanded first). */
function reorder<T>(arr: T[], idx: number): T[] {
  if (idx === 0) return arr;
  const copy = [...arr];
  const [item] = copy.splice(idx, 1);
  if (item !== undefined) copy.unshift(item);
  return copy;
}

// ── Compound schedule-paragraph resolution (M2) ──────────────────────────────

/**
 * Resolve "paragraph N of Schedule M" (spec §5.3 / §5.2).
 *
 * Queries the shared Section Index with the compound key `schedule:M/para:N`
 * (container-scoped by design in the index builder — never matches a body clause).
 * This is the correct "no new section-parsing" approach: the index already
 * partitions hits by container (spec §3).
 */
function resolveSchedPara(
  desc: RefDescriptor,
  index: SectionIndex,
  allParas: ReadonlyArray<ParaNode>,
  maxChars: number,
): LookupResult {
  const scheduleNum = desc.primary;
  const paraNum = desc.secondary ?? '1';
  const phrase = `paragraph ${paraNum} of Schedule ${scheduleNum}`;

  // Direct index query — the builder already keys compound schedule paragraphs
  // as "schedule:M/para:N", container-scoped (section-index/builder.ts §207-211).
  const key = `schedule:${scheduleNum}/para:${paraNum}`;
  const hits = index.resolve(key);

  if (hits.length === 0) {
    return { type: 'no-definition', phrase };
  }
  if (hits.length === 1) {
    return makeResolved(hits[0]!, allParas, maxChars);
  }
  // Multiple hits (unusual but possible in a malformed doc) → ambiguous.
  return resolveAmbiguousNodes(hits, allParas, maxChars, undefined, index);
}

// ── Para-only context resolution (spec §5.3 step 4) ─────────────────────────

function resolveParaOnly(
  desc: RefDescriptor,
  index: SectionIndex,
  allParas: ReadonlyArray<ParaNode>,
  maxChars: number,
  cursorParaId: string | undefined,
): LookupResult {
  const paraNum = desc.primary;

  // Determine if cursor is inside a schedule.
  if (cursorParaId) {
    const cursor = cursorSection(cursorParaId, allParas, index);
    if (cursor) {
      const container = rootContainer(cursor);
      if (container.startsWith('schedule:')) {
        // Cursor is inside a schedule → try container-scoped paragraph resolution.
        const schedNum = container.replace(/^schedule:/, '');
        const schedResult = resolveSchedPara(
          { kind: 'schedule-paragraph', primary: schedNum, secondary: paraNum },
          index,
          allParas,
          maxChars,
        );
        if (schedResult.type !== 'no-definition') return schedResult;
      }
    }
  }

  // Not in a schedule context: try as a body clause.
  const clauseKey = `clause:${paraNum}`;
  const clauseHits = index.resolve(clauseKey);
  if (clauseHits.length === 0) return { type: 'no-definition', phrase: `paragraph ${paraNum}` };
  return resolveAmbiguousNodes(clauseHits, allParas, maxChars, cursorParaId, index);
}

// ── One-hop forwardRef follow (M4) ───────────────────────────────────────────

function followForwardRef(
  forwardRef: string,
  index: SectionIndex,
  dictionary: TermsDictionary,
  allParas: ReadonlyArray<ParaNode>,
  maxChars: number,
): LookupResolved | undefined {
  // Classify the raw forwardRef string as a cross-reference.
  const cls = classify(forwardRef, dictionary);
  if (cls.type !== 'CROSS_REF' && cls.type !== 'SCHED_PARA') return undefined;
  if (!cls.descriptor) return undefined;

  const key = descriptorToKey(cls.descriptor);
  const hits = index.resolve(key);
  if (hits.length !== 1) return undefined; // 0 or ambiguous → don't follow

  return makeResolved(hits[0]!, allParas, maxChars);
}

// ── Defined-term resolution ───────────────────────────────────────────────────

function resolveTerm(
  termKey: string,
  dictionary: TermsDictionary,
  index: SectionIndex,
  allParas: ReadonlyArray<ParaNode>,
  maxChars: number,
): LookupResult {
  const entries = dictionary.byKey.get(termKey);
  if (!entries || entries.length === 0) {
    return { type: 'no-definition', phrase: termKey };
  }

  // Multiple entries → term defined more than once → ambiguous.
  if (entries.length > 1) {
    const candidates: AmbiguousCandidate[] = entries.map(term => ({
      kind: 'term' as const,
      label: term.display,
      containerBadge: clauseBadgeForTerm(term, allParas, index),
      previewText: term.previewText,
      term,
    }));
    return { type: 'ambiguous', candidates, bestGuessIndex: 0 };
  }

  const term = entries[0]!;

  if (term.defKind === 'external') {
    return { type: 'definition', term };
  }

  if (term.defKind === 'by-cross-ref' && term.forwardRef) {
    // Follow one hop — never recurse.
    const resolvedClause = followForwardRef(
      term.forwardRef, index, dictionary, allParas, maxChars,
    );
    return { type: 'def-by-xref', term, ...(resolvedClause ? { resolvedClause } : {}) };
  }

  return { type: 'definition', term };
}

/** Best-effort container badge for a defined term (the section it was defined in). */
function clauseBadgeForTerm(
  term: DefinedTerm,
  allParas: ReadonlyArray<ParaNode>,
  index: SectionIndex,
): string {
  const node = index.byParaId.get(term.anchor.paragraphId);
  if (node) return node.displayLabel;
  // Walk backwards from the definition paragraph to find enclosing section.
  const defPara = allParas.find(p => p.id === term.anchor.paragraphId);
  if (!defPara) return 'Body';
  const enclosing = cursorSection(term.anchor.paragraphId, allParas, index);
  return enclosing?.displayLabel ?? 'Body';
}

// ── Main entry point ─────────────────────────────────────────────────────────

/**
 * Resolve a classification to a `LookupResult`.
 *
 * @param cls            Output from `classify()`.
 * @param index          The shared Section Index.
 * @param dictionary     The in-memory defined-terms dictionary.
 * @param allParas       Full DocModel paragraph snapshot (for text + proximity).
 * @param profile        Active jurisdiction profile (for preview limits).
 * @param cursorParaId   DocModel paragraph ID of the cursor position (optional).
 */
export function resolve(
  cls: LookupClassification,
  index: SectionIndex,
  dictionary: TermsDictionary,
  allParas: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
  cursorParaId?: string,
): LookupResult {
  const { previewMaxChars } = getLookupConfig(profile);

  switch (cls.type) {
    case 'NEITHER':
      return { type: 'idle' };

    case 'EXTERNAL':
      return { type: 'external', docName: cls.docName ?? '' };

    case 'SCHED_PARA': {
      if (!cls.descriptor) return { type: 'no-definition', phrase: cls.rawText };
      return resolveSchedPara(cls.descriptor, index, allParas, previewMaxChars);
    }

    case 'CROSS_REF': {
      if (!cls.descriptor) return { type: 'no-definition', phrase: cls.rawText };

      // Bare-paragraph needs context-dependent resolution.
      if (cls.descriptor.kind === 'para-only') {
        return resolveParaOnly(cls.descriptor, index, allParas, previewMaxChars, cursorParaId);
      }

      const key = descriptorToKey(cls.descriptor);
      const hits = index.resolve(key);
      if (hits.length === 0) return { type: 'no-definition', phrase: cls.rawText };
      if (hits.length === 1) return makeResolved(hits[0]!, allParas, previewMaxChars);
      return resolveAmbiguousNodes(hits, allParas, previewMaxChars, cursorParaId, index);
    }

    case 'DEFINED_TERM': {
      const termKey = cls.termKey ?? '';
      return resolveTerm(termKey, dictionary, index, allParas, previewMaxChars);
    }
  }
}
