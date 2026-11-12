import type { ParaNode } from '@shared/doc-model/types.js';
import type { JurisdictionProfile } from '@shared/config/types.js';
import type { SectionIndex, SectionKey, SectionKind, SectionNode } from './types.js';
import { normaliseNumber, disambiguateRoman, isAmbiguousRomanLetter } from './normalise.js';

/**
 * Section Index builder.
 *
 * Derives the logical legal structure (clauses/sub-clauses, schedules, parts,
 * recitals, annexes, appendices) from the Document Model's paragraphs by
 * combining named styles with `listString` auto-numbers and a running
 * structural context stack.
 *
 * Implements spec §4.3 (key composition, compound sub-clause keys, static
 * fallback) and §4.4 (number normalisation, roman/letter disambiguation).
 * EXTENDED to support compound keys (clause:5(b)(i)), static-numbered unit
 * fallback indexing, config-driven displayLabels, bookmarkNames field, and
 * the has/resolve/ordered convenience interface.
 */

interface RegionContext {
  kind: SectionKind;
  rawNumber: string;
  key: SectionKey;
}

/** Tracks the compound key fragment at each list level for clause ancestor stack. */
interface ClauseLevel {
  /** Full compound key suffix without the "clause:" prefix, e.g. "5", "5(b)", "5(b)(i)". */
  compoundKey: string;
  /** The normalised token at this level, e.g. "5", "(b)", "(i)". */
  token: string;
}

const HEADING_REGION_RE = /^\s*(Schedule|Annex|Annexure|Appendix|Part)\b\s*([A-Z0-9]+|[IVXLC]+)?/i;
const RECITAL_HEADING_RE = /^\s*(Recitals?|Background)\b/i;
const RECITAL_LETTER_RE = /^\s*\(([A-Z])\)/;

/**
 * Leading static-numbering regexes (spec §5.2.1) — used only for fallback indexing.
 *
 * Conservative to avoid phantom index entries from ordinary prose:
 *  - DOTTED decimal ("4.2", "4.2.1") may be followed by `.`, `)`, tab, or space —
 *    a paragraph starting "4.2 " is almost always a clause, "4.2 days" is rare.
 *  - BARE integer ("5") must be followed by `.`, `)`, or tab — NOT a plain space.
 *    This rejects prose like "5 days after…", "1976 Act…", "30 June 2026" while
 *    keeping real static clauses "5. Payment", "5) Payment".
 */
const STATIC_DECIMAL_DOTTED_RE = /^\s*(\d+(?:\.\d+)+)[.)\t ]/;
const STATIC_DECIMAL_BARE_RE = /^\s*(\d+)[.)\t]/;
const STATIC_LETTER_RE = /^\s*\(?([a-z])\)[ \t]/;
const STATIC_ROMAN_RE = /^\s*\(?(i{1,3}|iv|vi{0,3}|ix|xi{0,2}|xii)\)[ \t]/i;
const STATIC_UPPER_RE = /^\s*\(?([A-Z])\)[ \t]/;

function normaliseRegionKind(word: string): SectionKind {
  const w = word.toLowerCase();
  if (w === 'schedule') return 'schedule';
  if (w === 'annex' || w === 'annexure') return 'annex';
  if (w === 'appendix') return 'appendix';
  if (w === 'part') return 'part';
  return 'unknown';
}

/** Strip trailing dots from a decimal token, e.g. "4.2." → "4.2", "1." → "1". */
function cleanDecimal(raw: string): string {
  return raw.replace(/\.*$/, '').trim();
}

/**
 * Normalise a raw listString token into its canonical key fragment.
 * Returns `{ token, isDecimal, isLetterOrRoman }` for use in the stack.
 */
function parseListToken(listString: string): { token: string; isDecimal: boolean } {
  const s = listString.trim();

  // Decimal with possible trailing dot: "1.", "1.1", "4.2.3."
  const decMatch = /^(\d+(?:\.\d+)*)\.*$/.exec(s);
  if (decMatch && decMatch[1]) {
    return { token: cleanDecimal(decMatch[1]), isDecimal: true };
  }

  // Paren forms: "(b)", "(iv)", "(B)"
  const parenMatch = /^\(([^)]+)\)$/.exec(s);
  const inner = (parenMatch && parenMatch[1]) ? parenMatch[1] : s.replace(/[.)\]]+$/, '').replace(/^[([\]]+/, '');

  return { token: inner, isDecimal: false };
}

/**
 * Determine the configured region noun for display labels.
 * Falls back gracefully if the profile doesn't have xref config (shouldn't happen).
 */
function regionNoun(kind: SectionKind, profile: JurisdictionProfile): string {
  const units = profile.xref?.units;
  switch (kind) {
    case 'schedule': return units?.annexurePrimary ?? 'Schedule';
    case 'annex': return 'Annex';
    case 'appendix': return 'Appendix';
    case 'part': return 'Part';
    default: return kind;
  }
}

function buildDisplayLabel(
  kind: SectionKind,
  rawNumber: string,
  compoundKey: string,
  parentKey: SectionKey | undefined,
  profile: JurisdictionProfile,
): string {
  const clauseNoun = profile.xref?.units.primaryBody ?? 'Clause';

  switch (kind) {
    case 'clause':
      return `${clauseNoun} ${compoundKey}`;
    case 'recital':
      return `Recital (${rawNumber})`;
    case 'part':
      return `Part ${rawNumber}`;
    case 'annex':
      return `Annex ${rawNumber}`;
    case 'appendix':
      return `Appendix ${rawNumber}`;
    case 'schedule': {
      if (parentKey?.startsWith('schedule:')) {
        const scheduleNum = parentKey.slice('schedule:'.length);
        return `paragraph ${rawNumber} of ${regionNoun('schedule', profile)} ${scheduleNum}`;
      }
      return `${regionNoun('schedule', profile)} ${rawNumber}`;
    }
    default:
      return rawNumber;
  }
}

export function buildSectionIndex(
  paragraphs: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
): SectionIndex {
  const nodes: SectionNode[] = [];
  const byKey = new Map<SectionKey, SectionNode[]>();
  const byParaId = new Map<string, SectionNode>();

  // Structural region context (last open Schedule/Annex/Appendix/Part/Recital region).
  let region: RegionContext | null = null;
  let inRecitals = false;

  // Clause ancestor stack: level → ClauseLevel.
  // Used to build compound keys like clause:5(b)(i) from individual list tokens.
  const clauseStack = new Map<number, ClauseLevel>();

  // Per-level previous sibling token for roman/letter disambiguation.
  const prevSiblingAtLevel = new Map<number, string>();

  for (const para of paragraphs) {
    const text = para.text ?? '';
    const trimmed = text.trim();
    if (trimmed.length === 0) continue;

    const tocStyles = profile.xref?.styles.tocStyles ?? [];
    if (tocStyles.includes(para.styleName)) continue;

    // 1. Does this paragraph open a new structural region (Schedule/Annex/Appendix/Part)?
    const headingMatch = HEADING_REGION_RE.exec(trimmed);
    if (headingMatch && headingMatch[1]) {
      const kind = normaliseRegionKind(headingMatch[1]);
      const rawNumber = normaliseNumber(
        headingMatch[2] ?? '1',
        kind === 'part' ? 'part' : 'schedule-or-annex',
      );
      const key: SectionKey = `${kind}:${rawNumber}`;
      region = { kind, rawNumber, key };
      inRecitals = false;
      clauseStack.clear();
      prevSiblingAtLevel.clear();
      const displayLabel = buildDisplayLabel(kind, rawNumber, rawNumber, undefined, profile);
      const node = makeNode(para, key, kind, rawNumber, displayLabel, undefined, 0, false);
      pushNode(node, nodes, byKey, byParaId);
      continue;
    }

    // 2. Recital region?
    if (RECITAL_HEADING_RE.test(trimmed)) {
      inRecitals = true;
      region = null;
      clauseStack.clear();
      prevSiblingAtLevel.clear();
      continue;
    }
    const recitalLetter = RECITAL_LETTER_RE.exec(trimmed);
    if (inRecitals && recitalLetter && recitalLetter[1]) {
      const rawNumber = recitalLetter[1]; // Uppercase letter A, B, C...
      const key: SectionKey = `recital:${rawNumber}`;
      const displayLabel = `Recital (${rawNumber})`;
      const node = makeNode(para, key, 'recital', rawNumber, displayLabel, undefined, 0, false);
      pushNode(node, nodes, byKey, byParaId);
      continue;
    }

    // 3. Numbered list item → clause (body) or paragraph (inside a region).
    if (para.listString) {
      const level = para.listLevel ?? 0;
      const { token, isDecimal } = parseListToken(para.listString);
      if (token.length === 0) continue;

      if (region) {
        // paragraph-of-Schedule: key like "schedule:2/para:3"
        const normToken = normaliseNumber(token, 'clause-body');
        const key: SectionKey = `${region.key}/para:${normToken}`;
        const displayLabel = buildDisplayLabel(region.kind, normToken, normToken, region.key, profile);
        const node = makeNode(para, key, region.kind, normToken, displayLabel, region.key, level + 1, false);
        pushNode(node, nodes, byKey, byParaId);
      } else {
        inRecitals = false;

        // Clear clause stack above the current level.
        for (const l of [...clauseStack.keys()]) {
          if (l > level) {
            clauseStack.delete(l);
            prevSiblingAtLevel.delete(l);
          }
        }

        let normToken: string;
        let ambiguousRomanLetter = false;

        if (isDecimal) {
          normToken = normaliseNumber(token, 'clause-body');
        } else {
          // Letter or roman sub-clause: normalise to paren form.
          const prev = prevSiblingAtLevel.get(level) ?? null;
          const rawInner = token.toLowerCase();

          if (isAmbiguousRomanLetter(rawInner)) {
            const { certain } = disambiguateRoman(rawInner, prev);
            // Flag informational ambiguity (spec §4.4/§9). NOTE: no alternate key is
            // needed for "accept either reading" — a single ambiguous character
            // (i/v/x/l/c/d/m) normalises to the SAME literal token under both the
            // roman and letter readings (e.g. "(i)" → "(i)"), so the cited form and
            // the stored key always collapse identically. The flag is surfaced for
            // UI/diagnostics only; resolution is reading-agnostic by construction.
            ambiguousRomanLetter = !certain;
          }

          normToken = normaliseNumber(token, 'clause-sub');
        }

        // Build compound key using ancestor stack.
        let compoundKey: string;
        if (isDecimal) {
          // Decimal tokens are already compound (e.g. "4.2.1").
          compoundKey = normToken;
        } else {
          // Letter/roman: prefix with nearest ancestor's compound key.
          let parentCompound: string | undefined;
          for (let l = level - 1; l >= 0; l--) {
            const entry = clauseStack.get(l);
            if (entry) {
              parentCompound = entry.compoundKey;
              break;
            }
          }
          compoundKey = parentCompound ? parentCompound + normToken : normToken;
        }

        clauseStack.set(level, { compoundKey, token: normToken });
        prevSiblingAtLevel.set(level, normToken);

        const key: SectionKey = `clause:${compoundKey}`;

        // parentKey: for decimal tokens use the dotted parent; for sub-clause tokens use the parent level's key.
        let parentKey: SectionKey | undefined;
        if (isDecimal && compoundKey.includes('.')) {
          parentKey = `clause:${compoundKey.slice(0, compoundKey.lastIndexOf('.'))}`;
        } else if (!isDecimal) {
          for (let l = level - 1; l >= 0; l--) {
            const entry = clauseStack.get(l);
            if (entry) {
              parentKey = `clause:${entry.compoundKey}`;
              break;
            }
          }
        }

        const displayLabel = buildDisplayLabel('clause', normToken, compoundKey, parentKey, profile);
        const node = makeNode(para, key, 'clause', compoundKey, displayLabel, parentKey, level, false);
        if (ambiguousRomanLetter) node.ambiguousRomanLetter = true;
        pushNode(node, nodes, byKey, byParaId);
      }
      continue;
    }

    // 4. Static-numbered fallback (spec §4.3 step 2 / §5.2.4):
    //    When listItem is null but text has a leading number, still index the unit
    //    with isStaticNumbered=true so Feature A doesn't generate false Not Found.
    if (!para.listString && para.container === 'body') {
      const staticNode = tryIndexStaticPara(para, trimmed, region, inRecitals, profile);
      if (staticNode) {
        pushNode(staticNode, nodes, byKey, byParaId);
      }
    }
  }

  const indexObj: SectionIndex = {
    nodes,
    byKey,
    byParaId,
    ordered: nodes,
    has(key: SectionKey): boolean {
      return (byKey.get(key)?.length ?? 0) > 0;
    },
    resolve(key: SectionKey): SectionNode[] {
      return byKey.get(key) ?? [];
    },
  };

  return indexObj;
}

/**
 * Attempt to index a static-numbered paragraph (no listString).
 * Returns a SectionNode with isStaticNumbered=true, or null if no leading number detected.
 */
function tryIndexStaticPara(
  para: ParaNode,
  trimmed: string,
  region: RegionContext | null,
  inRecitals: boolean,
  profile: JurisdictionProfile,
): SectionNode | null {
  // Try dotted decimal first (highest confidence), then bare integer (stricter delimiter).
  const decMatch = STATIC_DECIMAL_DOTTED_RE.exec(trimmed) ?? STATIC_DECIMAL_BARE_RE.exec(trimmed);
  if (decMatch && decMatch[1]) {
    const raw = cleanDecimal(decMatch[1]);
    const key: SectionKey = region ? `${region.key}/para:${raw}` : `clause:${raw}`;
    const parentKey: SectionKey | undefined = region ? region.key
      : raw.includes('.') ? `clause:${raw.slice(0, raw.lastIndexOf('.'))}` : undefined;
    const displayLabel = buildDisplayLabel(
      region ? region.kind : 'clause',
      raw,
      raw,
      parentKey,
      profile,
    );
    return makeNode(para, key, region ? region.kind : 'clause', raw, displayLabel, parentKey, 0, true);
  }

  // Recital uppercase letter.
  const upperMatch = STATIC_UPPER_RE.exec(trimmed);
  if (inRecitals && upperMatch && upperMatch[1]) {
    const raw = upperMatch[1];
    const key: SectionKey = `recital:${raw}`;
    return makeNode(para, key, 'recital', raw, `Recital (${raw})`, undefined, 0, true);
  }

  // Lowercase letter sub-clause.
  const letterMatch = STATIC_LETTER_RE.exec(trimmed);
  if (!region && !inRecitals && letterMatch && letterMatch[1]) {
    const raw = `(${letterMatch[1]})`;
    const key: SectionKey = `clause:${raw}`;
    return makeNode(para, key, 'clause', raw, buildDisplayLabel('clause', raw, raw, undefined, profile), undefined, 0, true);
  }

  // Roman sub-clause.
  const romanMatch = STATIC_ROMAN_RE.exec(trimmed);
  if (!region && !inRecitals && romanMatch && romanMatch[1]) {
    const raw = `(${romanMatch[1].toLowerCase()})`;
    const key: SectionKey = `clause:${raw}`;
    return makeNode(para, key, 'clause', raw, buildDisplayLabel('clause', raw, raw, undefined, profile), undefined, 0, true);
  }

  return null;
}

function makeNode(
  para: ParaNode,
  key: SectionKey,
  kind: SectionKind,
  rawNumber: string,
  displayLabel: string,
  parentKey: SectionKey | undefined,
  level: number,
  isStaticNumbered: boolean,
): SectionNode {
  const node: SectionNode = {
    key,
    kind,
    rawNumber,
    displayLabel,
    level,
    paraId: para.id,
    isStaticNumbered,
    bookmarkNames: [],
    getRange: para.getRange,
  };
  if (parentKey !== undefined) node.parentKey = parentKey;
  return node;
}

function pushNode(
  node: SectionNode,
  nodes: SectionNode[],
  byKey: Map<SectionKey, SectionNode[]>,
  byParaId: Map<string, SectionNode>,
): void {
  nodes.push(node);
  const bucket = byKey.get(node.key) ?? [];
  bucket.push(node);
  byKey.set(node.key, bucket);
  byParaId.set(node.paraId, node);
}
