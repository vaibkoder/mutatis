/**
 * Feature A — reference scanner (spec §5.1.2, §6).
 *
 * Compiles the reference regex library from config (Clause 6) and scans body
 * ParaNode.text to produce ReferenceMatch[] with TextAnchors. TOC-style
 * paragraphs and non-body containers are excluded. Self-reference phrases
 * (config §6.7) are discarded. Range/list references are expanded.
 *
 * NOTE: This module is pure (no Office JS calls). The REF-field detection and
 * final classification live in resolve.ts + fields.ts (T5).
 */

import type { ParaNode } from '@shared/doc-model/types.js';
import type { JurisdictionProfile } from '@shared/config/types.js';
import type { TextAnchor } from '@shared/doc-model/anchor.js';
import type { SectionIndex } from '@shared/section-index/types.js';
import { normaliseNumber } from '@shared/section-index/normalise.js';
import type { ReferenceMatch, UnitType } from './types.js';

let _matchCounter = 0;
function nextMatchId(): string {
  return `xref-match-${++_matchCounter}`;
}

/** Reset match counter (for tests). */
export function resetMatchCounter(): void {
  _matchCounter = 0;
}

interface CompiledPatterns {
  /** Combined regex for a single scan pass. Groups documented below. */
  combined: RegExp;
  selfRefPhrases: Set<string>;
  tocStyles: Set<string>;
  rangeConnectors: RegExp;
  listConnectors: RegExp;
  inclusiveMarker: RegExp;
}

function escapeRe(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Build the combined pattern from profile config (spec §6).
 *
 * Capture groups (1-indexed):
 *   1 — clause/section body unit + number+sub: e.g. "Clause 4.2(b)"
 *   2 — schedule/exhibit + number: e.g. "Schedule 2", "Schedule A"
 *   3 — paragraph-of-schedule compound: "paragraph N of Schedule M"
 *   4 — annex/appendix/part: "Annex A", "Part II"
 *   5 — recital named: "Recital (B)", "Recital B"
 *   6 — plural clauses list/range head (followed by connector + more)
 *   7 — plural schedules list/range head
 *
 * For simplicity of implementation we use separate named pass regexes and iterate.
 */
function compilePatterns(profile: JurisdictionProfile): CompiledPatterns {
  const u = profile.xref.units;
  const r = profile.xref.ranges;

  // Number patterns
  const NUM = String.raw`\d+(?:\.\d+)*`;
  const SUB = String.raw`(?:\([a-z]+\)|\([ivxlcdm]+\))*`;
  const LETTER_OR_NUM = String.raw`(?:${NUM}|[A-Z])`;

  // Clause/Section aliases
  const bodyWords = [...new Set(u.bodyAliases)].map(escapeRe).join('|');
  // Schedule/Exhibit aliases
  const schedWords = [...new Set(u.annexureAliases)].map(escapeRe).join('|');
  // Other divisions
  const otherWords = [...u.otherDivisions].map(escapeRe).join('|');
  // Recital nouns
  const recitalWords = [...u.recitalNouns].map(escapeRe).join('|');

  const selfRefPhrases = new Set(
    profile.xref.selfReferenceTerms.map(s => s.toLowerCase()),
  );
  const tocStyles = new Set(profile.xref.styles.tocStyles);

  const connectorStr = [...r.connectors].map(escapeRe).join('|');
  const listConnStr = [...r.listConnectors].map(escapeRe).join('|');

  const rangeConnectors = new RegExp(`^\\s*(?:${connectorStr})\\s*`, 'i');
  const listConnectors = new RegExp(`^\\s*(?:${listConnStr})\\s*`, 'i');
  // Anchored, non-global: consumed at the tail right after a range end-number.
  const inclusiveMarker = new RegExp(`^\\s*${escapeRe(r.inclusiveMarker)}`, 'i');

  return {
    combined: /unused/,
    selfRefPhrases,
    tocStyles,
    rangeConnectors,
    listConnectors,
    inclusiveMarker,

    // Store component patterns for use in scanText()
    ..._buildComponentPatterns(bodyWords, schedWords, otherWords, recitalWords, NUM, SUB, LETTER_OR_NUM),
  } as unknown as CompiledPatterns & ReturnType<typeof _buildComponentPatterns>;
}

function _buildComponentPatterns(
  bodyWords: string,
  schedWords: string,
  otherWords: string,
  recitalWords: string,
  NUM: string,
  SUB: string,
  LETTER_OR_NUM: string,
) {
  const PLURAL_BODY = bodyWords.replace(/\|/g, 's|') + 's';
  const PLURAL_SCHED = schedWords.replace(/\|/g, 's|') + 's';

  // Trailing guard: the matched identifier must be a whole token, not the start
  // of a real word. Without this, a unit noun at a sentence boundary swallows the
  // next word's first letter — e.g. "…in whole or in part. assign…" matches as a
  // reference "part. a" (the case-insensitive [A-Z] catching the "a" of "assign").
  const END = String.raw`(?![A-Za-z])`;

  return {
    // Clause references: "Clause 4.2", "Clause 4.2(b)(i)", "Section 4.2"
    clauseRe: new RegExp(
      `\\b(?:${bodyWords})(?:\\.)?\\s+(${NUM}${SUB})${END}`,
      'gi',
    ),
    // Schedule references: "Schedule 2", "Schedule A", "Exhibit B"
    schedRe: new RegExp(
      `\\b(?:${schedWords})(?:\\.)?\\s+(${LETTER_OR_NUM})${END}`,
      'gi',
    ),
    // paragraph-of-Schedule: "paragraph 3 of Schedule 2"
    paraOfSchedRe: new RegExp(
      `\\bparagraph(?:s)?\\s+(${NUM}${SUB})\\s+of\\s+(?:${schedWords}|Annex(?:ure)?|Appendix)\\s+(${LETTER_OR_NUM})${END}`,
      'gi',
    ),
    // that Schedule / bare "paragraph N" in context — handled separately
    bareParaRe: new RegExp(
      `\\bparagraph\\s+(${NUM}${SUB})\\s+of\\s+(?:that|the)\\s+(?:Schedule|Annex(?:ure)?|Appendix)`,
      'gi',
    ),
    // Annex / Appendix / Part
    otherRe: new RegExp(
      `\\b(?:${otherWords})(?:\\.)?\\s+(${LETTER_OR_NUM}|[IVXLC]+)${END}`,
      'gi',
    ),
    // Recital references: "Recital (B)", "Recital B"
    recitalRe: new RegExp(
      `\\b(?:${recitalWords})\\s+\\(?([A-Z])\\)?${END}`,
      'gi',
    ),
    // Plural clause list/range: "Clauses 4 and 5", "Clauses 4 to 7"
    pluralClauseRe: new RegExp(
      `\\b(?:${PLURAL_BODY})\\s+(${NUM}${SUB})${END}`,
      'gi',
    ),
    // Plural schedule list/range: "Schedules 2 and 3", "Schedules 2 to 5 (inclusive)"
    pluralSchedRe: new RegExp(
      `\\b(?:${PLURAL_SCHED})\\s+(${LETTER_OR_NUM})${END}`,
      'gi',
    ),
  };
}

type ComponentPatterns = ReturnType<typeof _buildComponentPatterns>;

/**
 * Normalise a cited number string to a canonical key fragment.
 * Symmetric with the builder (spec §4.4).
 */
function normaliseRef(raw: string, unitType: UnitType): string {
  switch (unitType) {
    case 'recital':
      return normaliseNumber(raw, 'recital');
    case 'part':
      return normaliseNumber(raw, 'part');
    case 'schedule':
    case 'annex':
    case 'appendix':
      return normaliseNumber(raw, 'schedule-or-annex');
    default:
      return normaliseNumber(raw, 'auto');
  }
}

/**
 * Determine the unit type from the matched keyword.
 * "Section" maps to 'clause' unless isTopLevelDivision=true (kept as clause space).
 */
function unitTypeFromKeyword(keyword: string, _profile: JurisdictionProfile): UnitType {
  const k = keyword.toLowerCase().replace(/\.*$/, '');
  if (k === 'schedule' || k === 'sch' || k === 'exhibit' || k === 'exh') return 'schedule';
  if (k === 'annex' || k === 'annexure') return 'annex';
  if (k === 'appendix' || k === 'app') return 'appendix';
  if (k === 'part') return 'part';
  if (k === 'recital' || k === 'background') return 'recital';
  if (k === 'paragraph') return 'paragraph';
  return 'clause';
}

/** Build a normalised key from unit type and raw number. */
function buildKey(unitType: UnitType, normNumber: string, parentScheduleKey?: string): string {
  if (unitType === 'paragraph' && parentScheduleKey) {
    return `${parentScheduleKey}/para:${normNumber}`;
  }
  return `${unitType}:${normNumber}`;
}

/** Expand a range of numbers: "2 to 5" → [2,3,4,5]. Supports numeric and alpha. */
function expandRange(start: string, end: string): string[] {
  const startN = parseInt(start, 10);
  const endN = parseInt(end, 10);
  if (!isNaN(startN) && !isNaN(endN) && endN >= startN) {
    const result: string[] = [];
    for (let i = startN; i <= endN; i++) result.push(String(i));
    return result;
  }
  // Alphabetic: "A" to "D"
  const startC = start.charCodeAt(0);
  const endC = end.charCodeAt(0);
  if (!isNaN(startC) && !isNaN(endC) && endC >= startC) {
    const result: string[] = [];
    for (let i = startC; i <= endC; i++) result.push(String.fromCharCode(i));
    return result;
  }
  return [start, end];
}

/**
 * After a leading reference token, attempt to parse a list or range continuation.
 * e.g. "Schedules 2 and 3" or "Schedules 2 to 5 (inclusive)".
 * Returns expanded keys (including the head) or just [headKey] if no continuation.
 */
function parseListOrRange(
  text: string,
  afterMatchEnd: number,
  headKey: string,
  unitType: UnitType,
  patterns: CompiledPatterns & ComponentPatterns,
  _profile: JurisdictionProfile,
): { expandedKeys: string[]; totalLength: number } {
  const remaining = text.slice(afterMatchEnd);
  const connectors = patterns.rangeConnectors;
  const listConn = patterns.listConnectors;
  const inclusive = patterns.inclusiveMarker;

  // Try range first: "to/–/—/- NUM"
  const rangeMatch = connectors.exec(remaining);
  if (rangeMatch) {
    const afterConn = remaining.slice(rangeMatch[0].length);
    const numMatch = /^(\d+(?:\.\d+)*|[A-Z])/.exec(afterConn);
    if (numMatch && numMatch[1]) {
      // Extract head number from headKey
      const headNum = headKey.split(':').pop()!.split('/').pop()!;
      const endNum = normaliseRef(numMatch[1], unitType);
      const expanded = expandRange(headNum, endNum).map(n => buildKey(unitType, n));
      let total = rangeMatch[0].length + numMatch[0].length;
      // Consume a trailing "(inclusive)" marker if present (anchored to the tail
      // immediately after the end number; `inclusive` is a non-global anchored regex).
      const tail = afterConn.slice(numMatch[0].length);
      const incMatch = inclusive.exec(tail);
      if (incMatch) total += incMatch[0].length;
      return { expandedKeys: expanded, totalLength: total };
    }
  }

  // Try list: "and/,/& NUM" repeatedly
  const keys: string[] = [headKey];
  let offset = 0;
  for (;;) {
    const rest = remaining.slice(offset);
    const lMatch = listConn.exec(rest);
    if (!lMatch) break;
    const afterList = rest.slice(lMatch[0].length);
    const numMatch = /^(\d+(?:\.\d+)*|[A-Z])/.exec(afterList);
    if (!numMatch || !numMatch[1]) break;
    const normN = normaliseRef(numMatch[1], unitType);
    keys.push(buildKey(unitType, normN));
    offset += lMatch[0].length + numMatch[0].length;
  }

  return { expandedKeys: keys, totalLength: offset };
}

/**
 * Scan a single paragraph's text for cross-reference patterns.
 * Returns an array of raw matches (classification not yet resolved).
 */
function scanText(
  para: ParaNode,
  patterns: CompiledPatterns & ComponentPatterns,
  profile: JurisdictionProfile,
  _index: SectionIndex,
): ReferenceMatch[] {
  const text = para.text;
  const results: ReferenceMatch[] = [];
  const selfRef = patterns.selfRefPhrases;

  function addMatch(
    rawText: string,
    startOffset: number,
    unitType: UnitType,
    normalizedKey: string,
    expandedKeys: string[],
  ): void {
    const lower = rawText.toLowerCase();
    for (const phrase of selfRef) {
      if (lower.startsWith(phrase)) return;
    }
    const anchor: TextAnchor = {
      paragraphId: para.id,
      startOffset,
      endOffset: startOffset + rawText.length,
    };
    results.push({
      matchId: nextMatchId(),
      rawText,
      anchor,
      paragraphId: para.id,
      unitType,
      normalizedKey,
      expandedKeys: expandedKeys.length > 0 ? expandedKeys : [normalizedKey],
      classification: 'notFound',
      candidateTargets: [],
      isFieldBacked: false,
      staticBadge: true,
      severity: 'warning',
    });
  }

  // --- Clause references ---
  patterns.clauseRe.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = patterns.clauseRe.exec(text)) !== null) {
    if (!m[1]) continue;
    const normNum = normaliseRef(m[1], 'clause');
    const key = buildKey('clause', normNum);
    const { expandedKeys, totalLength } = parseListOrRange(
      text, m.index + m[0].length, key, 'clause', patterns, profile,
    );
    addMatch(text.slice(m.index, m.index + m[0].length + totalLength), m.index, 'clause', key, expandedKeys);
  }

  // --- Schedule references ---
  patterns.schedRe.lastIndex = 0;
  while ((m = patterns.schedRe.exec(text)) !== null) {
    if (!m[1]) continue;
    const normNum = normaliseRef(m[1], 'schedule');
    const key = buildKey('schedule', normNum);
    const { expandedKeys, totalLength } = parseListOrRange(
      text, m.index + m[0].length, key, 'schedule', patterns, profile,
    );
    addMatch(text.slice(m.index, m.index + m[0].length + totalLength), m.index, 'schedule', key, expandedKeys);
  }

  // --- paragraph-of-Schedule ---
  patterns.paraOfSchedRe.lastIndex = 0;
  while ((m = patterns.paraOfSchedRe.exec(text)) !== null) {
    if (!m[1] || !m[2]) continue;
    const normPara = normaliseRef(m[1], 'clause');
    const normSched = normaliseRef(m[2], 'schedule');
    const schedKey = buildKey('schedule', normSched);
    const key = `${schedKey}/para:${normPara}`;
    addMatch(m[0], m.index, 'paragraph', key, [key]);
  }

  // --- bare "paragraph N of that Schedule" ---
  patterns.bareParaRe.lastIndex = 0;
  while ((m = patterns.bareParaRe.exec(text)) !== null) {
    if (!m[1]) continue;
    const normPara = normaliseRef(m[1], 'clause');
    // We don't know the enclosing schedule at this layer; use a placeholder key
    // that resolve.ts will fill in via context.
    const key = `schedule:?/para:${normPara}`;
    addMatch(m[0], m.index, 'paragraph', key, [key]);
  }

  // --- Annex / Appendix / Part ---
  patterns.otherRe.lastIndex = 0;
  while ((m = patterns.otherRe.exec(text)) !== null) {
    if (!m[1]) continue;
    // Determine kind from the matched keyword
    const keyword = m[0].split(/\s+/)[0] ?? '';
    const unitType = unitTypeFromKeyword(keyword, profile);
    const normNum = normaliseRef(m[1], unitType);
    const key = buildKey(unitType, normNum);
    addMatch(m[0], m.index, unitType, key, [key]);
  }

  // --- Recital references ---
  patterns.recitalRe.lastIndex = 0;
  while ((m = patterns.recitalRe.exec(text)) !== null) {
    if (!m[1]) continue;
    const normNum = normaliseRef(m[1], 'recital');
    const key = buildKey('recital', normNum);
    addMatch(m[0], m.index, 'recital', key, [key]);
  }

  // --- Plural clause list/range: "Clauses 4 and 5", "Clauses 4 to 7" ---
  patterns.pluralClauseRe.lastIndex = 0;
  while ((m = patterns.pluralClauseRe.exec(text)) !== null) {
    if (!m[1]) continue;
    const normNum = normaliseRef(m[1], 'clause');
    const key = buildKey('clause', normNum);
    const { expandedKeys, totalLength } = parseListOrRange(
      text, m.index + m[0].length, key, 'clause', patterns, profile,
    );
    addMatch(text.slice(m.index, m.index + m[0].length + totalLength), m.index, 'clause', key, expandedKeys);
  }

  // --- Plural schedule list/range: "Schedules 2 and 3", "Schedules 2 to 5 (inclusive)" ---
  patterns.pluralSchedRe.lastIndex = 0;
  while ((m = patterns.pluralSchedRe.exec(text)) !== null) {
    if (!m[1]) continue;
    const normNum = normaliseRef(m[1], 'schedule');
    const key = buildKey('schedule', normNum);
    const { expandedKeys, totalLength } = parseListOrRange(
      text, m.index + m[0].length, key, 'schedule', patterns, profile,
    );
    addMatch(text.slice(m.index, m.index + m[0].length + totalLength), m.index, 'schedule', key, expandedKeys);
  }

  return results;
}

/**
 * Compiled patterns cache per profile id.
 */
const _compiledCache = new Map<string, CompiledPatterns & ComponentPatterns>();

function getPatterns(profile: JurisdictionProfile): CompiledPatterns & ComponentPatterns {
  const cached = _compiledCache.get(profile.id);
  if (cached) return cached;
  const p = compilePatterns(profile) as CompiledPatterns & ComponentPatterns;
  _compiledCache.set(profile.id, p);
  return p;
}

/**
 * Scan all body paragraphs for cross-reference patterns.
 * Returns raw ReferenceMatch[] with classification='notFound' and staticBadge=true
 * as defaults; resolve.ts sets the final classification and fields.ts sets isFieldBacked.
 */
export function scanForReferences(
  paras: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
  index: SectionIndex,
): ReferenceMatch[] {
  const patterns = getPatterns(profile);
  const tocStyles = patterns.tocStyles;
  const results: ReferenceMatch[] = [];

  for (const para of paras) {
    if (para.container !== 'body') continue;
    if (tocStyles.has(para.styleName)) continue;
    const matches = scanText(para, patterns, profile, index);
    results.push(...matches);
  }

  return results;
}
