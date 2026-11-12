import type { ParaNode } from '@shared/doc-model/index.js';
import type { JurisdictionProfile } from '@shared/config/index.js';
import type { Finding } from '@shared/findings/index.js';
import type { DefinitionsMap, UsageIndex, TextAnchor } from './types.js';
import { normaliseTerm, resolveLookup, singularForms } from './normalise.js';
import { makeFinding, CATEGORIES, type Locator } from './locate.js';

// ---------------------------------------------------------------------------
// Shared helpers
// ---------------------------------------------------------------------------

function isSentenceInitialAt(text: string, idx: number): boolean {
  const pre = text.slice(0, idx).replace(/\s+$/, '');
  return pre === '' || /[.!?:\n]$/.test(pre);
}

/** The [start, end) character span of the sentence containing `idx`. */
function sentenceBounds(text: string, idx: number): [number, number] {
  let start = 0;
  for (let i = idx - 1; i >= 0; i--) {
    if (/[.!?:\n]/.test(text[i] ?? '')) { start = i + 1; break; }
  }
  let end = text.length;
  for (let i = idx; i < text.length; i++) {
    if (/[.!?\n]/.test(text[i] ?? '')) { end = i + 1; break; }
  }
  return [start, end];
}

/** Normalised allowlist set (party-roles, conventional, acronyms, statutory words, self-references). */
function buildAllowSet(profile: JurisdictionProfile): Set<string> {
  const ex = profile.definedTerms.exclusions;
  const s = new Set<string>();
  const add = (w: string): void => {
    const k = normaliseTerm(w);
    if (k) {
      s.add(k);
      s.add(k.replace(/^(?:the|this)\s+/, ''));
    }
  };
  ex.partyRoles.forEach(add);
  ex.conventional.forEach(add);
  ex.acronyms.forEach(add);
  ex.statutoryWords.forEach(add);
  ex.properNouns.forEach(add);
  // Self-reference nouns (this Agreement / this Deed / …) are never bespoke terms.
  profile.terminology.selfReferenceTerms.forEach(add);
  // Bare statutory nouns that appear capitalised on their own.
  add('Act');
  add('Regulations');
  return s;
}

/** Allowlist membership, with the same plural/possessive fallback as the definition lookup. */
function isAllowlisted(allow: Set<string>, key: string): boolean {
  if (allow.has(key) || allow.has(key.replace(/^(?:the|this)\s+/, ''))) return true;
  return singularForms(key).some(s => allow.has(s));
}

const CAND_WORD = "[A-Z][A-Za-z0-9&'’\\-]*";
const CAND_CONN = '(?:\\s+(?:of|and|&|the)\\s+|\\s+)';
const CAND_RE = new RegExp(`${CAND_WORD}(?:${CAND_CONN}${CAND_WORD})*`, 'g');

interface Candidate {
  raw: string;
  start: number;
  /** Strong evidence the capitalisation is intentional (mid-sentence, or after a stripped determiner). */
  midEvidence: boolean;
}

/** Harvest capitalised candidate terms from one paragraph (spec §5.4). */
function harvestCandidates(text: string, determiners: Set<string>): Candidate[] {
  const out: Candidate[] = [];
  CAND_RE.lastIndex = 0;
  for (let m = CAND_RE.exec(text); m; m = CAND_RE.exec(text)) {
    let raw = m[0];
    let start = m.index;
    let strippedDeterminer = false;

    const lead = /^(\S+)(\s+)([\s\S]+)$/.exec(raw);
    if (lead && determiners.has((lead[1] ?? '').toLowerCase())) {
      strippedDeterminer = true;
      start += (lead[1] ?? '').length + (lead[2] ?? '').length;
      raw = lead[3] ?? '';
    }

    const key = normaliseTerm(raw);
    if (key.length < 2 || !/[a-z]/i.test(key) || determiners.has(key)) continue;

    const sentenceInitial = !strippedDeterminer && isSentenceInitialAt(text, start);
    out.push({ raw, start, midEvidence: strippedDeterminer || !sentenceInitial });
  }
  return out;
}

/**
 * Classify the all-caps content of a candidate term for Sub-check 1.
 * An all-caps token = ≥2 letter chars, all uppercase (e.g. FEED, LIBOR, UAE).
 *   - 'unknown' : has an undefined, NON-allowlisted all-caps token (FEED, LIBOR).
 *                 Surfaced only if the candidate recurs (≥2 occurrences).
 *   - 'noise'   : all-caps tokens are present but every undefined one is an
 *                 allowlisted boilerplate/country word (THEREFORE, UAE) → suppress.
 *   - 'none'    : no undefined all-caps tokens (incl. when an all-caps token IS
 *                 defined — that takes precedence and the candidate is handled
 *                 normally).
 */
function classifyAllCaps(raw: string, map: DefinitionsMap, allow: Set<string>): 'unknown' | 'noise' | 'none' {
  let noise = false;
  for (const token of raw.split(/\s+/)) {
    const letters = token.replace(/[^A-Za-z]/g, '');
    if (letters.length >= 2 && letters === letters.toUpperCase()) {
      if (resolveLookup(map, token) !== null) continue;       // defined → precedence, treat as normal
      if (isAllowlisted(allow, normaliseTerm(token))) { noise = true; continue; }
      return 'unknown';
    }
  }
  return noise ? 'noise' : 'none';
}

// ---------------------------------------------------------------------------
// Sub-check 1 — Used-but-not-defined (Critical)
// ---------------------------------------------------------------------------

export function checkUsedNotDefined(
  map: DefinitionsMap,
  bodyParas: ReadonlyArray<ParaNode>,
  _tableParas: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
  locate: Locator,
): Finding[] {
  const ex = profile.definedTerms.exclusions;
  const allow = buildAllowSet(profile);
  const determiners = new Set(profile.definedTerms.determiners.map(d => d.toLowerCase()));

  // Definition term offsets per paragraph: a candidate in the same sentence as a
  // definition is definitional context (e.g. "Bank X" in '"Agent" means Bank X.')
  // and is not treated as usage that needs defining.
  const defOffsets = new Map<string, number[]>();
  for (const entry of map.values()) {
    for (const d of entry.definitions) {
      const arr = defOffsets.get(d.anchor.paragraphId) ?? [];
      arr.push(d.anchor.startOffset);
      defOffsets.set(d.anchor.paragraphId, arr);
    }
  }

  interface Agg { canonical: string; anchors: TextAnchor[]; midEvidence: boolean; }
  const agg = new Map<string, Agg>();

  const harvestFrom = (paras: ReadonlyArray<ParaNode>): void => {
    for (const p of paras) {
      const paraDefs = defOffsets.get(p.id) ?? [];
      for (const c of harvestCandidates(p.text, determiners)) {
        // Defined anywhere (incl. cross-ref / cross-doc) → not "undefined".
        if (resolveLookup(map, c.raw) !== null) continue;
        const key = normaliseTerm(c.raw);
        // Exclusion allowlist (plural/possessive-aware) + statutory-citation pattern.
        if (isAllowlisted(allow, key)) continue;
        if (ex.statutoryCitation.test(c.raw)) continue;

        // Occurrences sharing a sentence with a definition are definitional
        // context, not usage evidence.
        const [s, e] = sentenceBounds(p.text, c.start);
        const inDefContext = paraDefs.some(o => o >= s && o < e);
        const evidence = c.midEvidence && !inDefContext;

        const anchor: TextAnchor = {
          paragraphId: p.id, startOffset: c.start, endOffset: c.start + c.raw.length,
        };
        const existing = agg.get(key);
        if (existing) {
          existing.midEvidence = existing.midEvidence || evidence;
          existing.anchors.push(anchor);
        } else {
          agg.set(key, { canonical: c.raw, anchors: [anchor], midEvidence: evidence });
        }
      }
    }
  };

  // Candidates are harvested from body prose only; table/footnote definition
  // cells are definitions, not usage, and would otherwise raise false positives.
  harvestFrom(bodyParas);

  const findings: Finding[] = [];
  for (const a of agg.values()) {
    // Require ≥1 mid-sentence capitalised occurrence before flagging (spec §8).
    if (!a.midEvidence) continue;
    // All-caps handling (user rule): unknown all-caps surfaces only on recurrence;
    // allowlisted boilerplate/country all-caps (THEREFORE, UAE) stays suppressed.
    const cls = classifyAllCaps(a.canonical, map, allow);
    if (cls === 'noise') continue;
    if (cls === 'unknown' && a.anchors.length < 2) continue;

    const primary = a.anchors[0];
    if (!primary) continue;
    const title = `"${a.canonical}" is used but never defined`;
    const detail = `Define "${a.canonical}" or, if it is an ordinary word, remove the capitalisation.`;
    const occurrences = a.anchors.map(an => locate(an, a.canonical));
    findings.push(makeFinding(
      CATEGORIES.usedNotDefined, 'error', title,
      locate(primary, a.canonical), primary, detail, occurrences,
    ));
  }
  return findings;
}

// ---------------------------------------------------------------------------
// Sub-check 2 — Defined-but-not-used (Informational)
// ---------------------------------------------------------------------------

export function checkDefinedNotUsed(
  map: DefinitionsMap,
  usage: UsageIndex,
  locate: Locator,
): Finding[] {
  const findings: Finding[] = [];
  for (const entry of map.values()) {
    const uses = (usage.get(entry.key) ?? []).filter(u => !u.isDefinitionOccurrence);
    if (uses.length > 0) continue;
    const first = entry.definitions[0];
    if (!first) continue;
    const title = `"${entry.canonical}" is defined but never used`;
    const detail = `"${entry.canonical}" is defined but never used — consider removing the definition.`;
    findings.push(makeFinding(
      CATEGORIES.definedNotUsed, 'info', title,
      locate(first.anchor, entry.canonical), first.anchor, detail,
    ));
  }
  return findings;
}

// ---------------------------------------------------------------------------
// Sub-check 3 — Defined-multiple-times (Critical)
// ---------------------------------------------------------------------------

export function checkDefinedMultiple(
  map: DefinitionsMap,
  profile: JurisdictionProfile,
  locate: Locator,
): Finding[] {
  const flagLocal = profile.definedTerms.flagLocalRedefinitions;
  const findings: Finding[] = [];

  for (const entry of map.values()) {
    // Substantive records only: exclude pure cross-reference / cross-document.
    let substantive = entry.definitions.filter(d => !d.isCrossRef && !d.isCrossDoc);
    // Exclude differently-scoped local redefinitions unless configured to flag them.
    if (!flagLocal) {
      substantive = substantive.filter(d => d.contextScope === null);
    }
    if (substantive.length < 2) continue;

    const first = substantive[0];
    if (!first) continue;
    const locations = substantive.map(d => d.crossRefTarget ?? d.source).join(', ');
    const title = `"${entry.canonical}" is defined ${substantive.length} times`;
    const detail = `"${entry.canonical}" is defined ${substantive.length} times (${locations}). Keep one definition.`;
    findings.push(makeFinding(
      CATEGORIES.definedMultiple, 'error', title,
      locate(first.anchor, entry.canonical), first.anchor, detail,
    ));
  }
  return findings;
}

// ---------------------------------------------------------------------------
// Sub-check 4 — Capitalisation inconsistency (Informational) — T7
// ---------------------------------------------------------------------------

export function checkCapitalisation(
  map: DefinitionsMap,
  usage: UsageIndex,
  capMode: 'smart' | 'strict',
  _profile: JurisdictionProfile,
  locate: Locator,
  paras: ReadonlyArray<ParaNode>,
): Finding[] {
  const findings: Finding[] = [];
  for (const entry of map.values()) {
    const lowers = (usage.get(entry.key) ?? [])
      .filter(u => u.isLowercase && !u.isDefinitionOccurrence);

    const defParaId = entry.definitions[0]?.anchor.paragraphId;
    const sourcePara = defParaId ? paras.find(p => p.id === defParaId) : undefined;

    for (const u of lowers) {
      // smart (default): only flag a lowercase use immediately after a determiner.
      // strict: flag every lowercase use. The comparison is purely letter-case —
      // en-GB variants (Authorised/authorised) are a case hit, never a spelling flag.
      if (capMode === 'smart' && !u.precededByDeterminer) continue;
      const title = `"${u.surface}" is lowercase but "${entry.canonical}" is a defined term`;
      const detail = `"${u.surface}" appears in lowercase but is a defined term — capitalise it or confirm ordinary use.`;
      const finding = makeFinding(
        CATEGORIES.capitalisation, 'info', title,
        locate(u.anchor, entry.canonical), u.anchor, detail,
      );
      if (sourcePara) {
        finding.sourceSnippet = { label: entry.canonical, text: sourcePara.text };
      }
      findings.push(finding);
    }
  }
  return findings;
}
