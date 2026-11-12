import type { ParaNode } from '@shared/doc-model/index.js';
import type { JurisdictionProfile } from '@shared/config/index.js';
import type { DefinitionsMap, DefinitionRecord, DefinitionSource, DefinitionShape, TextAnchor } from './types.js';
import { normaliseTerm, cleanTerm } from './normalise.js';

// ---------------------------------------------------------------------------
// Shared regex building blocks
// ---------------------------------------------------------------------------

const QUOTES = '"\'“”‘’«»';
const Q = `[${QUOTES}]`;
const NOT_Q = `[^${QUOTES}]`;

/** A verb is a "reference" verb (points elsewhere) when it speaks of a meaning given/defined elsewhere. */
const REFERENCE_VERB = /meaning|defined\s+in/i;

/** A clause/schedule reference within THIS document. */
const CLAUSE_REF = /\b(?:Clause|Schedule|paragraph|Annex|Annexure|Appendix|Part)\s+\d[\w.()]*/i;

/** A named external instrument → cross-document. */
const EXTERNAL_DOC =
  /\b(?:the\s+)?([A-Z][A-Za-z]+(?:\s+[A-Z][A-Za-z]+){0,4}\s+(?:Agreement|Deed|Document|Instrument|SPA|Facility|Lease|Loan|Indenture|Mortgage|Charge|Letter))\b/;

/** Local-redefinition scope opener (spec §8). */
const CONTEXT_SCOPE =
  /for the purposes of (?:this )?(?:Clause|Schedule|paragraph)\s+[\w.]+/i;

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** Build a verb alternation, longest-first so "means and includes" wins over "means". */
function verbAlternation(verbs: readonly string[]): string {
  return [...verbs]
    .sort((a, b) => b.length - a.length)
    .map(v => escapeRegExp(v).replace(/\\?\s+/g, '\\s+'))
    .join('|');
}

function rawQuoteKind(ch: string): DefinitionRecord['rawQuote'] {
  if (ch === '"' || ch === '“' || ch === '”' || ch === '«' || ch === '»') return 'double';
  if (ch === "'" || ch === '‘' || ch === '’') return 'single';
  return 'none';
}

// ---------------------------------------------------------------------------
// Merge helper (shared with tables + footnote passes)
// ---------------------------------------------------------------------------

/** Merge a definition record under normaliseTerm(rawTerm); keep the earliest canonical. */
export function mergeDefinition(map: DefinitionsMap, rawTerm: string, record: DefinitionRecord): void {
  const key = normaliseTerm(rawTerm);
  if (!key || !/[a-z0-9]/i.test(key)) return;
  const existing = map.get(key);
  if (existing) {
    existing.definitions.push(record);
  } else {
    map.set(key, { key, canonical: cleanTerm(rawTerm), definitions: [record] });
  }
}

// ---------------------------------------------------------------------------
// Verb classification (cross-ref vs cross-doc vs substantive)
// ---------------------------------------------------------------------------

interface VerbClass {
  verb: DefinitionRecord['verb'];
  isCrossRef: boolean;
  isCrossDoc: boolean;
  crossRefTarget: string | null;
}

function classifyVerb(verb: string, trailing: string): VerbClass {
  if (REFERENCE_VERB.test(verb)) {
    const ext = EXTERNAL_DOC.exec(trailing);
    if (ext) {
      return { verb: 'cross-document', isCrossRef: false, isCrossDoc: true, crossRefTarget: ext[1]?.trim() ?? null };
    }
    const ref = CLAUSE_REF.exec(trailing);
    return {
      verb: 'cross-reference',
      isCrossRef: true,
      isCrossDoc: false,
      crossRefTarget: ref ? ref[0].trim().replace(/[.,;:]+$/, '') : null,
    };
  }
  return { verb, isCrossRef: false, isCrossDoc: false, crossRefTarget: null };
}

function makeRecord(
  source: DefinitionSource,
  shape: DefinitionShape,
  paraId: string,
  start: number,
  end: number,
  rawQuote: DefinitionRecord['rawQuote'],
  cls: VerbClass,
  contextScope: string | null,
): DefinitionRecord {
  const anchor: TextAnchor = { paragraphId: paraId, startOffset: start, endOffset: end };
  return {
    source,
    verb: cls.verb,
    shape,
    isCrossRef: cls.isCrossRef,
    isCrossDoc: cls.isCrossDoc,
    crossRefTarget: cls.crossRefTarget,
    anchor,
    styling: { bold: false, italic: false },
    rawQuote,
    contextScope,
  };
}

// ---------------------------------------------------------------------------
// Per-paragraph scan: Shapes A / A′ / B / C
// ---------------------------------------------------------------------------

interface Span { start: number; end: number; }

function overlaps(spans: Span[], start: number, end: number): boolean {
  return spans.some(s => start < s.end && end > s.start);
}

/**
 * Scan one paragraph's text for definitions (all ADR-0013 shapes) and merge
 * them into the map. Quoted shapes are recorded first; an unquoted Shape C
 * candidate whose span is already covered by a quoted match is skipped.
 */
export function scanParagraphDefinitions(
  text: string,
  paraId: string,
  source: DefinitionSource,
  profile: JurisdictionProfile,
  map: DefinitionsMap,
): void {
  const dt = profile.definedTerms;
  const verbAlt = verbAlternation(dt.verbs);
  const determiners = new Set(dt.determiners.map(d => d.toLowerCase()));

  const contextScope = CONTEXT_SCOPE.exec(text)?.[0]?.trim() ?? null;
  const covered: Span[] = [];

  // -- Shape A: quoted term + verb -----------------------------------------
  const shapeA = new RegExp(`(${Q})(${NOT_Q}{1,120})${Q}\\s*[,(]?\\s*(${verbAlt})\\b`, 'gi');
  for (let m = shapeA.exec(text); m; m = shapeA.exec(text)) {
    const rawTerm = m[2] ?? '';
    if (!/[a-z0-9]/i.test(rawTerm)) continue;
    const termStart = m.index + 1;
    const termEnd = termStart + rawTerm.length;
    const verb = (m[3] ?? '').replace(/\s+/g, ' ').toLowerCase();
    const trailing = text.slice(m.index + m[0].length, m.index + m[0].length + 100);
    const cls = classifyVerb(verb, trailing);
    mergeDefinition(map, rawTerm,
      makeRecord(source, 'A', paraId, termStart, termEnd, rawQuoteKind(m[1] ?? ''), cls, contextScope));
    covered.push({ start: termStart, end: termEnd });
  }

  // -- Shape A′: quoted term + colon (ADR-0013) -----------------------------
  if (dt.recogniseColon) {
    const shapeAp = new RegExp(`(${Q})(${NOT_Q}{1,120})${Q}\\s*:\\s*\\S`, 'gi');
    for (let m = shapeAp.exec(text); m; m = shapeAp.exec(text)) {
      const rawTerm = m[2] ?? '';
      if (!/[a-z0-9]/i.test(rawTerm)) continue;
      const termStart = m.index + 1;
      const termEnd = termStart + rawTerm.length;
      if (overlaps(covered, termStart, termEnd)) continue;
      const cls: VerbClass = { verb: null, isCrossRef: false, isCrossDoc: false, crossRefTarget: null };
      mergeDefinition(map, rawTerm,
        makeRecord(source, 'A-prime', paraId, termStart, termEnd, rawQuoteKind(m[1] ?? ''), cls, contextScope));
      covered.push({ start: termStart, end: termEnd });
    }
  }

  // -- Shape B: inline parenthetical ----------------------------------------
  const shapeB = new RegExp(`\\(\\s*(?:the|this|each|a|an)?\\s*(?:a\\s+)?(${Q})(${NOT_Q}{1,120})${Q}\\s*\\)`, 'gi');
  for (let m = shapeB.exec(text); m; m = shapeB.exec(text)) {
    const rawTerm = m[2] ?? '';
    if (!/[a-z0-9]/i.test(rawTerm)) continue;
    const quoteIdx = m[0].indexOf(m[1] ?? '"');
    const termStart = m.index + quoteIdx + 1;
    const termEnd = termStart + rawTerm.length;
    if (overlaps(covered, termStart, termEnd)) continue;
    const cls: VerbClass = { verb: null, isCrossRef: false, isCrossDoc: false, crossRefTarget: null };
    mergeDefinition(map, rawTerm,
      makeRecord(source, 'B', paraId, termStart, termEnd, rawQuoteKind(m[1] ?? ''), cls, contextScope));
    covered.push({ start: termStart, end: termEnd });
  }

  // -- Shape C: unquoted candidate + mandatory verb (ADR-0013) --------------
  if (dt.recogniseUnquoted) {
    scanShapeC(text, paraId, source, verbAlt, determiners, contextScope, covered, map);
  }
}

/**
 * Shape C — an unquoted Title-Case / ALL-CAPS candidate immediately followed by
 * a definitional verb. The mandatory verb is the false-positive control; the
 * determiner + sentence-position guards keep "The Company means business." and
 * "This Clause means that …" quiet.
 */
function scanShapeC(
  text: string,
  paraId: string,
  source: DefinitionSource,
  verbAlt: string,
  determiners: Set<string>,
  contextScope: string | null,
  covered: Span[],
  map: DefinitionsMap,
): void {
  // Case-sensitive: the candidate must start with an uppercase letter.
  const cand = "[A-Z][A-Za-z0-9&'’\\-]*(?:\\s+[A-Z][A-Za-z0-9&'’\\-]*){0,5}";
  const cRe = new RegExp(`(${cand})\\s+(${verbAlt})\\b`, 'g');

  for (let m = cRe.exec(text); m; m = cRe.exec(text)) {
    let candidate = m[1] ?? '';
    let candStart = m.index;

    // Strip a leading determiner that the Title-Case run swallowed
    // (e.g. "The Company" → "Company"), recording that a determiner preceded it.
    let strippedLeading = false;
    const lead = /^(\S+)(\s+)([\s\S]+)$/.exec(candidate);
    if (lead && determiners.has((lead[1] ?? '').toLowerCase())) {
      strippedLeading = true;
      candStart += (lead[1] ?? '').length + (lead[2] ?? '').length;
      candidate = lead[3] ?? '';
    }
    if (!candidate || !/[A-Za-z]/.test(candidate)) continue;

    const candEnd = candStart + candidate.length;
    if (overlaps(covered, candStart, candEnd)) continue;

    const verb = (m[2] ?? '').replace(/\s+/g, ' ').toLowerCase();
    const isReference = REFERENCE_VERB.test(verb);
    const isShallMean = /^shall mean/.test(verb);
    const eitherPosition = isReference || isShallMean;

    // Determiner immediately before the candidate in the surrounding prose.
    const preTrim = text.slice(0, candStart).replace(/\s+$/, '');
    const preLastWord = (preTrim.match(/(\S+)$/)?.[1] ?? '').toLowerCase().replace(/[^a-z'’]/g, '');
    const determiner = strippedLeading || determiners.has(preLastWord);
    const sentenceInitial = !strippedLeading && (preTrim === '' || /[.!?:\n]$/.test(preTrim));

    if (eitherPosition) {
      if (determiner) continue;
    } else {
      // bare means / includes / refers to — sentence-initial only, no determiner
      if (!sentenceInitial || determiner) continue;
    }

    const trailing = text.slice(m.index + m[0].length, m.index + m[0].length + 100);
    const cls = classifyVerb(verb, trailing);
    mergeDefinition(map, candidate,
      makeRecord(source, 'C', paraId, candStart, candEnd, 'none', cls, contextScope));
    covered.push({ start: candStart, end: candEnd });
  }
}

// ---------------------------------------------------------------------------
// Map builder (prose + footnote passes)
// ---------------------------------------------------------------------------

/**
 * Build the DefinitionsMap from the prose pass (body paragraphs) and the
 * footnote pass (footnote bodies). Table definitions are merged separately.
 */
export function buildDefinitionsMap(
  bodyParas: ReadonlyArray<ParaNode>,
  footnoteParas: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
): DefinitionsMap {
  const map: DefinitionsMap = new Map();
  for (const p of bodyParas) {
    scanParagraphDefinitions(p.text, p.id, 'prose', profile, map);
  }
  for (const p of footnoteParas) {
    scanParagraphDefinitions(p.text, p.id, 'footnote', profile, map);
  }
  return map;
}
