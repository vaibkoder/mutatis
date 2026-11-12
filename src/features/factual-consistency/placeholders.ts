import type { ParaNode } from '@shared/doc-model/index.js';
import type { JurisdictionProfile } from '@shared/config/index.js';
import type { Finding } from '@shared/findings/index.js';
import type { PlaceholderHit } from './types.js';
import type { TextAnchor } from './types.js';

/**
 * Scan all paragraphs for unfilled placeholder markers.
 *
 * Categories emitted: 'factual-consistency/placeholder'
 * Severity: all Informational ('info')
 * All patterns are config-driven (profile.placeholderVocab).
 */
export function scanPlaceholders(
  paras: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
): Finding[] {
  return analysePlaceholders(paras, profile).findings;
}

/**
 * Same single pass as {@link scanPlaceholders}, but also returns the
 * de-duplicated non-suppressed {@link PlaceholderHit}s for the AI-export
 * deal-facts serialiser (feature 11). Findings are byte-identical to
 * {@link scanPlaceholders} — no new scanning.
 */
export function analysePlaceholders(
  paras: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
): { findings: Finding[]; records: PlaceholderHit[] } {
  const findings: Finding[] = [];
  const collected: PlaceholderHit[] = [];
  const vocab = profile.placeholderVocab;

  // Build pattern list from config.
  const patterns = buildPatterns(vocab);

  const seen = new Set<string>();

  for (const para of paras) {
    const text = para.text;
    const hits = extractHits(text, para.id, patterns, vocab);

    for (const hit of hits) {
      if (hit.suppressed) continue;
      const key = `${hit.anchor.paragraphId}:${hit.anchor.startOffset}:${hit.rawText}`;
      if (seen.has(key)) continue;
      seen.add(key);

      collected.push(hit);
      const anchor = hit.anchor;
      findings.push(makeFinding(hit, anchor, para));
    }
  }

  return { findings, records: collected };
}

// ---------------------------------------------------------------------------
// Pattern registry
// ---------------------------------------------------------------------------

interface PatternEntry {
  re: RegExp;
  category: PlaceholderHit['category'];
  confidence: PlaceholderHit['confidence'];
  enabled: boolean;
}

function buildPatterns(vocab: JurisdictionProfile['placeholderVocab']): PatternEntry[] {
  // Build bullet char class from config.
  const bulletClass = vocab.bulletChars
    .filter(c => c.length > 0)
    .map(escapeRe)
    .join('');

  const patterns: PatternEntry[] = [
    // percent-bullet  [●]%  — check before filled-bullet so % suffix is captured
    {
      re: new RegExp(`\\[[${bulletClass}]\\]\\s*%`, 'g'),
      category: 'percent-bullet',
      confidence: 'high',
      enabled: true,
    },
    // filled-bullet  [●]  [•]  [ ]  [*]
    {
      re: new RegExp(`\\[[${bulletClass} *]\\]`, 'g'),
      category: 'filled-bullet',
      confidence: 'high',
      enabled: true,
    },
    // note-to-draft  [NTD: …]  [drafting note …]  — before keyword so NTD isn't caught twice
    {
      re: /\[(?:NTD|drafting note|note)\b[^\]]*\]/gi,
      category: 'note-to-draft',
      confidence: 'high',
      enabled: true,
    },
    // insert-directive  [INSERT …]  [insert …]
    {
      re: /\[(?:INSERT|insert)\b[^\]]*\]/g,
      category: 'insert-directive',
      confidence: 'high',
      enabled: true,
    },
    // keyword  [TBC]  [TBD]  [DATE]  [NAME]  [AMOUNT]  [PARTY]  [ADDRESS]
    {
      re: new RegExp(`\\[(${vocab.keywords.filter(k => !['NTD','drafting note','INSERT','note'].includes(k)).map(escapeRe).join('|')})\\]`, 'gi'),
      category: 'keyword',
      confidence: 'high',
      enabled: true,
    },
    // bare-brackets  [anything up to 80 chars]  — opt-in, default off
    {
      re: /\[[^\]]{1,80}\]/g,
      category: 'bare-brackets',
      confidence: 'medium',
      enabled: vocab.flagBareBrackets,
    },
  ];

  return patterns.filter(p => p.enabled);
}

// ---------------------------------------------------------------------------
// Extraction
// ---------------------------------------------------------------------------

function extractHits(
  text: string,
  paraId: string,
  patterns: PatternEntry[],
  vocab: JurisdictionProfile['placeholderVocab'],
): PlaceholderHit[] {
  const hits: PlaceholderHit[] = [];
  const coveredRanges: Array<[number, number]> = [];

  for (const { re, category, confidence } of patterns) {
    re.lastIndex = 0;
    let m: RegExpExecArray | null;
    while ((m = re.exec(text)) !== null) {
      const start = m.index;
      const end = start + m[0].length;

      // Skip if this range was already captured by a higher-priority pattern.
      if (coveredRanges.some(([s, e]) => start >= s && end <= e)) continue;

      const rawText = m[0];
      const anchor: TextAnchor = { paragraphId: paraId, startOffset: start, endOffset: end };
      const { suppressed, suppressReason } = classify(rawText, start, text, vocab);

      const hit: PlaceholderHit = suppressReason !== undefined
        ? { anchor, rawText, category, suppressed, suppressReason, confidence }
        : { anchor, rawText, category, suppressed, confidence };
      hits.push(hit);
      if (!suppressed) coveredRanges.push([start, end]);
    }
  }

  return hits;
}

// ---------------------------------------------------------------------------
// Suppression
// ---------------------------------------------------------------------------

function classify(
  rawText: string,
  start: number,
  text: string,
  vocab: JurisdictionProfile['placeholderVocab'],
): { suppressed: boolean; suppressReason?: PlaceholderHit['suppressReason'] } {
  // Whitelist check (full string match)
  if (vocab.whitelist.some(w => w === rawText)) {
    return { suppressed: true, suppressReason: 'whitelist' };
  }

  // Case citation: [2024] followed by a reporter abbreviation
  if (isCaseCitation(rawText, start, text, vocab.citationReporters)) {
    return { suppressed: true, suppressReason: 'case-citation' };
  }

  // Signature line proximity
  if (isNearSignature(start, text, vocab.signatureAnchors)) {
    return { suppressed: true, suppressReason: 'signature-line' };
  }

  return { suppressed: false };
}

/** [2019|20xx] immediately followed by a citation reporter abbreviation. */
function isCaseCitation(
  rawText: string,
  start: number,
  fullText: string,
  reporters: readonly string[],
): boolean {
  // Must look like [YYYY]
  if (!/^\[\s*(19|20)\d{2}\s*\]$/.test(rawText)) return false;

  // What follows the bracket?
  const after = fullText.slice(start + rawText.length).trimStart();
  return reporters.some(r => after.startsWith(r));
}

/** Whether the placeholder is within ~60 chars of a signature anchor keyword. */
function isNearSignature(
  start: number,
  text: string,
  anchors: readonly string[],
): boolean {
  const WINDOW = 60;
  const contextBefore = text.slice(Math.max(0, start - WINDOW), start);
  const lowerBefore = contextBefore.toLowerCase();
  return anchors.some(a => lowerBefore.includes(a.toLowerCase()));
}

// ---------------------------------------------------------------------------
// Finding construction
// ---------------------------------------------------------------------------

function makeFinding(
  hit: PlaceholderHit,
  anchor: TextAnchor,
  para: ParaNode,
): Finding {
  const categoryLabel = categoryDisplayName(hit.category);
  return {
    id: `factual-consistency/placeholder/${anchor.paragraphId}-${anchor.startOffset}-${encodeURIComponent(hit.rawText)}`,
    category: 'factual-consistency/placeholder',
    featureId: 'factual-consistency',
    severity: 'info',
    title: `Unfinished placeholder ${hit.rawText} (${categoryLabel})`,
    location: {
      paraId: para.id,
      charStart: anchor.startOffset,
      charEnd: anchor.endOffset,
      label: para.id,
    },
  };
}

function categoryDisplayName(cat: PlaceholderHit['category']): string {
  switch (cat) {
    case 'filled-bullet': return 'filled bullet';
    case 'percent-bullet': return 'percent bullet';
    case 'keyword': return 'keyword';
    case 'note-to-draft': return 'note to drafter';
    case 'insert-directive': return 'insert directive';
    case 'bare-brackets': return 'bare brackets';
  }
}

function escapeRe(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
