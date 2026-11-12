/**
 * Single-phrase classifier for Interactive Reference Lookup (feature 04).
 *
 * Applies the spec §5.3 precedence order against one bounded text window.
 * Pure — no Office JS, no document reads.
 *
 * Precedence (first match wins):
 *  1. RX_SCHED_PARA  — compound "paragraph N of Schedule M"
 *  2. RX_EXTERNAL    — "Clause X of the Facility Agreement"
 *  3. RX_CLAUSE / RX_SCHEDULE / RX_ANNEXURE / RX_RECITAL
 *  4. RX_PARA_ONLY   — bare "paragraph N" (context-dependent resolution in resolver)
 *  5. Defined-term   — only if key exists in the dictionary
 *  6. NEITHER
 *
 * Vocabulary is SG/UK-fixed per the jurisdictional invariant (ADR-0002):
 * "Clause" not "Section", "Schedule" not "Exhibit", parenthesised recital letters.
 */

import { normaliseNumber } from '@shared/section-index/normalise.js';
import { resolveLookup } from '@features/defined-terms/normalise.js';
import type { TermsDictionary } from './types.js';
import type { LookupClassification } from './types.js';

// ── Numbering token (spec §5.3) ───────────────────────────────────────────────
// Captures: 14, 14.2, 14.2(a), 14.2(a)(iii), 4(1), 4(1)(a)
const NUM = String.raw`\d+(?:\.\d+)*(?:\([a-z0-9ivxlcdm]+\))*`;

// ── Regexes (spec §5.3, exact order) ─────────────────────────────────────────

/** 1. Compound Schedule paragraph — MUST be first. */
const RX_SCHED_PARA = new RegExp(
  String.raw`\b(?:para(?:graph)?s?\.?)\s+(${NUM})(?:\s+to\s+${NUM})?\s+of\s+(?:Sch(?:edule|\.)?)\s+(${NUM}|[IVXLC]+)\b`,
  'i',
);

/**
 * 2. External reference — "… of the Facility Agreement". Tested before plain refs.
 * Uses .*? (not [^.]*?) because clause numbers like "4.2" contain dots;
 * the spec's [^.]*? would fail on "Clause 4.2 of the Facility Agreement".
 */
const RX_EXTERNAL = new RegExp(
  String.raw`\b(?:Clause|Schedule|paragraph|Annexures?|Appendix)\b.*?\bof\s+the\s+((?:[A-Z][A-Za-z]+\s+){0,4}(?:[A-Z][A-Za-z]+\s+)?(?:Agreement|Deed|Letter|Schedule|Lease|Charge|Guarantee))\b`,
  'i',
);

/** 3a. Plain clause reference. */
const RX_CLAUSE = new RegExp(String.raw`\bClause\s+(${NUM})\b`, 'i');

/** 3b. Schedule reference. */
const RX_SCHEDULE = new RegExp(String.raw`\bSchedule\s+(${NUM}|[IVXLC]+)\b`, 'i');

/** 3c. Annexure reference. */
const RX_ANNEXURE = new RegExp(String.raw`\bAnnexures?\s+(${NUM}|[A-Z])\b`, 'i');

/** 3d. Recital reference — "Recital (B)" or "Recital B". */
const RX_RECITAL = new RegExp(String.raw`\bRecital\s+\(?([A-Z])\)?\b`, 'i');

/** 4. Bare paragraph — context-dependent (resolver checks cursor container). */
const RX_PARA_ONLY = new RegExp(String.raw`\bpara(?:graph)?s?\.?\s+(${NUM})\b`, 'i');

// ── Classifier ────────────────────────────────────────────────────────────────

/**
 * Find the defined term whose surface form spans `offset` within `text`.
 *
 * Used for the bare-cursor case (spec §4.2 / T3): the caret sits *inside* a
 * defined term rather than the whole term being selected, so the candidate is
 * the enclosing sentence window and we must locate the specific term under the
 * cursor. Scans with the dictionary's pre-built longest-match-first alternation
 * and returns the key of the first term whose [start, end] span contains the
 * cursor — so two terms in one window resolve to the one the caret is in, not
 * merely the first in the paragraph.
 *
 * @returns the resolved dictionary key, or null when the cursor is not inside a term.
 */
function findTermAtOffset(
  text: string,
  offset: number,
  dictionary: TermsDictionary,
): string | null {
  const re = dictionary.termMatcher;
  if (!re) return null;
  re.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    const start = m.index;
    const end = start + m[0].length;
    // Inclusive bounds: clicking at either edge of the term still resolves it.
    if (offset >= start && offset <= end) {
      const key = resolveLookup(dictionary.byKey, m[0]);
      if (key !== null) return key;
    }
    // Guard against zero-length matches looping forever (defensive; alternation
    // never produces them, but lastIndex must advance).
    if (re.lastIndex === m.index) re.lastIndex++;
  }
  return null;
}

/**
 * Classify a candidate text string into a LookupClassification.
 *
 * @param text         The candidate phrase (selection text or sentence window).
 * @param dictionary   In-memory defined-terms dictionary (for step 5 check).
 * @param cursorOffset For a bare cursor, the caret's character offset within
 *                     `text`. When set, the defined-term step (5) also matches a
 *                     term the cursor sits *inside* — not just a whole-string
 *                     selection. Omitted for explicit selections.
 */
export function classify(
  text: string,
  dictionary: TermsDictionary,
  cursorOffset?: number,
): LookupClassification {
  const raw = text.trim();

  // 1. Compound Schedule paragraph.
  let m = RX_SCHED_PARA.exec(raw);
  if (m) {
    const para = m[1]!;
    const sched = m[2]!;
    return {
      type: 'SCHED_PARA',
      descriptor: {
        kind: 'schedule-paragraph',
        primary: normaliseNumber(sched, 'schedule-or-annex'),
        secondary: normaliseNumber(para, 'auto'),
      },
      rawText: raw,
    };
  }

  // 2. External reference.
  m = RX_EXTERNAL.exec(raw);
  if (m) {
    return {
      type: 'EXTERNAL',
      docName: m[1]!.trim(),
      rawText: raw,
    };
  }

  // 3a. Plain clause.
  m = RX_CLAUSE.exec(raw);
  if (m) {
    return {
      type: 'CROSS_REF',
      descriptor: {
        kind: 'clause',
        primary: normaliseNumber(m[1]!, 'auto'),
      },
      rawText: raw,
    };
  }

  // 3b. Schedule.
  m = RX_SCHEDULE.exec(raw);
  if (m) {
    return {
      type: 'CROSS_REF',
      descriptor: {
        kind: 'schedule',
        primary: normaliseNumber(m[1]!, 'schedule-or-annex'),
      },
      rawText: raw,
    };
  }

  // 3c. Annexure.
  m = RX_ANNEXURE.exec(raw);
  if (m) {
    return {
      type: 'CROSS_REF',
      descriptor: {
        kind: 'annexure',
        primary: normaliseNumber(m[1]!, 'schedule-or-annex'),
      },
      rawText: raw,
    };
  }

  // 3d. Recital.
  m = RX_RECITAL.exec(raw);
  if (m) {
    return {
      type: 'CROSS_REF',
      descriptor: {
        kind: 'recital',
        primary: normaliseNumber(m[1]!, 'recital'),
      },
      rawText: raw,
    };
  }

  // 4. Bare paragraph (context-dependent — resolver handles container scoping).
  m = RX_PARA_ONLY.exec(raw);
  if (m) {
    return {
      type: 'CROSS_REF',
      descriptor: {
        kind: 'para-only',
        primary: normaliseNumber(m[1]!, 'auto'),
      },
      rawText: raw,
    };
  }

  // 5. Defined-term test — only if key already exists in the dictionary.
  //   (a) Exact whole-candidate match — the user selected the term itself.
  //       resolveLookup checks exact key + naive singular forms; never invents a key.
  const resolved = resolveLookup(dictionary.byKey, raw);
  if (resolved !== null) {
    return {
      type: 'DEFINED_TERM',
      termKey: resolved,
      rawText: raw,
    };
  }
  //   (b) Bare cursor inside a term — the caret sits within a defined term in
  //       the sentence window, but the whole window is not itself the term
  //       (spec §4.2 / T3). Locate the term under the cursor offset.
  if (cursorOffset !== undefined) {
    const atCursor = findTermAtOffset(raw, cursorOffset, dictionary);
    if (atCursor !== null) {
      return {
        type: 'DEFINED_TERM',
        termKey: atCursor,
        rawText: raw,
      };
    }
  }

  // 6. Nothing matched.
  return { type: 'NEITHER', rawText: raw };
}
