/**
 * Feature 06 — Numbering upgrade (ADR-0023): sequence-anomaly surfacing.
 *
 * Pure. Given the ordered typed-number tokens of a candidate block, surface
 * duplicate and gap anomalies — each carrying BOTH readings (an accidental typo
 * OR a deliberately reserved number). The tool NEVER auto-resequences
 * (constraint 4); it only surfaces. The user retypes, ignores, or converts
 * (Step 1+2 dissolves a typo because the dynamic list recomputes).
 */

export type AnomalyKind = 'duplicate' | 'gap';

export interface SequenceAnomaly {
  kind: AnomalyKind;
  /** Index into the input array where the anomaly is observed. */
  index: number;
  /** The token at `index`, e.g. "1.5". */
  token: string;
  /** The token immediately before, e.g. "1.5" (duplicate) or "1.5" before "1.7" (gap). */
  previousToken: string;
  /** Both readings, surfaced verbatim — the tool does not choose between them. */
  readings: { typo: string; reserved: string };
}

/** Numeric value of the final segment of a dotted-decimal / limb token. */
function tailValue(token: string): number | null {
  const t = token.trim();

  // Dotted decimal: take the last segment, e.g. "1.5.2" → 2.
  const dotted = /^(\d+(?:\.\d+)*)$/.exec(t);
  if (dotted) {
    const segs = t.split('.');
    const last = segs[segs.length - 1];
    const n = Number(last);
    return Number.isFinite(n) ? n : null;
  }

  // Parenthesised lower-letter limb: (a) → 1, (b) → 2 …
  const letter = /^\(([a-z])\)$/.exec(t);
  if (letter && letter[1]) return letter[1].charCodeAt(0) - 96;

  // Parenthesised upper-letter limb: (A) → 1 …
  const upper = /^\(([A-Z])\)$/.exec(t);
  if (upper && upper[1]) return upper[1].charCodeAt(0) - 64;

  // Parenthesised lower-roman limb: (i),(ii),(iii),(iv)…
  const roman = /^\(([ivxl]+)\)$/i.exec(t);
  if (roman && roman[1]) return romanToInt(roman[1].toLowerCase());

  return null;
}

function romanToInt(s: string): number | null {
  const map: Record<string, number> = { i: 1, v: 5, x: 10, l: 50, c: 100 };
  let total = 0;
  for (let i = 0; i < s.length; i++) {
    const cur = map[s[i] ?? ''];
    const next = map[s[i + 1] ?? ''];
    if (cur === undefined) return null;
    total += next !== undefined && cur < next ? -cur : cur;
  }
  return total;
}

/** Same dotted-decimal parent prefix? ("1.5" vs "1.7" share parent "1"; "1.5" vs "2.1" do not). */
function sameGroup(a: string, b: string): boolean {
  const pa = a.trim().split('.').slice(0, -1).join('.');
  const pb = b.trim().split('.').slice(0, -1).join('.');
  return pa === pb;
}

/**
 * Detect duplicate and gap anomalies across an ordered list of typed-number
 * tokens (as they appear in the candidate block, top to bottom).
 *
 * Only adjacent comparisons within the same numbering group are made — a level
 * change (e.g. "1.1" → "(a)") is not an anomaly. Never auto-fixes.
 */
export function detectSequenceAnomalies(tokens: readonly string[]): SequenceAnomaly[] {
  const out: SequenceAnomaly[] = [];

  for (let i = 1; i < tokens.length; i++) {
    const prev = tokens[i - 1] ?? '';
    const cur = tokens[i] ?? '';

    const pv = tailValue(prev);
    const cv = tailValue(cur);
    if (pv === null || cv === null) continue;

    const dotted = prev.includes('.') || cur.includes('.');
    if (dotted && !sameGroup(prev, cur)) continue;

    if (cv === pv) {
      out.push({
        kind: 'duplicate',
        index: i,
        token: cur,
        previousToken: prev,
        readings: {
          typo: `"${cur}" repeats the previous number — likely a typo.`,
          reserved: `"${cur}" may intentionally reuse the number (e.g. a reserved or split clause).`,
        },
      });
    } else if (cv > pv + 1) {
      out.push({
        kind: 'gap',
        index: i,
        token: cur,
        previousToken: prev,
        readings: {
          typo: `"${cur}" skips ${cv - pv - 1} number(s) after "${prev}" — possibly a typo or deleted clause.`,
          reserved: `The gap after "${prev}" may be a deliberately reserved number.`,
        },
      });
    }
  }

  return out;
}
