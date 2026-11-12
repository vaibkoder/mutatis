import * as chrono from 'chrono-node';
import type { ParaNode } from '@shared/doc-model/index.js';
import type { JurisdictionProfile } from '@shared/config/index.js';
import type { Finding } from '@shared/findings/index.js';
import type { DateRecord, TextAnchor } from './types.js';

// ---------------------------------------------------------------------------
// Date extraction patterns
// ---------------------------------------------------------------------------

const MONTH_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
].join('|');

const ORDINAL_SUFFIX = '(?:st|nd|rd|th)?';

// DD Month YYYY  e.g. "6 June 2026", "13th December 2025"
const DD_MONTH_YYYY = new RegExp(
  `\\b(\\d{1,2})${ORDINAL_SUFFIX}\\s+(${MONTH_NAMES})\\s+(\\d{4})\\b`,
  'gi',
);

// Deed form: "this 6th day of June 2026" or "this ___ day of ____ 20__"
const ORDINAL_LONG = new RegExp(
  `\\bthis\\s+(\\d{1,2}|[A-Za-z]+)${ORDINAL_SUFFIX}\\s+day\\s+of\\s+(${MONTH_NAMES})\\s*,?\\s*(\\d{4}|\\d{2})\\b`,
  'gi',
);

// Template form: "this ___ day of ____ 20__" (incompletely filled deed date)
const ORDINAL_TEMPLATE = /\bthis\s+_+\s+day\s+of\s+_+\s+20_*/gi;

// Numeric: d/m/yyyy, d-m-yyyy, d.m.yyyy — always day-first (little-endian)
const NUMERIC_DATE = /\b(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})\b/g;

// ---------------------------------------------------------------------------
// Main scanner
// ---------------------------------------------------------------------------

export function scanDates(
  paras: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
): Finding[] {
  return analyseDates(paras, profile).findings;
}

/**
 * Same single pass as {@link scanDates}, but also returns the non-template
 * {@link DateRecord}s (with `roleHint`, `dayOfWeek` and `flaggedWeekend`
 * resolved) for the AI-export deal-facts serialiser (feature 11). The findings
 * are byte-identical to {@link scanDates} — no new scanning.
 */
export function analyseDates(
  paras: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
): { findings: Finding[]; records: DateRecord[] } {
  const findings: Finding[] = [];
  const collected: DateRecord[] = [];
  const { weekendDays, executionAnchorKeywords, effectiveDateKeywords,
          flagAllWeekendDates, roleHintWindow } = profile.dateLocale;

  for (let i = 0; i < paras.length; i++) {
    const para = paras[i]!;
    const prevText = i > 0 ? (paras[i - 1]?.text ?? '') : '';
    const contextText = prevText + ' ' + para.text;

    // Extract all date records from this paragraph.
    const records = extractDates(para.text, para.id);

    for (const rec of records) {
      // Template dates → skip weekend check; placeholder sweep handles them.
      if (rec.isTemplate) continue;

      // Retained for AI-export (deal facts). Real dates only — templates above.
      collected.push(rec);
      rec.flaggedWeekend = false;

      // Assign role hint from keyword context.
      const role = detectRoleHint(rec.rawText, rec.anchor.startOffset, contextText,
        executionAnchorKeywords, effectiveDateKeywords, roleHintWindow);
      rec.roleHint = role;

      // Century-ambiguous: flag as Informational, don't compute weekend.
      if (rec.centuryAmbiguous) {
        findings.push(makeFinding(rec, para, 'info',
          `Century-ambiguous date "${rec.rawText}" — cannot verify day of week`,
          'factual-consistency/century-ambiguous'));
        continue;
      }

      if (rec.parsed === null) continue;

      const day = rec.parsed.getDay() as 0 | 1 | 2 | 3 | 4 | 5 | 6;
      rec.dayOfWeek = day;
      const isWeekend = (weekendDays as readonly number[]).includes(day);

      if (!isWeekend) continue;

      // Effective/as-of dates: weekend not flagged as a Critical execution error.
      if (role === 'effective' || role === 'asof') continue;

      // Only flag execution/completion roles by default; flag all if config says so.
      if (role === 'execution' || role === 'completion') {
        const dayName = day === 0 ? 'Sunday' : 'Saturday';
        rec.flaggedWeekend = true;
        findings.push(makeFinding(rec, para, 'error',
          `Stated ${role} date "${rec.rawText}" falls on a ${dayName}`));
      } else if (flagAllWeekendDates) {
        const dayName = day === 0 ? 'Sunday' : 'Saturday';
        rec.flaggedWeekend = true;
        findings.push(makeFinding(rec, para, 'info',
          `Date "${rec.rawText}" falls on a ${dayName}`));
      }
    }
  }

  return { findings, records: collected };
}

// ---------------------------------------------------------------------------
// Date extraction
// ---------------------------------------------------------------------------

function extractDates(text: string, paraId: string): DateRecord[] {
  const records: DateRecord[] = [];
  const covered: Array<[number, number]> = [];

  // Template dates first (highest specificity — no further parsing needed).
  ORDINAL_TEMPLATE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = ORDINAL_TEMPLATE.exec(text)) !== null) {
    const start = m.index;
    const end = start + m[0].length;
    covered.push([start, end]);
    const anchor: TextAnchor = { paragraphId: paraId, startOffset: start, endOffset: end };
    records.push({
      anchor, rawText: m[0], parsed: null,
      format: 'template', centuryAmbiguous: false, isTemplate: true,
      roleHint: 'other', dayOfWeek: null,
    });
  }

  // Deed ordinal form: "this 6th day of June 2026"
  // Parse from capture groups (not the raw text) because chrono misreads "this N day of".
  ORDINAL_LONG.lastIndex = 0;
  while ((m = ORDINAL_LONG.exec(text)) !== null) {
    const start = m.index;
    const end = start + m[0].length;
    if (covered.some(([s, e]) => start >= s && end <= e)) continue;
    covered.push([start, end]);
    const yearStr = m[3] ?? '';
    const centuryAmbiguous = yearStr.length === 2;
    const rawText = m[0];
    // Reconstruct as "D Month YYYY" so chrono parses the day number correctly.
    const dayStr = m[1] ?? '';
    const monthStr = m[2] ?? '';
    const canonicalForm = `${dayStr} ${monthStr} ${yearStr}`;
    const parsed = centuryAmbiguous ? null : chronoParse(canonicalForm);
    const anchor: TextAnchor = { paragraphId: paraId, startOffset: start, endOffset: end };
    records.push({
      anchor, rawText, parsed,
      format: 'ordinal-long', centuryAmbiguous, isTemplate: false,
      roleHint: 'other',
      dayOfWeek: parsed ? (parsed.getDay() as 0 | 1 | 2 | 3 | 4 | 5 | 6) : null,
    });
  }

  // DD Month YYYY
  DD_MONTH_YYYY.lastIndex = 0;
  while ((m = DD_MONTH_YYYY.exec(text)) !== null) {
    const start = m.index;
    const end = start + m[0].length;
    if (covered.some(([s, e]) => start >= s && end <= e)) continue;
    covered.push([start, end]);
    const rawText = m[0];
    const parsed = chronoParse(rawText);
    const anchor: TextAnchor = { paragraphId: paraId, startOffset: start, endOffset: end };
    records.push({
      anchor, rawText, parsed,
      format: 'dd-month-yyyy', centuryAmbiguous: false, isTemplate: false,
      roleHint: 'other',
      dayOfWeek: parsed ? (parsed.getDay() as 0 | 1 | 2 | 3 | 4 | 5 | 6) : null,
    });
  }

  // Numeric dates (day-first): d/m/yyyy
  NUMERIC_DATE.lastIndex = 0;
  while ((m = NUMERIC_DATE.exec(text)) !== null) {
    const start = m.index;
    const end = start + m[0].length;
    if (covered.some(([s, e]) => start >= s && end <= e)) continue;
    const d = parseInt(m[1] ?? '0', 10);
    const mo = parseInt(m[2] ?? '0', 10);
    const yearStr = m[3] ?? '';
    const centuryAmbiguous = yearStr.length <= 2;
    // Validate: day 1-31, month 1-12 (rejects US-style month-first ambiguity
    // e.g. 13/6/2026 is clearly d/m since month 13 is invalid, confirming day-first).
    if (d < 1 || d > 31 || mo < 1 || mo > 12) continue;
    covered.push([start, end]);
    const rawText = m[0];
    const parsed = centuryAmbiguous ? null : chronoParse(rawText);
    const anchor: TextAnchor = { paragraphId: paraId, startOffset: start, endOffset: end };
    records.push({
      anchor, rawText, parsed,
      format: rawText.includes('/') ? 'numeric-slash' : 'numeric-dash',
      centuryAmbiguous, isTemplate: false,
      roleHint: 'other',
      dayOfWeek: parsed ? (parsed.getDay() as 0 | 1 | 2 | 3 | 4 | 5 | 6) : null,
    });
  }

  return records;
}

// ---------------------------------------------------------------------------
// Role hint detection
// ---------------------------------------------------------------------------

function detectRoleHint(
  rawDate: string,
  dateOffset: number,
  contextText: string,
  executionKeywords: readonly string[],
  effectiveKeywords: readonly string[],
  window: number,
): DateRecord['roleHint'] {
  // We look in a window around the date position in the combined context text.
  // `contextText` = prevParagraph + ' ' + currentParagraph; dateOffset is into
  // the current para, so offset in contextText = prevLen + 1 + dateOffset.
  // For simplicity use a case-insensitive search over a slice of contextText.
  const lower = contextText.toLowerCase();
  const dateIdx = lower.lastIndexOf(rawDate.toLowerCase());
  if (dateIdx === -1) return 'other';

  const start = Math.max(0, dateIdx - window);
  const end = Math.min(lower.length, dateIdx + rawDate.length + window);
  const ctx = lower.slice(start, end);

  for (const kw of effectiveKeywords) {
    if (ctx.includes(kw.toLowerCase())) return kw === 'as of' || kw === 'effective as of' ? 'asof' : 'effective';
  }

  for (const kw of executionKeywords) {
    const kwLower = kw.toLowerCase();
    if (ctx.includes(kwLower)) {
      if (kwLower.includes('completion')) return 'completion';
      return 'execution';
    }
  }

  return 'other';
}

// ---------------------------------------------------------------------------
// Chrono parsing (little-endian / day-first)
// ---------------------------------------------------------------------------

function chronoParse(rawText: string): Date | null {
  // chrono-node en.GB parses numeric dates day-first (little-endian).
  const results = chrono.en.GB.parse(rawText);
  if (results.length === 0) return null;
  return results[0]?.date() ?? null;
}

// ---------------------------------------------------------------------------
// Finding construction
// ---------------------------------------------------------------------------

function makeFinding(
  rec: DateRecord,
  para: ParaNode,
  severity: 'error' | 'info',
  title: string,
  category = 'factual-consistency/weekend-date',
): Finding {
  return {
    id: `${category}/${rec.anchor.paragraphId}-${rec.anchor.startOffset}`,
    category,
    featureId: 'factual-consistency',
    severity,
    title,
    location: { paraId: para.id, charStart: rec.anchor.startOffset, charEnd: rec.anchor.endOffset, label: para.id },
  };
}
