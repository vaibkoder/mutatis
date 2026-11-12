/**
 * Tests for B5 — Insert logic: jurisdiction gate + XML escaping + CC tag helpers.
 *
 * Acceptance checks:
 *   B-T3: cross-jurisdiction → blocking confirm naming X vs Y, insert only after confirm.
 *   B-T4: same-jurisdiction → no warning; HTML insert; wrapped in CLAUSE CC with version tag.
 *   B-T5: preserve-tables → OOXML path used.
 *
 * Office JS is NOT available in jsdom; we test:
 *   - The jurisdiction-gate pure function (requiresJurisdictionWarning).
 *   - The human-readable label helpers (jurisdictionLabel, docJurisdictionLabel).
 *   - XML escaping / control-char sanitisation.
 *   - CC tag builders (buildClauseTag, buildClauseTitle).
 *   - normaliseQuotesAndEncoding.
 *
 * Live host verification required for:
 *   - The actual insertHtml / insertOoxml call.
 *   - Content-control wrapping of the inserted range.
 *   - Whether the cursor range after insertHtml encompasses exactly the inserted content.
 */

import { describe, it, expect } from 'vitest';
import {
  requiresJurisdictionWarning,
  jurisdictionLabel,
  docJurisdictionLabel,
  buildClauseTag,
  buildClauseTitle,
  xmlEscape,
} from '@features/content/library/insert.js';
import { parseClauseTag } from '@features/content/library/drift.js';
import {
  stripControlChars,
  normaliseQuotesAndEncoding,
} from '@features/content/library/insert.js';

// ---------------------------------------------------------------------------
// B-T3 / B-T4: jurisdiction gate (pure function)
// ---------------------------------------------------------------------------

describe('requiresJurisdictionWarning (B-T3 / B-T4)', () => {
  // Cross-jurisdiction cases → should warn
  it('returns true when clause=SG, doc=UK (cross-jurisdiction)', () => {
    expect(requiresJurisdictionWarning('SG', 'UK')).toBe(true);
  });

  it('returns true when clause=EW, doc=SG (cross-jurisdiction)', () => {
    expect(requiresJurisdictionWarning('EW', 'SG')).toBe(true);
  });

  it('returns true when clause=OTHER, doc=SG (OTHER is always warned)', () => {
    expect(requiresJurisdictionWarning('OTHER', 'SG')).toBe(true);
  });

  it('returns true when clause=OTHER, doc=UK (OTHER is always warned)', () => {
    expect(requiresJurisdictionWarning('OTHER', 'UK')).toBe(true);
  });

  // Same-jurisdiction cases → no warning
  it('returns false when clause=SG, doc=SG (same jurisdiction)', () => {
    expect(requiresJurisdictionWarning('SG', 'SG')).toBe(false);
  });

  it('returns false when clause=EW, doc=UK (EW maps to UK; same jurisdiction)', () => {
    expect(requiresJurisdictionWarning('EW', 'UK')).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// Jurisdiction label helpers (used in warning modal copy)
// ---------------------------------------------------------------------------

describe('jurisdictionLabel + docJurisdictionLabel', () => {
  it('jurisdictionLabel returns "Singapore" for SG', () => {
    expect(jurisdictionLabel('SG')).toBe('Singapore');
  });

  it('jurisdictionLabel returns "England & Wales" for EW', () => {
    expect(jurisdictionLabel('EW')).toBe('England & Wales');
  });

  it('jurisdictionLabel returns "Other (New York)" for OTHER with other', () => {
    expect(jurisdictionLabel('OTHER', 'New York')).toBe('Other (New York)');
  });

  it('jurisdictionLabel returns "Other / unspecified jurisdiction" for OTHER without other', () => {
    expect(jurisdictionLabel('OTHER')).toBe('Other / unspecified jurisdiction');
  });

  it('docJurisdictionLabel returns "England & Wales" for UK', () => {
    expect(docJurisdictionLabel('UK')).toBe('England & Wales');
  });

  it('docJurisdictionLabel returns "Singapore" for SG', () => {
    expect(docJurisdictionLabel('SG')).toBe('Singapore');
  });
});

// ---------------------------------------------------------------------------
// CC tag builders (B-T4: version tag format)
// ---------------------------------------------------------------------------

describe('buildClauseTag + buildClauseTitle', () => {
  it('buildClauseTag produces CLAUSE:<id>:v<version>', () => {
    const tag = buildClauseTag('clz-001', 3);
    expect(tag).toBe('CLAUSE:clz-001:v3');
  });

  it('buildClauseTitle produces "<title> (rev <lastReviewed>)"', () => {
    const title = buildClauseTitle('Limitation of Liability', '2026-02-14');
    expect(title).toBe('Limitation of Liability (rev 2026-02-14)');
  });

  it('buildClauseTag keeps the id verbatim and round-trips through parseClauseTag', () => {
    // .tag is a typed Office setter (plain string) — no XML escaping. The tag
    // must round-trip so drift detection can recover the clause id + version.
    const tag = buildClauseTag('clz-001', 7);
    const parsed = parseClauseTag(tag);
    expect(parsed).toEqual({ clauseId: 'clz-001', version: 7 });
  });

  it('buildClauseTitle does not XML-escape — special chars display verbatim', () => {
    // A clause title with an ampersand must show "Indemnity & Liability", not
    // "Indemnity &amp; Liability" (the .title setter is a plain string).
    const title = buildClauseTitle('Indemnity & Liability', '2026-01-01');
    expect(title).toBe('Indemnity & Liability (rev 2026-01-01)');
    expect(title).not.toContain('&amp;');
  });
});

// ---------------------------------------------------------------------------
// XML escaping (spec §B.8 — escape special XML chars)
// ---------------------------------------------------------------------------

describe('xmlEscape', () => {
  it('escapes & to &amp;', () => {
    expect(xmlEscape('a & b')).toBe('a &amp; b');
  });

  it('escapes < and > to &lt; and &gt;', () => {
    expect(xmlEscape('<tag>')).toBe('&lt;tag&gt;');
  });

  it('escapes " to &quot;', () => {
    expect(xmlEscape('"hello"')).toBe('&quot;hello&quot;');
  });

  it('escapes \' to &apos;', () => {
    expect(xmlEscape("it's")).toBe('it&apos;s');
  });

  it('does not alter plain ASCII text', () => {
    expect(xmlEscape('Hello World')).toBe('Hello World');
  });
});

// ---------------------------------------------------------------------------
// Control character rejection (spec §B.8)
// ---------------------------------------------------------------------------

describe('stripControlChars', () => {
  it('removes ASCII control chars below 0x09', () => {
    const input = 'Hello\x00World\x01';
    expect(stripControlChars(input)).toBe('HelloWorld');
  });

  it('preserves TAB (0x09), LF (0x0A), CR (0x0D)', () => {
    const input = 'a\tb\nc\rd';
    expect(stripControlChars(input)).toBe('a\tb\nc\rd');
  });

  it('removes 0x0B (VT) and 0x0C (FF)', () => {
    const input = 'a\x0Bb\x0Cc';
    expect(stripControlChars(input)).toBe('abc');
  });

  it('removes 0x7F (DEL)', () => {
    expect(stripControlChars('ab\x7Fcd')).toBe('abcd');
  });

  it('preserves printable ASCII and Unicode', () => {
    const input = 'Hello, Wörld! "clause" — \'test\'';
    expect(stripControlChars(input)).toBe(input);
  });
});

// ---------------------------------------------------------------------------
// Smart-quote normalisation (spec §B.8)
// ---------------------------------------------------------------------------

describe('normaliseQuotesAndEncoding', () => {
  it('normalises curly double quotes to straight', () => {
    expect(normaliseQuotesAndEncoding('“Hello”')).toBe('"Hello"');
  });

  it('normalises curly single quotes / apostrophes to straight', () => {
    expect(normaliseQuotesAndEncoding('it’s ‘nice’')).toBe("it's 'nice'");
  });

  it('normalises HTML entity curly quotes', () => {
    expect(normaliseQuotesAndEncoding('&ldquo;test&rdquo;')).toBe('"test"');
    expect(normaliseQuotesAndEncoding('&lsquo;test&rsquo;')).toBe("'test'");
  });

  it('replaces &ndash; and &mdash; with Unicode chars', () => {
    expect(normaliseQuotesAndEncoding('&ndash;')).toBe('–');
    expect(normaliseQuotesAndEncoding('&mdash;')).toBe('—');
  });

  it('leaves plain ASCII text unchanged', () => {
    const input = '<p>Hello "World"</p>';
    expect(normaliseQuotesAndEncoding(input)).toBe(input);
  });
});
