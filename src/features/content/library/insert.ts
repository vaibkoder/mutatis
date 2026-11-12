/**
 * Precedent / Clause Library — click-to-insert logic.
 *
 * Spec: spec/features/07-toc-precedent-templates.md §B.5.3
 *
 * CRITICAL SAFETY GATE: The jurisdiction warning is MANDATORY and NON-SUPPRESSIBLE
 * for any cross-jurisdiction insert (entry.jurisdiction ≠ doc jurisdiction OR
 * entry.jurisdiction === 'OTHER'). The insert MUST NOT proceed without explicit
 * user confirmation when this gate fires. This is a core risk-management invariant.
 *
 * This module contains:
 *   - XML escaping + control-char rejection (shared utility, used in both HTML and OOXML paths)
 *   - Jurisdiction gate decision function (pure, testable)
 *   - Insert logic (Office JS — requires a Word.RequestContext)
 *   - Content-control wrapping (tracking CC: tag=CLAUSE:<id>:v<version>)
 *
 * Office JS notes (spec/ARCHITECTURE.md §Office JS Constraints Reference):
 *   - insertHtml(html, 'Replace') maps to the document's named styles; absent styles are substituted.
 *   - insertOoxml(ooxml, 'Replace') preserves table fidelity but may introduce style conflicts.
 *   - Content controls: set .tag and .title on the returned proxy BEFORE ctx.sync().
 *   - The content control wraps the INSERTED range — after insertHtml/insertOoxml the returned range
 *     is the new content; we then call range.insertContentControl() on that range.
 *   - Content-control-in-table-cell corruption risk: for table-bearing clauses the OOXML path
 *     is preferred (spec §B.8); HTML path is default for portability.
 *
 * DONE_WITH_CONCERNS: Range returned by insertHtml/insertOoxml is NOT guaranteed in all
 * Office JS versions to be the newly inserted content; the approach of calling
 * insertContentControl() on the cursor range after the insert is used instead (see below).
 * Live host verification required to confirm the CC wraps precisely the inserted content.
 */

import type { ClauseEntry } from './types.js';
import { wordRun } from '@shared/office/index.js';
import { getStoredJurisdiction } from '@shared/config/index.js';

// ---------------------------------------------------------------------------
// XML escaping / control-character hardening
// ---------------------------------------------------------------------------

/** Characters that must be XML-escaped when appearing in OOXML content. */
const XML_ESCAPE_MAP: Readonly<Record<string, string>> = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&apos;',
};

/**
 * XML-escape a plain-text string for safe embedding in OOXML attributes.
 * Used for tag and title values.
 */
export function xmlEscape(value: string): string {
  return value.replace(/[&<>"']/g, ch => XML_ESCAPE_MAP[ch] ?? ch);
}

/**
 * Reject strings containing ASCII control characters (U+0000–U+001F, U+007F),
 * excluding TAB (0x09), LF (0x0A), and CR (0x0D) which are acceptable in text.
 * Returns the string with control characters removed (never throws — safe fallback).
 */
export function stripControlChars(value: string): string {
  return value.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/gu, '');
}

/**
 * Normalise smart/curly quotes and common Unicode dashes in HTML content captured
 * from Word's getHtml(). Ensures round-trip fidelity when inserting into another document.
 */
export function normaliseQuotesAndEncoding(html: string): string {
  return html
    // Curly double quotes → straight
    .replace(/[“”]/g, '"')
    // Curly single quotes / apostrophe → straight
    .replace(/[‘’]/g, "'")
    // HTML entity equivalents of the above (Office sometimes emits these)
    .replace(/&ldquo;|&rdquo;/g, '"')
    .replace(/&lsquo;|&rsquo;|&apos;/g, "'")
    // En-dash and em-dash entities → literal Unicode (already safe in UTF-8 HTML)
    .replace(/&ndash;/g, '–')
    .replace(/&mdash;/g, '—');
}

// ---------------------------------------------------------------------------
// Jurisdiction gate (pure function — fully testable without Office JS)
// ---------------------------------------------------------------------------

export type DocJurisdiction = 'SG' | 'UK';

/**
 * Determine whether a cross-jurisdiction warning is required before inserting this entry.
 *
 * Rule (spec §B.5.3):
 *   Show the blocking modal if entry.metadata.jurisdiction ≠ docJurisdiction
 *   OR entry.metadata.jurisdiction === 'OTHER'.
 *
 * Mapping: The config layer uses 'UK' for England & Wales; the clause schema uses 'EW'.
 * Both refer to the same jurisdiction — they match.
 *
 * Returns true = show the warning modal before inserting.
 */
export function requiresJurisdictionWarning(
  clauseJurisdiction: 'SG' | 'EW' | 'OTHER',
  docJurisdiction: DocJurisdiction,
): boolean {
  if (clauseJurisdiction === 'OTHER') return true;
  // Map EW ↔ UK for comparison
  const clauseNorm: string = clauseJurisdiction === 'EW' ? 'UK' : clauseJurisdiction;
  return clauseNorm !== docJurisdiction;
}

/**
 * Human-readable label for a clause jurisdiction (for use in the warning modal copy).
 */
export function jurisdictionLabel(j: 'SG' | 'EW' | 'OTHER', other?: string): string {
  if (j === 'SG') return 'Singapore';
  if (j === 'EW') return 'England & Wales';
  return other ? `Other (${other})` : 'Other / unspecified jurisdiction';
}

/**
 * Human-readable label for the document jurisdiction (config layer uses 'UK' for E&W).
 */
export function docJurisdictionLabel(dj: DocJurisdiction): string {
  if (dj === 'UK') return 'England & Wales';
  return 'Singapore';
}

// ---------------------------------------------------------------------------
// Content-control tag helpers
// ---------------------------------------------------------------------------

/**
 * Build the content control tag for a CLAUSE tracking CC.
 * Format: CLAUSE:<id>:v<version>  (spec §B.5.3)
 *
 * NOTE: this value is assigned to the typed `ContentControl.tag` setter, which
 * takes a PLAIN string — Office serialises it internally. XML-escaping here
 * would store a literal `&amp;`/`&lt;` and break the drift round-trip
 * (drift.ts parseClauseTag reads the raw tag back). The id is an internal
 * UUID, so it carries no XML metacharacters in practice.
 */
export function buildClauseTag(id: string, version: number): string {
  return `CLAUSE:${id}:v${version}`;
}

/**
 * Build the content control title for a CLAUSE tracking CC.
 * Format: <title> (rev <lastReviewed ISO>)
 *
 * Assigned to the typed `ContentControl.title` setter — a PLAIN string. Do not
 * XML-escape: a clause title such as "Indemnity & Liability" must display
 * verbatim, not as "Indemnity &amp; Liability".
 */
export function buildClauseTitle(title: string, lastReviewed: string): string {
  return `${title} (rev ${lastReviewed})`;
}

// ---------------------------------------------------------------------------
// Insert options
// ---------------------------------------------------------------------------

export interface InsertOptions {
  /**
   * When true and the entry has an OOXML blob, use insertOoxml instead of insertHtml.
   * Use for table-bearing clauses to preserve table fidelity.
   */
  preserveTables?: boolean;
  /**
   * For testing: skip the actual Office JS calls (mock-friendly).
   * When true, resolves immediately without calling wordRun.
   */
  _dryRun?: boolean;
}

// ---------------------------------------------------------------------------
// Insert function
// ---------------------------------------------------------------------------

/**
 * Insert a clause entry at the current cursor position.
 *
 * IMPORTANT: the caller MUST have already shown and received confirmation from the
 * jurisdiction warning modal (requiresJurisdictionWarning() gate) BEFORE calling this.
 * This function does NOT check the jurisdiction gate — that is the caller's responsibility
 * (implemented in driver.ts). This separation allows the pure gate logic to be tested
 * independently.
 *
 * Steps (spec §B.5.3):
 * 1. Get the current selection (cursor).
 * 2. insertHtml or insertOoxml (per preserveTables flag).
 * 3. After insert, get the selection again (now surrounding the inserted content).
 * 4. Wrap in a RichText content control with CLAUSE tag + title.
 * 5. ctx.sync().
 *
 * DONE_WITH_CONCERNS: The exact behaviour of insertHtml/insertOoxml with respect to the
 * returned range and subsequent content-control wrapping requires live host verification.
 * In particular, whether the selection after insertHtml encompasses exactly the inserted
 * content varies by Office version. The approach below (re-selecting after insert and
 * calling insertContentControl on that range) is the recommended pattern but must be
 * verified on a live Word host.
 */
export async function insertClause(
  entry: ClauseEntry,
  opts: InsertOptions = {},
): Promise<void> {
  if (opts._dryRun) return;

  const useOoxml = opts.preserveTables === true && typeof entry.body.ooxml === 'string';
  const tag = buildClauseTag(entry.id, entry.audit.version);
  // Title must be ≤ 255 chars (Office JS limit); truncate if necessary.
  const rawTitle = buildClauseTitle(entry.title, entry.metadata.lastReviewed);
  const ccTitle = rawTitle.length > 255 ? rawTitle.slice(0, 252) + '…' : rawTitle;

  await wordRun(async ctx => {
    const sel = ctx.document.getSelection();

    if (useOoxml && entry.body.ooxml) {
      // OOXML path: preserve table fidelity
      const safeOoxml = stripControlChars(entry.body.ooxml);
      sel.insertOoxml(safeOoxml, Word.InsertLocation.replace);
    } else {
      // HTML path: portable default; normalise smart-quotes on insert
      const safeHtml = normaliseQuotesAndEncoding(stripControlChars(entry.body.html));
      sel.insertHtml(safeHtml, Word.InsertLocation.replace);
    }

    // After the insert the selection proxy represents the replaced range.
    // Re-get the selection to obtain the range covering the newly inserted content.
    const insertedRange = ctx.document.getSelection();

    // Wrap in a RichText content control for drift tracking.
    const cc = insertedRange.insertContentControl();
    cc.tag = tag;
    cc.title = ccTitle;
    cc.appearance = Word.ContentControlAppearance.hidden;

    await ctx.sync();
  });
}

// ---------------------------------------------------------------------------
// Document jurisdiction helper (reads from the config layer)
// ---------------------------------------------------------------------------

/**
 * Read the document's effective jurisdiction from the config layer.
 * The config layer stores 'SG' or 'UK'; maps cleanly to DocJurisdiction.
 */
export function getDocJurisdiction(): DocJurisdiction {
  const stored = getStoredJurisdiction();
  return stored === 'SG' ? 'SG' : 'UK';
}
