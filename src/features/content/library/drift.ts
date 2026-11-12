/**
 * Precedent / Clause Library — drift detection.
 *
 * Spec: spec/features/07-toc-precedent-templates.md §B.8 ("Clause version drift")
 *
 * "Check clauses" action:
 *   1. Load all content controls (tag, title).
 *   2. Parse tags matching CLAUSE:<id>:v<n>.
 *   3. Compare stored version (n) vs current library entry version.
 *   4. Return DriftResult[] for entries where insertedVersion < currentVersion.
 *
 * No auto-update. The lawyer decides whether to replace an inserted copy.
 *
 * Pure parsing helpers are exported separately so unit tests can test them
 * without needing Office JS.
 */

import type { DriftResult } from './types.js';
import { getClause } from './library-store.js';
import { wordRun } from '@shared/office/index.js';

// ---------------------------------------------------------------------------
// Tag parsing (pure — testable without Office JS)
// ---------------------------------------------------------------------------

/** Pattern: CLAUSE:<id>:v<version> (spec §B.5.3) */
const CLAUSE_TAG_PATTERN = /^CLAUSE:([^:]+):v(\d+)$/;

export interface ParsedClauseTag {
  clauseId: string;
  version: number;
}

/**
 * Parse a content-control tag string.
 * Returns null when the tag does not match the CLAUSE pattern.
 */
export function parseClauseTag(tag: string): ParsedClauseTag | null {
  const m = CLAUSE_TAG_PATTERN.exec(tag.trim());
  if (!m) return null;
  const version = parseInt(m[2] ?? '', 10);
  if (isNaN(version)) return null;
  return { clauseId: m[1] ?? '', version };
}

// ---------------------------------------------------------------------------
// Date formatting helper
// ---------------------------------------------------------------------------

/** Format an ISO date string as DD Month YYYY (spec §B.6 / CONVENTIONS.md §Dates). */
export function formatDateDDMonthYYYY(iso: string): string {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) return iso;
  const d = new Date(`${iso}T00:00:00Z`);
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleDateString('en-GB', {
    day: '2-digit',
    month: 'long',
    year: 'numeric',
    timeZone: 'UTC',
  });
}

// ---------------------------------------------------------------------------
// Main scan function (Office JS required)
// ---------------------------------------------------------------------------

/**
 * Scan all content controls in the active document, parse CLAUSE tracking tags,
 * and compare against the library. Returns a list of DriftResult for each
 * inserted clause whose library version is newer than the stored version.
 *
 * Returns an empty array when no drift is found or when the document contains
 * no CLAUSE content controls.
 */
export async function scanDrift(): Promise<DriftResult[]> {
  const results: DriftResult[] = [];

  // Load all content-control tags in one batched sync.
  let parsedTags: Array<{ ccTag: string; parsed: ParsedClauseTag }> = [];

  await wordRun(async ctx => {
    const controls = ctx.document.contentControls;
    controls.load('tag,title');
    await ctx.sync();

    for (const cc of controls.items) {
      const tag = cc.tag ?? '';
      const parsed = parseClauseTag(tag);
      if (parsed) {
        parsedTags.push({ ccTag: tag, parsed });
      }
    }
  });

  // For each parsed CLAUSE tag, look up the library entry and compare versions.
  // Done outside the Word.run to keep the context lean (no range holds).
  for (const { ccTag, parsed } of parsedTags) {
    const entry = await getClause(parsed.clauseId);
    if (!entry) {
      // Entry deleted from library — not flagged as drift (it's a missing entry).
      continue;
    }

    if (parsed.version < entry.audit.version) {
      results.push({
        ccTag,
        clauseId: parsed.clauseId,
        insertedVersion: parsed.version,
        currentVersion: entry.audit.version,
        currentLastReviewed: formatDateDDMonthYYYY(entry.metadata.lastReviewed),
        title: entry.title,
      });
    }
  }

  return results;
}
