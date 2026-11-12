/**
 * Precedent / Clause Library — shared type definitions.
 *
 * Spec: spec/features/07-toc-precedent-templates.md §B.4
 *
 * Target jurisdictions: Singapore (SG) and England & Wales (EW) only.
 * "OTHER" is explicitly elevate-risk — the jurisdiction warning gate (B5)
 * is mandatory and non-suppressible for OTHER entries.
 *
 * These types are pure data — no Office JS, no DOM, no I/O.
 */

// ---------------------------------------------------------------------------
// Jurisdiction
// ---------------------------------------------------------------------------

/** The three clause-jurisdiction values. */
export type ClauseJurisdiction = 'SG' | 'EW' | 'OTHER';

// ---------------------------------------------------------------------------
// Clause entry body
// ---------------------------------------------------------------------------

export interface ClauseBody {
  /** HTML representation of the clause (portable default, always present). */
  html: string;
  /**
   * OOXML blob for table-fidelity insert ("preserve tables" path).
   * Captured from range.getOoxml() at save time; optional.
   */
  ooxml?: string | undefined;
  /** Plain text, derived from range.text at save time; used for Fuse.js indexing. */
  plainText: string;
}

// ---------------------------------------------------------------------------
// Metadata
// ---------------------------------------------------------------------------

export interface ClauseMetadata {
  /** REQUIRED. Jurisdiction enum — target is SG + EW; anything else is OTHER (elevated risk). */
  jurisdiction: ClauseJurisdiction;
  /**
   * REQUIRED iff jurisdiction === 'OTHER'.
   * Free-text label for the non-target jurisdiction (e.g. "NY", "Delaware").
   * Optional in type but required in validation when jurisdiction === 'OTHER'.
   */
  jurisdictionOther?: string | undefined;
  /** REQUIRED. Practice area from the firm taxonomy (config-controlled but stored as string). */
  practiceArea: string;
  /**
   * REQUIRED. ISO 8601 date string (YYYY-MM-DD).
   * Displayed as DD Month YYYY per the Dates convention.
   * Must not be in the future. Warn when older than the staleness window (default 24 months).
   */
  lastReviewed: string;
  /** REQUIRED. Non-empty name of the approving practitioner (risk governance). */
  approvedBy: string;
  /**
   * REQUIRED.
   * - 'personal' — stored in localStorage (local store); available only to the current user.
   * - 'shared'   — stored in the shared backend (disabled when no backend configured).
   */
  visibility: 'personal' | 'shared';
  /** Optional tags for search ranking. */
  tags?: string[] | undefined;
  /** Language: en-GB (England & Wales / standard SG drafting) or en-SG (SG variant). */
  language?: 'en-GB' | 'en-SG' | undefined;
  /** Provenance note — which precedent this was sourced from. */
  sourcePrecedent?: string;
  /** Drafting notes surfaced at insert time. */
  notes?: string;
}

// ---------------------------------------------------------------------------
// Audit trail
// ---------------------------------------------------------------------------

export interface ClauseAudit {
  createdBy: string;
  createdAt: string;   // ISO 8601
  updatedAt: string;   // ISO 8601
  /** Increments on each edit. Used by drift detection (B6) — inserted copies carry :v<n>. */
  version: number;
}

// ---------------------------------------------------------------------------
// Clause entry (full record)
// ---------------------------------------------------------------------------

export interface ClauseEntry {
  /** Stable UUID assigned at save time. Never changes on edit. */
  id: string;
  /** REQUIRED. Short human label for search and card display. */
  title: string;
  body: ClauseBody;
  metadata: ClauseMetadata;
  audit: ClauseAudit;
}

// ---------------------------------------------------------------------------
// Index record (lightweight — carried in index.json for Fuse.js loading)
// ---------------------------------------------------------------------------

/**
 * Lightweight record stored in the library index (index.json / localStorage index key).
 * Contains only the fields needed for Fuse.js search + facet filtering.
 * Full body is loaded on demand (B5 insert).
 */
export interface ClauseIndexRecord {
  id: string;
  title: string;
  /** Subset of metadata fields needed for search and facet filtering. */
  metadata: Pick<
    ClauseMetadata,
    'jurisdiction' | 'jurisdictionOther' | 'practiceArea' | 'lastReviewed' |
    'approvedBy' | 'visibility' | 'tags' | 'language'
  >;
  /** Plain text copy for Fuse.js indexing (body.plainText). */
  plainText: string;
  /** Audit.version — carried so drift detection can compare without loading the full entry. */
  version: number;
}

// ---------------------------------------------------------------------------
// Save-clause input shape (before id/audit are assigned)
// ---------------------------------------------------------------------------

export interface SaveClauseInput {
  title: string;
  body: ClauseBody;
  metadata: ClauseMetadata;
  /** The author's display name (from Office context or settings). Used to populate audit.createdBy. */
  createdBy: string;
}

// ---------------------------------------------------------------------------
// Drift detection result
// ---------------------------------------------------------------------------

export interface DriftResult {
  /** Content-control tag that contains the inserted clause reference. */
  ccTag: string;
  /** Clause id extracted from the tag. */
  clauseId: string;
  /** Version stored in the document (the version at insert time). */
  insertedVersion: number;
  /** Current library version (audit.version of the library entry). */
  currentVersion: number;
  /** Current library lastReviewed (DD Month YYYY formatted by the caller). */
  currentLastReviewed: string;
  /** Entry title from the library. */
  title: string;
}

// ---------------------------------------------------------------------------
// Staleness window
// ---------------------------------------------------------------------------

/** Default staleness window in months (24 months = 2 years). */
export const DEFAULT_STALENESS_MONTHS = 24;

/** Large-library warning threshold. Above this count the UI warns + recommends SharePoint. */
export const LARGE_LIBRARY_THRESHOLD = 2000;
