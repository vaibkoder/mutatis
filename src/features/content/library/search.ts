/**
 * Precedent / Clause Library — Fuse.js search + faceted pre-filter.
 *
 * Spec: spec/features/07-toc-precedent-templates.md §B.4.2 / §B.5.2
 *
 * Search keys and weights per spec §B.4.2:
 *   title           weight 0.50
 *   metadata.tags   weight 0.20
 *   metadata.practiceArea weight 0.15
 *   plainText       weight 0.15
 *
 * Facet filters applied BEFORE Fuse on the candidate set:
 *   jurisdiction, practiceArea, visibility, stale (older than 24 months)
 *
 * The Fuse instance is rebuilt incrementally on save (addEntry / removeEntry).
 * For a full rebuild, call buildIndex(records).
 *
 * Large-library guidance: above LARGE_LIBRARY_THRESHOLD the caller should warn
 * and recommend the SharePoint backend (B7 UI responsibility).
 */

import Fuse, { type IFuseOptions } from 'fuse.js';
import type { ClauseIndexRecord } from './types.js';
import { DEFAULT_STALENESS_MONTHS } from './types.js';

// ---------------------------------------------------------------------------
// Fuse.js configuration (spec §B.4.2)
// ---------------------------------------------------------------------------

type FuseRecord = {
  id: string;
  title: string;
  'metadata.practiceArea': string;
  'metadata.tags': string[];
  plainText: string;
};

const FUSE_OPTIONS: IFuseOptions<FuseRecord> = {
  keys: [
    { name: 'title',                    weight: 0.50 },
    { name: 'metadata.tags',            weight: 0.20 },
    { name: 'metadata.practiceArea',    weight: 0.15 },
    { name: 'plainText',                weight: 0.15 },
  ],
  threshold: 0.34,
  ignoreLocation: true,
  minMatchCharLength: 3,
  shouldSort: true,
  includeScore: true,
};

// ---------------------------------------------------------------------------
// Facet filter types
// ---------------------------------------------------------------------------

export interface FacetFilters {
  /** Restrict to a specific jurisdiction (e.g. 'SG', 'EW', 'OTHER'). Undefined = no filter. */
  jurisdiction?: string;
  /** Restrict to a specific practice area string. Undefined = no filter. */
  practiceArea?: string;
  /** Restrict to 'personal' or 'shared'. Undefined = no filter. */
  visibility?: 'personal' | 'shared';
  /** When true, restrict to stale entries only (older than stalenessMonths). */
  staleOnly?: boolean;
  /** Staleness window in months (default DEFAULT_STALENESS_MONTHS). */
  stalenessMonths?: number;
}

// ---------------------------------------------------------------------------
// Search result
// ---------------------------------------------------------------------------

export interface SearchResult {
  record: ClauseIndexRecord;
  /** Fuse.js relevance score (lower = better match; 0 = perfect). Undefined for facet-only results. */
  score?: number;
}

// ---------------------------------------------------------------------------
// Stale helper
// ---------------------------------------------------------------------------

function isStale(lastReviewed: string, stalenessMonths: number, now: Date): boolean {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(lastReviewed)) return false;
  const reviewed = new Date(`${lastReviewed}T00:00:00Z`);
  if (isNaN(reviewed.getTime())) return false;
  const staleAfter = new Date(reviewed.getTime());
  staleAfter.setUTCMonth(staleAfter.getUTCMonth() + stalenessMonths);
  const todayUtc = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate()));
  return staleAfter < todayUtc;
}

// ---------------------------------------------------------------------------
// Flat projection for Fuse
// ---------------------------------------------------------------------------

function toFuseRecord(record: ClauseIndexRecord): FuseRecord {
  return {
    id: record.id,
    title: record.title,
    'metadata.practiceArea': record.metadata.practiceArea,
    'metadata.tags': record.metadata.tags ?? [],
    plainText: record.plainText,
  };
}

// ---------------------------------------------------------------------------
// LibrarySearch class
// ---------------------------------------------------------------------------

/** Stateful Fuse.js search instance with incremental index management. */
export class LibrarySearch {
  private _records: ClauseIndexRecord[] = [];
  private _fuse: Fuse<FuseRecord>;

  constructor() {
    this._fuse = new Fuse<FuseRecord>([], FUSE_OPTIONS);
  }

  /** Replace the entire index with the supplied records. */
  buildIndex(records: ClauseIndexRecord[]): void {
    this._records = [...records];
    this._fuse.setCollection(records.map(toFuseRecord));
  }

  /** Add or update a single record in the index (incremental, called after save). */
  upsertEntry(record: ClauseIndexRecord): void {
    const existing = this._records.findIndex(r => r.id === record.id);
    if (existing >= 0) {
      this._records[existing] = record;
    } else {
      this._records.push(record);
    }
    // Rebuild the Fuse collection from the updated records array.
    // Fuse.js 7.x exposes setCollection for full replacement; there is no partial update API.
    this._fuse.setCollection(this._records.map(toFuseRecord));
  }

  /** Remove a record from the index by id. */
  removeEntry(id: string): void {
    this._records = this._records.filter(r => r.id !== id);
    this._fuse.setCollection(this._records.map(toFuseRecord));
  }

  /** Total number of records in the index (before any filtering). */
  get size(): number {
    return this._records.length;
  }

  /**
   * Search the library.
   *
   * Facet filters are applied first to narrow the candidate set.
   * If the query string is non-empty and at least minMatchCharLength chars, Fuse runs on
   * the filtered candidate set. If the query is empty, the filtered set is returned as-is.
   *
   * @param query  Free-text search query. Empty/blank = browse (facets only).
   * @param facets Optional facet constraints (jurisdiction, practiceArea, visibility, stale).
   * @param now    Override "now" for staleness calculation (testing).
   */
  search(
    query: string,
    facets: FacetFilters = {},
    now: Date = new Date(),
  ): SearchResult[] {
    const stalenessMonths = facets.stalenessMonths ?? DEFAULT_STALENESS_MONTHS;

    // 1. Facet pre-filter
    const candidates = this._records.filter(r => {
      if (facets.jurisdiction !== undefined && r.metadata.jurisdiction !== facets.jurisdiction) return false;
      if (facets.practiceArea !== undefined && r.metadata.practiceArea !== facets.practiceArea) return false;
      if (facets.visibility !== undefined && r.metadata.visibility !== facets.visibility) return false;
      if (facets.staleOnly === true && !isStale(r.metadata.lastReviewed, stalenessMonths, now)) return false;
      return true;
    });

    const trimmed = query.trim();

    // 2. If no query or too short, return the facet candidates ranked by lastReviewed desc.
    if (trimmed.length < 3) {
      return candidates
        .slice()
        .sort((a, b) => b.metadata.lastReviewed.localeCompare(a.metadata.lastReviewed))
        .map(record => ({ record }));
    }

    // 3. Run Fuse on the candidate set (create a temporary Fuse over the candidate slice).
    const candidateFuse = new Fuse<FuseRecord>(candidates.map(toFuseRecord), FUSE_OPTIONS);
    const fuseResults = candidateFuse.search(trimmed);

    // Map Fuse results back to full ClauseIndexRecord by id
    const recordMap = new Map<string, ClauseIndexRecord>(candidates.map(r => [r.id, r]));
    const out: SearchResult[] = [];
    for (const r of fuseResults) {
      const record = recordMap.get(r.item.id);
      if (!record) continue;
      const sr: SearchResult = r.score !== undefined
        ? { record, score: r.score }
        : { record };
      out.push(sr);
    }
    return out;
  }
}

// ---------------------------------------------------------------------------
// Module-level singleton (shared across the panel)
// ---------------------------------------------------------------------------

let _librarySearch: LibrarySearch | null = null;

/** Get (or lazily create) the module-level LibrarySearch singleton. */
export function getLibrarySearch(): LibrarySearch {
  if (!_librarySearch) {
    _librarySearch = new LibrarySearch();
  }
  return _librarySearch;
}

/** Reset the singleton (for testing). */
export function _resetLibrarySearch(): void {
  _librarySearch = null;
}
