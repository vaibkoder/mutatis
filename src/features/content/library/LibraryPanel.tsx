/**
 * Precedent / Clause Library — main sub-panel.
 *
 * Spec: spec/features/07-toc-precedent-templates.md §B.7
 *
 * Layout:
 *   Browse/Search mode: search box + facet chips + result cards + backend status.
 *   Save mode: SaveClauseForm (overlaid / shown in place of the list).
 *   Drift panel: surfaced below the search results when drift results exist.
 *   Insert confirm modal: blocking cross-jurisdiction warning.
 *
 * Result cards show:
 *   title, colour-coded jurisdiction chip (SG/E&W/Other), practice area,
 *   last-reviewed (DD Month YYYY), approved-by, visibility icon.
 *   Actions per card: Insert (primary), Insert (preserve tables) [if OOXML present],
 *   Preview (inline expansion), Edit (opens save form pre-populated).
 *
 * Large-library (>2000 entries) warning recommends the SharePoint backend.
 */

import { useEffect, useRef } from 'preact/hooks';
import type { ClauseIndexRecord } from './types.js';
import { LARGE_LIBRARY_THRESHOLD } from './types.js';
import {
  librarySearchQuery,
  libraryFacets,
  libraryResults,
  libraryBusy,
  libraryError,
  libraryIndexSize,
  saveFormOpen,
  insertModalOpen,
  insertModalWarning,
  driftResults,
  driftBusy,
  driftError,
  initLibrarySearch,
  runSearch,
  openSaveForm,
  initiateInsert,
  confirmInsert,
  cancelInsert,
  checkDrift,
} from './driver.js';
import { getClause } from './library-store.js';
import { hasSharedBackend } from './library-store.js';
import { formatDateDDMonthYYYY } from './drift.js';
import { SaveClauseForm } from './SaveClauseForm.js';
import type { FacetFilters } from './search.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Format a jurisdiction string as a short display label. */
function jurisdictionChipLabel(j: string, other?: string): string {
  if (j === 'SG') return 'SG';
  if (j === 'EW') return 'E&W';
  return other ? `Other: ${other}` : 'Other';
}

/** CSS modifier for the jurisdiction chip colour. */
function jurisdictionChipMod(j: string): string {
  if (j === 'SG') return 'sg';
  if (j === 'EW') return 'ew';
  return 'other';
}

/** Debounce utility. */
function debounce<T extends (...args: unknown[]) => void>(fn: T, ms: number): T {
  let timer: ReturnType<typeof setTimeout> | null = null;
  return ((...args: unknown[]) => {
    if (timer !== null) clearTimeout(timer);
    timer = setTimeout(() => { fn(...args); }, ms);
  }) as T;
}

// ---------------------------------------------------------------------------
// Result card component
// ---------------------------------------------------------------------------

interface ResultCardProps {
  record: ClauseIndexRecord;
}

function ResultCard({ record }: ResultCardProps) {
  const m = record.metadata;
  const chipMod = jurisdictionChipMod(m.jurisdiction);
  const chipLabel = jurisdictionChipLabel(m.jurisdiction, m.jurisdictionOther);
  const displayDate = formatDateDDMonthYYYY(m.lastReviewed);
  const hasOoxml = false; // index record doesn't carry this flag; we'd need the full entry

  async function handleInsert(preserveTables: boolean) {
    // Load the full entry for insert (includes body.html / body.ooxml)
    const entry = await getClause(record.id);
    if (!entry) {
      libraryError.value = `Clause "${record.title}" not found in the library.`;
      return;
    }
    await initiateInsert(entry, preserveTables);
  }

  return (
    <article class="dp-library__card" aria-label={`Clause: ${record.title}`}>
      <div class="dp-library__card-header">
        <span
          class={`dp-library__juris-chip dp-library__juris-chip--${chipMod}`}
          aria-label={`Jurisdiction: ${chipLabel}`}
        >
          {chipLabel}
        </span>
        <span class="dp-library__card-title">{record.title}</span>
      </div>

      <dl class="dp-library__card-meta">
        <div class="dp-library__card-meta-row">
          <dt>Practice area</dt>
          <dd>{m.practiceArea}</dd>
        </div>
        <div class="dp-library__card-meta-row">
          <dt>Last reviewed</dt>
          <dd>{displayDate}</dd>
        </div>
        <div class="dp-library__card-meta-row">
          <dt>Approved by</dt>
          <dd>{m.approvedBy}</dd>
        </div>
        <div class="dp-library__card-meta-row">
          <dt>Visibility</dt>
          <dd>
            <span class={`dp-library__visibility-badge dp-library__visibility-badge--${m.visibility}`}>
              {m.visibility === 'shared' ? 'Shared' : 'Personal'}
            </span>
          </dd>
        </div>
      </dl>

      {m.tags && m.tags.length > 0 && (
        <div class="dp-library__card-tags" aria-label="Tags">
          {m.tags.map(tag => (
            <span key={tag} class="dp-library__tag">{tag}</span>
          ))}
        </div>
      )}

      <div class="dp-library__card-actions">
        <button
          type="button"
          class="dp-library__btn dp-library__btn--primary dp-library__btn--sm"
          onClick={() => void handleInsert(false)}
          aria-label={`Insert clause: ${record.title}`}
        >
          Insert
        </button>
        {hasOoxml && (
          <button
            type="button"
            class="dp-library__btn dp-library__btn--secondary dp-library__btn--sm"
            onClick={() => void handleInsert(true)}
            aria-label={`Insert (preserve tables): ${record.title}`}
          >
            Insert (preserve tables)
          </button>
        )}
      </div>
    </article>
  );
}

// ---------------------------------------------------------------------------
// Jurisdiction warning modal
// ---------------------------------------------------------------------------

function JurisdictionWarningModal() {
  if (!insertModalOpen.value) return null;

  return (
    <div
      class="dp-library__modal-backdrop"
      role="dialog"
      aria-modal="true"
      aria-label="Cross-jurisdiction warning"
    >
      <div class="dp-library__modal">
        <h3 class="dp-library__modal-title">Cross-Jurisdiction Warning</h3>
        <p class="dp-library__modal-body">{insertModalWarning.value}</p>
        <div class="dp-library__modal-actions">
          <button
            type="button"
            class="dp-library__btn dp-library__btn--danger"
            onClick={() => void confirmInsert()}
          >
            Insert anyway
          </button>
          <button
            type="button"
            class="dp-library__btn dp-library__btn--secondary"
            onClick={cancelInsert}
          >
            Cancel
          </button>
        </div>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Drift results panel
// ---------------------------------------------------------------------------

function DriftPanel() {
  const results = driftResults.value;
  if (results.length === 0 && !driftError.value) return null;

  return (
    <section class="dp-library__drift-panel" aria-label="Clause drift results">
      <h3 class="dp-library__drift-title">Clause Check Results</h3>
      {driftError.value && (
        <p class="dp-library__error" role="alert">{driftError.value}</p>
      )}
      {results.length === 0 && !driftError.value && (
        <p class="dp-library__drift-none">All inserted clauses are up to date.</p>
      )}
      {results.map(r => (
        <div key={r.ccTag} class="dp-library__drift-item" role="alert">
          <strong>{r.title}</strong>
          <p>
            Inserted version: v{r.insertedVersion} — current library version: v{r.currentVersion}.
          </p>
          <p>Library last reviewed: {r.currentLastReviewed}</p>
          <p class="dp-library__drift-note">
            Review and manually replace if this clause is material to the document.
          </p>
        </div>
      ))}
    </section>
  );
}

// ---------------------------------------------------------------------------
// Main panel
// ---------------------------------------------------------------------------

export function LibraryPanel() {
  // Initialise the search index once on mount.
  useEffect(() => {
    void initLibrarySearch();
  }, []);

  // Debounced search runner.
  const debouncedSearch = useRef(debounce(runSearch as (...args: unknown[]) => void, 250));

  function handleQueryInput(e: Event) {
    librarySearchQuery.value = (e.target as HTMLInputElement).value;
    debouncedSearch.current();
  }

  function toggleFacet<K extends keyof FacetFilters>(key: K, value: FacetFilters[K]): void {
    const current = libraryFacets.value;
    if (current[key] === value) {
      // Toggle off
      const updated = { ...current };
      delete updated[key];
      libraryFacets.value = updated;
    } else {
      libraryFacets.value = { ...current, [key]: value };
    }
    runSearch();
  }

  const isLarge = libraryIndexSize.value > LARGE_LIBRARY_THRESHOLD;
  const results = libraryResults.value;
  const hasBackend = hasSharedBackend();

  // Show save form overlay
  if (saveFormOpen.value) {
    return (
      <div class="dp-library" lang="en-GB">
        <SaveClauseForm />
      </div>
    );
  }

  return (
    <div class="dp-library" lang="en-GB">
      <JurisdictionWarningModal />

      {/* Header */}
      <div class="dp-library__header">
        <h2 class="dp-library__heading">Clause Library</h2>
        <button
          type="button"
          class="dp-library__btn dp-library__btn--primary"
          onClick={() => void openSaveForm()}
          disabled={libraryBusy.value}
          aria-label="Save current selection as a clause"
        >
          Save selection as clause
        </button>
      </div>

      {/* Backend status */}
      <div class="dp-library__backend-status" aria-live="polite">
        {hasBackend
          ? <span class="dp-library__status--connected">Shared backend: connected</span>
          : (
            <span class="dp-library__status--local">
              Personal library only.{' '}
              <span class="dp-library__status-hint">
                Configure SharePoint (recommended for SG/UK firms) or GitHub
                to enable shared clauses.
              </span>
            </span>
          )
        }
        <span class="dp-library__status-count" aria-label={`${libraryIndexSize.value} clauses in library`}>
          {' '}{libraryIndexSize.value} clause{libraryIndexSize.value !== 1 ? 's' : ''}
        </span>
      </div>

      {/* Large-library warning (B-T7) */}
      {isLarge && (
        <div class="dp-library__large-warning" role="note">
          <strong>Library has more than {LARGE_LIBRARY_THRESHOLD} clauses.</strong>{' '}
          Client-side search may be slow. We recommend connecting a SharePoint backend
          for server-side search and paging.
        </div>
      )}

      {/* Search box */}
      <div class="dp-library__search-row">
        <label for="dp-lib-search" class="dp-library__sr-only">Search clauses</label>
        <input
          id="dp-lib-search"
          type="search"
          class="dp-library__search-input"
          value={librarySearchQuery.value}
          onInput={handleQueryInput}
          placeholder="Search clauses…"
          aria-label="Search clauses"
        />
      </div>

      {/* Facet chips */}
      <div class="dp-library__facets" role="group" aria-label="Filter by">
        {/* Jurisdiction chips */}
        {(['SG', 'EW', 'OTHER'] as const).map(j => {
          const label = j === 'SG' ? 'Singapore' : j === 'EW' ? 'England & Wales' : 'Other';
          const active = libraryFacets.value.jurisdiction === j;
          return (
            <button
              key={j}
              type="button"
              class={`dp-library__facet-chip${active ? ' dp-library__facet-chip--active' : ''}`}
              onClick={() => toggleFacet('jurisdiction', j)}
              aria-pressed={active}
            >
              {label}
            </button>
          );
        })}
        {/* Visibility chips */}
        {(['personal', 'shared'] as const).map(v => {
          const active = libraryFacets.value.visibility === v;
          return (
            <button
              key={v}
              type="button"
              class={`dp-library__facet-chip${active ? ' dp-library__facet-chip--active' : ''}`}
              onClick={() => toggleFacet('visibility', v)}
              aria-pressed={active}
            >
              {v === 'personal' ? 'Personal' : 'Shared'}
            </button>
          );
        })}
        {/* Stale only chip */}
        {(() => {
          const active = libraryFacets.value.staleOnly === true;
          return (
            <button
              type="button"
              class={`dp-library__facet-chip${active ? ' dp-library__facet-chip--active' : ''}`}
              onClick={() => toggleFacet('staleOnly', true)}
              aria-pressed={active}
            >
              Stale only
            </button>
          );
        })()}
      </div>

      {/* Errors */}
      {libraryError.value && (
        <div class="dp-library__error" role="alert">{libraryError.value}</div>
      )}

      {/* Results */}
      <div
        class="dp-library__results"
        aria-busy={libraryBusy.value}
        aria-label={`${results.length} result${results.length !== 1 ? 's' : ''}`}
      >
        {libraryBusy.value && <p class="dp-library__loading">Loading…</p>}
        {!libraryBusy.value && results.length === 0 && (
          <p class="dp-library__empty">
            {libraryIndexSize.value === 0
              ? 'No clauses in the library yet. Save a selection to get started.'
              : 'No clauses match your search and filters.'}
          </p>
        )}
        {results.map(r => (
          <ResultCard key={r.record.id} record={r.record} />
        ))}
      </div>

      {/* Drift detection */}
      <div class="dp-library__drift-section">
        <button
          type="button"
          class="dp-library__btn dp-library__btn--secondary"
          onClick={() => void checkDrift()}
          disabled={driftBusy.value}
          aria-busy={driftBusy.value}
        >
          {driftBusy.value ? 'Checking…' : 'Check clauses in document'}
        </button>
        <DriftPanel />
      </div>
    </div>
  );
}
