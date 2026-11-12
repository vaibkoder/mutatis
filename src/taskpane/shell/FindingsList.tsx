import { useState } from 'preact/hooks';
import type { TabId } from '@shared/findings/index.js';
import { findingsByCategory } from '@shared/findings/index.js';
import type { Finding } from '@shared/findings/index.js';
import { DocModel } from '@shared/doc-model/index.js';
import { getActiveProfile } from '@shared/config/index.js';
import { wordRun } from '@shared/office/index.js';
import {
  clearThenEmit as runFactualConsistency,
  type FactualConsistencyAnalysis,
} from '@features/factual-consistency/index.js';
import {
  clearThenEmit as runDefinedTerms,
  type DefinedTermsAnalysis,
} from '@features/defined-terms/index.js';
import { clearThenEmit as runCrossReference, type CrossReferenceAnalysis } from '@features/cross-reference/index.js';
import { clearThenEmit as runPartyRole, type PartyRoleAnalysis } from '@features/party-role/index.js';
import { FindingRow } from '../components/FindingRow.js';
import { AssignSourcePicker } from '../components/AssignSourcePicker.js';
import { categoryLabel, splitFindingsByFeature } from '../components/findings-grouping.js';
import { ProofExports } from './ProofExports.js';

function sortByFrequency(findings: Finding[]): Finding[] {
  return [...findings].sort((a, b) => (b.occurrences?.length ?? 1) - (a.occurrences?.length ?? 1));
}

/**
 * True when the finding matches a free-text filter query (already lowercased).
 * Matches against the surfaced text the user can see: title, detail, the
 * category label, every occurrence's location label, and any source snippet —
 * so a query can be a word, a defined term, or a clause reference like "4.2".
 */
function findingMatches(f: Finding, query: string): boolean {
  const haystack = [
    f.title,
    f.detail,
    categoryLabel(f.category),
    f.location.label,
    ...(f.occurrences?.map(o => o.label) ?? []),
    f.sourceSnippet?.label,
    f.sourceSnippet?.text,
  ];
  return haystack.some(part => part != null && part.toLowerCase().includes(query));
}

/** Filter a category→findings map, dropping categories with no remaining matches. */
function filterGrouped(grouped: Map<string, Finding[]>, query: string): Map<string, Finding[]> {
  const out = new Map<string, Finding[]>();
  for (const [cat, findings] of grouped.entries()) {
    const matched = findings.filter(f => findingMatches(f, query));
    if (matched.length > 0) out.set(cat, matched);
  }
  return out;
}

function countFindings(grouped: Map<string, Finding[]>): number {
  let total = 0;
  for (const findings of grouped.values()) total += findings.length;
  return total;
}

interface FindingsListProps {
  tabId: TabId;
}

export function FindingsList({ tabId }: FindingsListProps) {
  const [running, setRunning] = useState(false);
  const [capMode, setCapMode] = useState<'smart' | 'strict'>('smart');
  const [sortByFreq, setSortByFreq] = useState(false);
  const [query, setQuery] = useState('');
  const [analysis, setAnalysis] = useState<DefinedTermsAnalysis | null>(null);
  const [xrefAnalysis, setXrefAnalysis] = useState<CrossReferenceAnalysis | null>(null);
  const [factualAnalysis, setFactualAnalysis] = useState<FactualConsistencyAnalysis | null>(null);
  const [partyAnalysis, setPartyAnalysis] = useState<PartyRoleAnalysis | null>(null);
  const [assignTarget, setAssignTarget] = useState<Finding | null>(null);
  const allGrouped = findingsByCategory(tabId);
  const trimmedQuery = query.trim().toLowerCase();
  // Client-side filter only — purely narrows what is already on screen.
  const grouped = trimmedQuery ? filterGrouped(allGrouped, trimmedQuery) : allGrouped;

  async function handleRunProof() {
    if (running) return;
    setRunning(true);
    const profile = getActiveProfile();
    let paras = DocModel.allParagraphs.value;
    try {
      await DocModel.load().catch(() => { /* non-fatal: analyse whatever snapshot exists */ });
      paras = DocModel.allParagraphs.value;
      // Sync checks run first so the UI responds immediately. Retain each
      // analysis so the AI-export panel can serialise the raw records.
      setFactualAnalysis(runFactualConsistency(paras, profile));
      setAnalysis(runDefinedTerms(paras, profile, { capModeOverride: capMode }));
      // Party-Role Consistency (feature 09) — reuses the same snapshot.
      setPartyAnalysis(runPartyRole(paras, profile));
    } finally {
      setRunning(false);
    }
    // Cross-reference: try with a live Word.run context (for REF-field detection);
    // fall back to null context (ADR-0014 degraded mode) if Word API is unavailable.
    try {
      await wordRun(async ctx => {
        const result = await runCrossReference(paras, profile, ctx);
        setXrefAnalysis(result);
      });
    } catch {
      try {
        const result = await runCrossReference(paras, profile, null);
        setXrefAnalysis(result);
      } catch {
        // Silently ignore complete failure.
      }
    }
  }

  // Toggling smart/strict re-runs only the defined-terms checks (sub-check 4
  // changes); factual-consistency findings are untouched (separate prefix).
  function handleToggleCap() {
    const next = capMode === 'smart' ? 'strict' : 'smart';
    setCapMode(next);
    if (analysis) {
      setAnalysis(runDefinedTerms(DocModel.allParagraphs.value, getActiveProfile(), { capModeOverride: next }));
    }
  }

  // Split findings into: cross-reference (Feature A), static-numbering (Feature B),
  // defined-terms (Feature 01), others (factual-consistency). Shared with the
  // standalone web view via findings-grouping.ts.
  const {
    xrefOrdered,
    staticNum: staticNumFindings,
    definedTermsOrdered,
    other: otherFindings,
  } = splitFindingsByFeature(grouped);

  const hasAnyFindings = allGrouped.size > 0;
  const hasVisibleFindings = grouped.size > 0;
  const totalCount = countFindings(allGrouped);
  const visibleCount = countFindings(grouped);

  return (
    <div class="dp-findings-list" role="tabpanel" id={`dp-panel-${tabId}`} aria-labelledby={`dp-tab-${tabId}`}>
      {tabId === 'proof' && (
        <div class="dp-proof-actions">
          <button
            class="dp-proof-actions__run"
            type="button"
            onClick={handleRunProof}
            aria-busy={running ? 'true' : 'false'}
            disabled={running}
          >
            {running ? 'Running…' : 'Run Proof checks'}
          </button>
          <button
            class="dp-cap-toggle"
            type="button"
            onClick={handleToggleCap}
            aria-pressed={capMode === 'strict' ? 'true' : 'false'}
            title="Smart flags lowercase uses of defined terms only after a determiner like 'the'. Strict flags every lowercase use."
          >
            Capitalisation: {capMode === 'smart' ? 'Smart' : 'Strict'}
          </button>
          <button
            class="dp-freq-sort"
            type="button"
            onClick={() => setSortByFreq(v => !v)}
            aria-pressed={sortByFreq ? 'true' : 'false'}
            title="Sort findings by how many times the term is used, most-used first."
          >
            Sort: {sortByFreq ? 'Most used' : 'Document order'}
          </button>
        </div>
      )}
      {hasAnyFindings && (
        <div class="dp-findings-search" role="search">
          <input
            type="search"
            class="dp-findings-search__input"
            value={query}
            onInput={e => setQuery((e.currentTarget as HTMLInputElement).value)}
            placeholder="Filter findings — word, term or clause ref…"
            aria-label="Filter findings by word, defined term or clause reference"
          />
          {trimmedQuery && (
            <span class="dp-findings-search__count" role="status">
              {visibleCount} of {totalCount}
            </span>
          )}
        </div>
      )}
      {tabId === 'proof' && analysis && (
        <div class="dp-defined-terms-banner" role="status">
          {analysis.noDefinitions && <p class="dp-defined-terms-banner__empty">No definitions detected.</p>}
          {analysis.crossRefCount > 0 && (
            <p class="dp-defined-terms-banner__crossref">
              {analysis.crossRefCount} term(s) defined by cross-reference within this document (treated as defined).
            </p>
          )}
          {analysis.crossDocCount > 0 && (
            <p class="dp-defined-terms-banner__crossdoc">
              {analysis.crossDocCount} term(s) defined by cross-reference to external documents (not checked).
            </p>
          )}
        </div>
      )}
      {tabId === 'proof' && xrefAnalysis?.fieldsApiUnavailable && (
        <div class="dp-capability-banner" role="status">
          REF-field detection not available on this host — all cross-references shown as manually typed.
        </div>
      )}
      {!hasAnyFindings && (
        <p class="dp-findings-empty" role="status">No issues found.</p>
      )}
      {hasAnyFindings && !hasVisibleFindings && (
        <p class="dp-findings-empty" role="status">No findings match “{query.trim()}”.</p>
      )}

      {/* Feature A — Cross-references (featureId: 'xref'). MUST render before Feature B. */}
      {tabId === 'proof' && xrefOrdered.length > 0 && (
        <section class="dp-feature-group" aria-label="Cross-references">
          <h2 class="dp-feature-group__heading">Cross-references</h2>
          {xrefOrdered.map(({ cat, findings }) => (
            <section key={cat} class="dp-findings-group">
              <h3 class="dp-findings-group__heading">
                {categoryLabel(cat)}
                <span class="dp-findings-group__count" aria-label={`${findings.length} issue${findings.length === 1 ? '' : 's'}`}>
                  {findings.length}
                </span>
              </h3>
              <ul class="dp-findings-group__list" aria-label={categoryLabel(cat)}>
                {findings.map(f => (
                  <li key={f.id}>
                    <FindingRow finding={f} />
                    {(f.category === 'cross-reference/not-found' || f.category === 'cross-reference/ambiguous') && (
                      <button
                        class="dp-assign-source-btn"
                        type="button"
                        onClick={() => setAssignTarget(f)}
                        aria-label={`Assign source for: ${f.title}`}
                      >
                        Assign source
                      </button>
                    )}
                  </li>
                ))}
              </ul>
            </section>
          ))}
        </section>
      )}

      {/* Feature B — Manually-typed clause numbers (featureId: 'staticNumbering'). Separate group. */}
      {tabId === 'proof' && staticNumFindings.length > 0 && (
        <section class="dp-feature-group" aria-label="Manually-typed clause numbers">
          <h2 class="dp-feature-group__heading">Manually-typed clause numbers</h2>
          <ul class="dp-findings-group__list" aria-label="Manually-typed clause numbers">
            {staticNumFindings.map(f => (
              <li key={f.id}>
                <FindingRow finding={f} />
                {/* No fix/convert action for Feature B — ADR-0006 */}
              </li>
            ))}
          </ul>
        </section>
      )}

      {/* Feature 01 — Defined terms */}
      {tabId === 'proof' && definedTermsOrdered.length > 0 && (
        <section class="dp-feature-group" aria-label="Defined terms">
          <h2 class="dp-feature-group__heading">Defined terms</h2>
          {definedTermsOrdered.map(({ cat, findings }) => (
            <section key={cat} class="dp-findings-group">
              <h3 class="dp-findings-group__heading">
                {categoryLabel(cat)}
                <span class="dp-findings-group__count" aria-label={`${findings.length} issue${findings.length === 1 ? '' : 's'}`}>
                  {findings.length}
                </span>
              </h3>
              <ul class="dp-findings-group__list" aria-label={categoryLabel(cat)}>
                {(sortByFreq ? sortByFrequency(findings) : findings).map(f => (
                  <li key={f.id}>
                    <FindingRow finding={f} />
                  </li>
                ))}
              </ul>
            </section>
          ))}
        </section>
      )}

      {/* Other proof findings (factual-consistency). */}
      {otherFindings.map(({ cat: category, findings }) => (
        <section key={category} class="dp-findings-group">
          <h3 class="dp-findings-group__heading">
            {categoryLabel(category)}
            <span class="dp-findings-group__count" aria-label={`${findings.length} issue${findings.length === 1 ? '' : 's'}`}>
              {findings.length}
            </span>
          </h3>
          <ul class="dp-findings-group__list" aria-label={categoryLabel(category)}>
            {(sortByFreq ? sortByFrequency(findings) : findings).map(f => (
              <li key={f.id}>
                <FindingRow finding={f} />
              </li>
            ))}
          </ul>
        </section>
      ))}

      {/* Copy for AI — per-checker exports + composite briefing pack (feature 11). */}
      {tabId === 'proof' && (
        <ProofExports
          defined={analysis}
          xref={xrefAnalysis}
          factual={factualAnalysis}
          party={partyAnalysis}
        />
      )}

      {/* Assign-source picker modal (Feature A only). */}
      {assignTarget && xrefAnalysis && (
        <AssignSourcePicker
          nodes={xrefAnalysis.index.ordered}
          candidateTargets={(assignTarget as Finding & { candidateTargets?: import('@shared/section-index/types.js').SectionNode[] }).candidateTargets ?? []}
          {...(assignTarget.category === 'cross-reference/not-found' ? { unitType: 'clause' } : {})}
          anchor={{ paragraphId: assignTarget.location.paraId ?? '', startOffset: 0, endOffset: 0 }}
          paras={DocModel.allParagraphs.value}
          onComplete={() => setAssignTarget(null)}
          onCancel={() => setAssignTarget(null)}
        />
      )}
    </div>
  );
}
