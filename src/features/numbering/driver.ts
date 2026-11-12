/**
 * Feature 06 — Numbering: driver (signals + orchestration).
 *
 * Owns the reactive state for the Numbering tab. Called by NumberingPanel.tsx.
 * Uses Preact signals for fine-grained reactivity.
 */

import { signal, computed } from '@preact/signals';
import type { NumberingScheme } from './types.js';
import type { ApplyScope, ApplyReport, ReviewState, BoundaryState } from './types.js';
import { wordRun } from '@shared/office/index.js';
import { getActiveProfile } from '@shared/config/index.js';
import { getNumberingConfig } from '@shared/config/index.js';
import { DocModel } from '@shared/doc-model/index.js';
import { buildSectionIndex } from '@shared/section-index/index.js';
import { findingsByCategory } from '@shared/findings/index.js';
import {
  applyToDocumentBody,
  applyToCurrentSelection,
  classifyParagraphs,
} from './apply.js';
import { changeLevelOnSelection, boundaryState } from './level-change.js';
import { reviewCrossReferences, clearReviewHighlights } from './review.js';
import {
  listUserSchemes,
  saveScheme,
  exportSchemes,
  importSchemes,
} from './scheme-store.js';
import type { CollisionResolution } from './types.js';

// ---------------------------------------------------------------------------
// Signals
// ---------------------------------------------------------------------------

/** All schemes available in the library (built-in + user). */
export const allSchemes = signal<NumberingScheme[]>([]);

/** Id of the currently selected scheme. */
export const selectedSchemeId = signal<string>('preset-uk-sg-commercial');

/** Apply scope: 'selection' or 'document'. */
export const applyScope = signal<ApplyScope>('selection');

/** Result of the last apply/promote/demote operation. */
export const lastApplyReport = signal<ApplyReport | null>(null);

/** Whether an apply/promote/demote operation is in progress. */
export const operationRunning = signal(false);

/** State of the cross-reference review panel. */
export const reviewState = signal<ReviewState>({ count: 0, items: [], active: false });

/** Boundary state for promote/demote buttons. Updated from DocModel selection context. */
export const promoteDemoteState = signal<BoundaryState>({ canPromote: false, canDemote: true });

/** Non-blocking banners (Word Online, tracked changes, style-linked). */
export const banners = signal<string[]>([]);

/** Error message for the last failed operation. */
export const operationError = signal<string | null>(null);

/** The scheme currently open in the level editor (null = none). */
export const editingScheme = signal<NumberingScheme | null>(null);

/** Whether the edited scheme has unsaved changes. */
export const editingDirty = signal(false);

// ---------------------------------------------------------------------------
// Derived
// ---------------------------------------------------------------------------

export const selectedScheme = computed<NumberingScheme | undefined>(
  () => allSchemes.value.find(s => s.id === selectedSchemeId.value),
);

// ---------------------------------------------------------------------------
// Initialise (called once on Numbering tab mount)
// ---------------------------------------------------------------------------

export async function activateNumbering(): Promise<void> {
  const profile = getActiveProfile();
  const config = getNumberingConfig(profile);

  // Build the library: built-in presets + user schemes.
  const builtIn = profile.numberingPresets as NumberingScheme[];
  let userSchemes: NumberingScheme[] = [];
  try { userSchemes = await listUserSchemes(); } catch { /* non-fatal */ }
  allSchemes.value = [...builtIn, ...userSchemes];

  // Pre-select the default scheme.
  selectedSchemeId.value = config.defaultSchemeId;

  // Reset state.
  banners.value = [];
  operationError.value = null;
  reviewState.value = { count: 0, items: [], active: false };
}

// ---------------------------------------------------------------------------
// Apply scheme
// ---------------------------------------------------------------------------

export async function runApply(): Promise<void> {
  if (operationRunning.value) return;
  const scheme = selectedScheme.value;
  if (!scheme) { operationError.value = 'No scheme selected.'; return; }

  const profile = getActiveProfile();

  // Gather static-numbering findings (read-only consumer of feature 02).
  const staticFindings = Array.from(findingsByCategory('proof').values()).flat()
    .filter(f => f.category === 'static-numbering/manually-typed');

  operationRunning.value = true;
  operationError.value = null;
  banners.value = [];

  try {
    let report: ApplyReport;

    await wordRun(async ctx => {
      if (applyScope.value === 'document') {
        report = await applyToDocumentBody(scheme, staticFindings, ctx);
      } else {
        // Use the live selection path — do NOT pass DocModel.paragraphs here.
        // DocModel nodes carry a capturedPara closure from a prior Word.run; using
        // them in a new Word.run throws "Cannot use the object across different
        // request contexts." applyToCurrentSelection() fetches paragraphs fresh
        // via ctx.document.getSelection() within this same context.
        report = await applyToCurrentSelection(scheme, staticFindings, ctx);
      }
      lastApplyReport.value = report;
      _updateBannersFromReport(report);

      // Post-apply cross-reference review.
      const body = ctx.document.body;
      reviewState.value = await reviewCrossReferences(body, profile, ctx);
    });

  } catch (err) {
    operationError.value = err instanceof Error ? err.message : 'Apply failed.';
  } finally {
    operationRunning.value = false;
  }
}

// ---------------------------------------------------------------------------
// Promote / Demote
// ---------------------------------------------------------------------------

export async function runPromote(): Promise<void> {
  await _runLevelChange(-1);
}

export async function runDemote(): Promise<void> {
  await _runLevelChange(1);
}

async function _runLevelChange(delta: -1 | 1): Promise<void> {
  if (operationRunning.value) return;
  const profile = getActiveProfile();

  operationRunning.value = true;
  operationError.value = null;

  try {
    await wordRun(async ctx => {
      // Same cross-context issue as apply: DocModel nodes must not be passed
      // into a new Word.run. changeLevelOnSelection() uses ctx.document.getSelection().
      await changeLevelOnSelection(delta, ctx);

      // Post-op cross-reference review.
      const body = ctx.document.body;
      reviewState.value = await reviewCrossReferences(body, profile, ctx);
    });
  } catch (err) {
    operationError.value = err instanceof Error ? err.message : 'Level change failed.';
  } finally {
    operationRunning.value = false;
  }
}

// ---------------------------------------------------------------------------
// Clear review highlights
// ---------------------------------------------------------------------------

export async function runClearReview(): Promise<void> {
  const profile = getActiveProfile();
  const config = getNumberingConfig(profile);

  await wordRun(async ctx => {
    await clearReviewHighlights(ctx.document.body, config.reviewHighlightColour, ctx);
  }).catch(() => { /* best-effort */ });

  reviewState.value = { count: 0, items: [], active: false };
}

// ---------------------------------------------------------------------------
// Scheme management
// ---------------------------------------------------------------------------

export async function selectScheme(id: string): Promise<void> {
  selectedSchemeId.value = id;
  editingScheme.value = null;
  editingDirty.value = false;
}

export async function saveCurrentEdit(): Promise<void> {
  const scheme = editingScheme.value;
  if (!scheme) return;
  await saveScheme(scheme);
  // Refresh library.
  const existing = allSchemes.value.findIndex(s => s.id === scheme.id);
  if (existing >= 0) {
    allSchemes.value = [...allSchemes.value.slice(0, existing), scheme, ...allSchemes.value.slice(existing + 1)];
  } else {
    allSchemes.value = [...allSchemes.value, scheme];
  }
  editingDirty.value = false;
}

export async function runExport(schemeIds?: string[]): Promise<void> {
  const toExport = schemeIds
    ? allSchemes.value.filter(s => schemeIds.includes(s.id))
    : allSchemes.value.filter(s => !s.builtIn);

  const blob = exportSchemes(toExport);
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'numbering-schemes.json';
  a.click();
  URL.revokeObjectURL(url);
}

export async function runImport(blob: Blob, resolution: CollisionResolution = 'replace'): Promise<string[]> {
  const profile = getActiveProfile();
  const builtIn = profile.numberingPresets as NumberingScheme[];
  const result = await importSchemes(blob, resolution, builtIn);

  if (result.errors.length === 0 && result.imported.length > 0) {
    // Refresh library.
    const existing = new Set(allSchemes.value.map(s => s.id));
    const newSchemes = result.imported.filter(s => !existing.has(s.id));
    const updated = allSchemes.value.map(s => result.imported.find(i => i.id === s.id) ?? s);
    allSchemes.value = [...updated, ...newSchemes];
  }

  return result.errors;
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

function _updateBannersFromReport(report: ApplyReport): void {
  const msgs: string[] = [];

  if (report.degradedLevels.length > 0) {
    msgs.push(
      `Word Online: levels ${report.degradedLevels.join(', ')} could not be fully applied. Scheme applied as far as supported.`,
    );
  }
  if (report.hasMultipleSections) {
    msgs.push(
      'This document has multiple sections; top-level clause numbering may restart at a section break — verify manually.',
    );
  }
  if (report.styleLinkedSkipped) {
    msgs.push(
      'A style-linked list was detected; it was not modified to avoid breaking your template.',
    );
  }
  if (report.staticParasSkipped > 0) {
    msgs.push(
      `${report.staticParasSkipped} paragraph(s) appear to use manually typed numbers and were not converted (see Proof tab).`,
    );
  }

  banners.value = msgs;
}

/** Update boundary state from current DocModel selection (called from panel on mount/update). */
export function updateBoundaryState(): void {
  const paras = DocModel.paragraphs.value;
  const staticFindings = Array.from(findingsByCategory('proof').values()).flat()
    .filter(f => f.category === 'static-numbering/manually-typed');
  const classified = classifyParagraphs(paras, staticFindings);
  // We don't have live listItem.level here without a Word.run — use a conservative default.
  // The panel shows promote/demote as always available unless the DocModel is empty.
  const hasParas = classified.some(c => c.kind === 'convertible');
  promoteDemoteState.value = boundaryState(hasParas ? [1] : []);
}

/** Build the section index (used for section-aware apply — not yet wired to UI, §3). */
export function buildIndex() {
  const profile = getActiveProfile();
  const paras = DocModel.paragraphs.value;
  return buildSectionIndex(paras, profile);
}
