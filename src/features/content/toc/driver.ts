/**
 * Feature 07-A — Table of Contents: driver (signals + orchestration).
 *
 * Owns all UI-facing state for the TocPanel. Calls the toc-service layer for
 * Office JS operations; never calls Office JS directly.
 *
 * Signals pattern mirrors the Repair and Numbering drivers for consistency.
 */

import { signal } from '@preact/signals';
import { isCapable } from '@shared/office/index.js';
import type { TocProfile, TocRegistry } from './types.js';
import { loadActiveProfile, listAvailableProfiles, saveProfileOverride } from './config.js';
import {
  insertToc,
  insertScopedToc,
  insertScheduleIndexField,
  refreshOneToc,
  refreshAllTocs,
  loadRegistry,
  navigateToToc,
} from './toc-service.js';
import { DEFAULT_SCHEDULE_INDEX_CONFIG } from './config.js';

// ---------------------------------------------------------------------------
// Signals
// ---------------------------------------------------------------------------

/** Currently selected TOC profile. */
export const activeProfile = signal<TocProfile | null>(null);

/** All available profiles for the profile selector. */
export const availableProfiles = signal<TocProfile[]>([]);

/** Registry of inserted TOCs (fieldId, bookmark, scopeLabel). */
export const registry = signal<TocRegistry>([]);

/** True when an Office operation is in progress. */
export const busy = signal(false);

/** User-facing error message for the last failed operation. */
export const error = signal<string | null>(null);

/** True when wordApi15 is absent — triggers the F9 banner. */
export const showF9Banner = signal(false);

/** Informational banners (e.g. "TOC had no entries"). */
export const banners = signal<string[]>([]);

// ---------------------------------------------------------------------------
// Initialise (called on TocPanel mount)
// ---------------------------------------------------------------------------

export async function activateToc(): Promise<void> {
  showF9Banner.value = !isCapable('wordApi15');
  error.value = null;
  banners.value = [];

  const profiles = listAvailableProfiles();
  availableProfiles.value = profiles;

  const loaded = await loadActiveProfile();
  activeProfile.value = loaded;

  await _refreshRegistry();
}

// ---------------------------------------------------------------------------
// Profile management
// ---------------------------------------------------------------------------

/** Select a profile from the available list. */
export async function selectProfile(profileId: string): Promise<void> {
  const found = availableProfiles.value.find(p => p.id === profileId);
  if (!found) return;
  activeProfile.value = found;
  await saveProfileOverride(found);
}

// ---------------------------------------------------------------------------
// Insert TOC
// ---------------------------------------------------------------------------

/** Insert a TOC at the cursor using the active profile. */
export async function runInsertTocAtCursor(): Promise<void> {
  await _runInsert('atCursor', 'TOC (cursor)');
}

/** Insert a TOC at the top of document using the active profile. */
export async function runInsertTocAtTop(): Promise<void> {
  await _runInsert('topOfDocument', 'TOC (document start)');
}

async function _runInsert(
  placement: 'atCursor' | 'topOfDocument',
  label: string,
): Promise<void> {
  const profile = activeProfile.value;
  if (!profile) { error.value = 'No profile selected.'; return; }
  if (busy.value) return;

  busy.value = true;
  error.value = null;
  banners.value = [];

  try {
    const result = await insertToc(profile, { placement, scopeLabel: label });

    if (result.isEmpty) {
      banners.value = [
        'The table of contents has no entries. Apply heading styles (e.g. Heading 1, Clause Heading) to the document first.',
      ];
    }
    if (result.usedFallback) {
      showF9Banner.value = true;
    }

    await _refreshRegistry();
  } catch (err) {
    error.value = err instanceof Error ? err.message : 'Insert failed.';
  } finally {
    busy.value = false;
  }
}

// ---------------------------------------------------------------------------
// Scoped TOC
// ---------------------------------------------------------------------------

/**
 * Insert a scoped TOC for the current selection.
 * The UI must verify the selection is non-empty before calling this.
 */
export async function runInsertScopedToc(label: string): Promise<void> {
  const profile = activeProfile.value;
  if (!profile) { error.value = 'No profile selected.'; return; }
  if (busy.value) return;

  busy.value = true;
  error.value = null;
  banners.value = [];

  try {
    const result = await insertScopedToc(profile, label);

    if (result.isEmpty) {
      banners.value = [
        'The scoped TOC has no entries within the selected range.',
      ];
    }
    if (result.usedFallback) {
      showF9Banner.value = true;
    }

    await _refreshRegistry();
  } catch (err) {
    error.value = err instanceof Error ? err.message : 'Scoped TOC insert failed.';
  } finally {
    busy.value = false;
  }
}

// ---------------------------------------------------------------------------
// Index of Schedules
// ---------------------------------------------------------------------------

/** Generate (or regenerate) the Index of Schedules using Strategy 1. */
export async function runGenerateScheduleIndex(): Promise<void> {
  if (busy.value) return;

  busy.value = true;
  error.value = null;
  banners.value = [];

  try {
    const result = await insertScheduleIndexField(
      DEFAULT_SCHEDULE_INDEX_CONFIG,
      DEFAULT_SCHEDULE_INDEX_CONFIG.placement === 'atCursor' ? 'atCursor' : 'topOfDocument',
    );

    if (result.isEmpty) {
      banners.value = [
        `The Index of Schedules has no entries. Ensure paragraphs use one of these styles: ${DEFAULT_SCHEDULE_INDEX_CONFIG.scheduleStyles.join(', ')}.`,
      ];
    }
    if (result.usedFallback) {
      showF9Banner.value = true;
    }
  } catch (err) {
    error.value = err instanceof Error ? err.message : 'Index of Schedules generation failed.';
  } finally {
    busy.value = false;
  }
}

// ---------------------------------------------------------------------------
// Refresh
// ---------------------------------------------------------------------------

/** Refresh a single TOC from the registry. */
export async function runRefreshOne(fieldId: string): Promise<void> {
  if (busy.value) return;
  busy.value = true;
  error.value = null;

  try {
    await refreshOneToc(fieldId);
    await _refreshRegistry();
  } catch (err) {
    error.value = err instanceof Error ? err.message : 'Refresh failed.';
  } finally {
    busy.value = false;
  }
}

/** Refresh all TOCs and reconcile the registry. */
export async function runRefreshAll(): Promise<void> {
  if (busy.value) return;
  busy.value = true;
  error.value = null;

  try {
    await refreshAllTocs();
    await _refreshRegistry();
  } catch (err) {
    error.value = err instanceof Error ? err.message : 'Refresh all failed.';
  } finally {
    busy.value = false;
  }
}

// ---------------------------------------------------------------------------
// Navigate
// ---------------------------------------------------------------------------

/** Navigate (select) to the TOC identified by fieldId. */
export async function runNavigateToToc(fieldId: string): Promise<void> {
  try {
    await navigateToToc(fieldId);
  } catch {
    /* best-effort navigation */
  }
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

async function _refreshRegistry(): Promise<void> {
  try {
    registry.value = await loadRegistry();
  } catch { /* non-fatal */ }
}
