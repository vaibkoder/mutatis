/**
 * Feature 07-A — Table of Contents: firm-level default config + doc-override loading.
 *
 * The "house-default" TocProfile ships with SG/UK legal-drafting conventions:
 *  - Clause Heading / Sub-Clause as primary navigation styles
 *  - Dot leaders + right-aligned page numbers
 *  - Hyperlinked entries
 *
 * Per-document overrides are read from docSettings key "toc.profile".
 * All label values use the Localisation/Config layer (scheduleWord from the
 * active profile) — never hardcoded US-isms.
 *
 * Spec reference: spec/features/07-toc-precedent-templates.md §A.4.
 */

import type { TocProfile, ScheduleIndexConfig } from './types.js';
import { getStorageProvider } from '@shared/storage/index.js';

// ---------------------------------------------------------------------------
// Built-in "house-default" profile
// ---------------------------------------------------------------------------

/**
 * Firm standard TOC profile: Heading 1–3 + bespoke UK/SG firm styles.
 * styleLevelMap is non-empty so \t is emitted (not \o), giving explicit style
 * control rather than relying on Word's built-in outline levels.
 */
export const HOUSE_DEFAULT_PROFILE: TocProfile = {
  id: 'house-default',
  label: 'Firm Standard TOC',
  switches: {
    levels: '1-3',
    useHyperlinks: true,
    hidePageNumbers: false,
    rightAlignPageNo: true,
    preserveTabAndLeader: true,
  },
  styleLevelMap: [
    { style: 'Heading 1',      level: 1 },
    { style: 'Heading 2',      level: 2 },
    { style: 'Heading 3',      level: 3 },
    { style: 'Clause Heading', level: 1 },
    { style: 'Sub-Clause',     level: 2 },
  ],
  tabLeader: 'dots',
};

// ---------------------------------------------------------------------------
// Default Index of Schedules config
// ---------------------------------------------------------------------------

/**
 * Default configuration for the Index of Schedules.
 * Label and scheduleStyles must never reference US-isms like "Exhibit".
 */
export const DEFAULT_SCHEDULE_INDEX_CONFIG: ScheduleIndexConfig = {
  enabled: true,
  label: 'Index of Schedules',           // en-GB, per spec CONVENTIONS.md Terminology
  scheduleStyles: ['Schedule Heading', 'Schedule Title'],
  level: 1,
  separateField: true,                   // Strategy 1 (live TOC field) preferred
  placement: 'afterMainToc',
};

// ---------------------------------------------------------------------------
// Storage key
// ---------------------------------------------------------------------------

const DOC_SETTINGS_KEY_PROFILE = 'toc.profile';

// ---------------------------------------------------------------------------
// Profile loading
// ---------------------------------------------------------------------------

/**
 * Load the active TocProfile.
 * Priority: docSettings override → firm default (house-default).
 * Falls back to HOUSE_DEFAULT_PROFILE on any error.
 */
export async function loadActiveProfile(): Promise<TocProfile> {
  try {
    const store = getStorageProvider().docSettings;
    const override = await store.get(DOC_SETTINGS_KEY_PROFILE) as TocProfile | undefined;
    if (override && typeof override.id === 'string') {
      return override;
    }
  } catch {
    // Non-fatal; fall through to default.
  }
  return HOUSE_DEFAULT_PROFILE;
}

/**
 * Persist a TocProfile override to the current document's settings.
 * Pass `null` to clear the override and revert to the firm default.
 */
export async function saveProfileOverride(profile: TocProfile | null): Promise<void> {
  const store = getStorageProvider().docSettings;
  if (profile === null) {
    await store.delete(DOC_SETTINGS_KEY_PROFILE);
  } else {
    await store.set(DOC_SETTINGS_KEY_PROFILE, profile as unknown);
  }
}

// ---------------------------------------------------------------------------
// All available profiles (firm library)
// ---------------------------------------------------------------------------

/**
 * Returns the full list of available TOC profiles.
 * Currently only the built-in house-default; firms extend by adding entries.
 */
export function listAvailableProfiles(): TocProfile[] {
  return [HOUSE_DEFAULT_PROFILE];
}
