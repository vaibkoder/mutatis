/**
 * Feature 06 — Numbering upgrade (ADR-0023): whitespace-override persistence.
 *
 * The user sets ONE global whitespace rule for assisted removal (constraint 3).
 * Stored machine-local via the Storage abstraction; falls back to the jurisdiction
 * profile's NumberingConfig default ('auto') when unset.
 */

import { getStorageProvider } from '@shared/storage/index.js';
import { getActiveProfile, getNumberingConfig } from '@shared/config/index.js';
import type { NumberingWhitespaceOverride } from '@shared/config/index.js';

const PREF_KEY = 'numbering.upgrade.whitespaceOverride';

const VALID: readonly NumberingWhitespaceOverride[] = [
  'auto',
  'collapse-to-one-space',
  'strip-to-flush',
  'keep-one-tab',
];

function isValid(v: unknown): v is NumberingWhitespaceOverride {
  return typeof v === 'string' && (VALID as readonly string[]).includes(v);
}

/** The config-default override for the active jurisdiction profile. */
export function defaultWhitespaceOverride(): NumberingWhitespaceOverride {
  return getNumberingConfig(getActiveProfile()).whitespaceOverride;
}

/**
 * Read the user's persisted override, or the profile default when unset/invalid.
 * Never throws — storage failure falls back to the default.
 */
export async function getWhitespaceOverride(): Promise<NumberingWhitespaceOverride> {
  try {
    const stored = await getStorageProvider().local.get(PREF_KEY);
    if (isValid(stored)) return stored;
  } catch {
    /* fall through to default */
  }
  return defaultWhitespaceOverride();
}

/**
 * Persist the user's override. Rejects an unknown value (no silent write of a
 * bad setting that would later fall back anyway).
 */
export async function setWhitespaceOverride(value: NumberingWhitespaceOverride): Promise<void> {
  if (!isValid(value)) throw new Error(`Invalid whitespace override: "${String(value)}"`);
  await getStorageProvider().local.set(PREF_KEY, value);
}
