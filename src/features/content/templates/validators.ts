/**
 * SG/UK template validators.
 *
 * All patterns come from the Config layer (shared/config) — never hardcoded.
 * Jurisdiction-aware: UK_CH accepted under EW profile; SG_UEN accepted under
 * SG profile; each is rejected under the other jurisdiction (C-T3).
 *
 * NO US conventions: Inc./LLC/EIN/State-of-incorporation are forbidden from
 * any option list or validation here.
 */

import { getProfile } from '@shared/config/index.js';
import type { TemplateJurisdiction } from './types.js';

// ---------------------------------------------------------------------------
// Registration number validation (UK_CH / SG_UEN)
// ---------------------------------------------------------------------------

/** Result of a registration-number validation attempt. */
export interface RegNumberValidationResult {
  valid: boolean;
  /** Human-readable error if not valid. */
  error?: string;
}

/**
 * Validate a UK Companies House number against the profile's regex.
 * Pattern: ^[A-Z0-9]{8}$ — two-letter prefixes SC/NI/OC/SO are letters, so
 * they are covered. Reads from UK profile config, never hardcoded.
 */
export function validateUkCh(value: string): RegNumberValidationResult {
  const profile = getProfile('UK');
  const pattern = profile.entityFields.registrationNumbers['UK_CH'];
  if (!pattern) {
    return { valid: false, error: 'UK_CH validator not found in config.' };
  }
  const normalised = value.trim().toUpperCase();
  if (pattern.test(normalised)) {
    return { valid: true };
  }
  return {
    valid: false,
    error: `"${value}" is not a valid Companies House number (must be 8 alphanumeric characters, e.g. 00012345 or SC012345).`,
  };
}

/**
 * Validate a Singapore UEN against the profile's regex.
 * Reads from SG profile config, never hardcoded.
 */
export function validateSgUen(value: string): RegNumberValidationResult {
  const profile = getProfile('SG');
  const pattern = profile.entityFields.registrationNumbers['SG_UEN'];
  if (!pattern) {
    return { valid: false, error: 'SG_UEN validator not found in config.' };
  }
  const normalised = value.trim().toUpperCase();
  if (pattern.test(normalised)) {
    return { valid: true };
  }
  return {
    valid: false,
    error: `"${value}" is not a valid Singapore UEN (expected format: 201912345A, 19990012345K, or similar).`,
  };
}

/**
 * Validate a registration number for the given scheme.
 * Enforces jurisdiction-awareness: UK_CH is rejected under SG profile and
 * vice versa — see C-T3.
 */
export function validateRegNumber(
  value: string,
  scheme: 'UK_CH' | 'SG_UEN',
  jurisdictionProfile: TemplateJurisdiction,
): RegNumberValidationResult {
  if (scheme === 'UK_CH') {
    if (jurisdictionProfile === 'SG') {
      return {
        valid: false,
        error: 'UK Companies House numbers are not accepted under the Singapore jurisdiction profile. Use SG_UEN instead.',
      };
    }
    return validateUkCh(value);
  }

  // scheme === 'SG_UEN'
  if (jurisdictionProfile === 'EW') {
    return {
      valid: false,
      error: 'Singapore UEN numbers are not accepted under the England & Wales jurisdiction profile. Use UK_CH instead.',
    };
  }
  return validateSgUen(value);
}

// ---------------------------------------------------------------------------
// Entity-type options (jurisdiction-aware)
// ---------------------------------------------------------------------------

/**
 * Return the entity-type suffix options for the given jurisdiction.
 * Reads from the profile's entityFields.suffixes — never hardcoded.
 * NO Inc./LLC/EIN/State-of-incorporation are in these lists.
 */
export function getEntityTypeOptions(jurisdictionProfile: TemplateJurisdiction): string[] {
  const profileId = jurisdictionProfile === 'EW' ? 'UK' : 'SG';
  const profile = getProfile(profileId);
  // Return a mutable copy so callers can sort/filter safely.
  return [...profile.entityFields.suffixes];
}

// ---------------------------------------------------------------------------
// General variable validation
// ---------------------------------------------------------------------------

/**
 * Validate a form field value against a TemplateVariable's constraints.
 * Returns an error string if invalid, or null if valid.
 */
export function validateFieldValue(
  value: string,
  constraints: {
    required?: boolean;
    maxLen?: number;
    pattern?: string;
    min?: number;
    max?: number;
  },
  label: string,
): string | null {
  if (constraints.required && value.trim() === '') {
    return `${label} is required.`;
  }
  if (constraints.maxLen !== undefined && value.length > constraints.maxLen) {
    return `${label} must not exceed ${constraints.maxLen} characters.`;
  }
  if (constraints.pattern) {
    try {
      const re = new RegExp(constraints.pattern);
      if (!re.test(value)) {
        return `${label} does not match the required format.`;
      }
    } catch {
      // Malformed pattern in schema — skip pattern validation gracefully.
    }
  }
  if (constraints.min !== undefined) {
    const n = parseFloat(value);
    if (!isNaN(n) && n < constraints.min) {
      return `${label} must be at least ${constraints.min}.`;
    }
  }
  if (constraints.max !== undefined) {
    const n = parseFloat(value);
    if (!isNaN(n) && n > constraints.max) {
      return `${label} must be at most ${constraints.max}.`;
    }
  }
  return null;
}
