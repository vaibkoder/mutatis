import type { ModelChoice, ProviderId } from '@common/ipc.js';

/**
 * Parsing persisted settings.
 *
 * Separate from the provider so it can be tested without React or a DOM, and
 * because it is the place that decides what a corrupt or outdated stored
 * preference does. The answer is always: fall back to the default for the field
 * that is wrong, keep the fields that are fine, and never throw. A user whose
 * localStorage was written by an older build must still get a working window.
 */

export type Jurisdiction = 'UK' | 'SG';

export interface Settings {
  readonly jurisdiction: Jurisdiction;
  readonly model: ModelChoice;
}

export const DEFAULT_SETTINGS: Settings = {
  jurisdiction: 'UK',
  model: { provider: 'deepseek', model: 'deepseek-v4-flash' },
};

function isProviderId(value: unknown): value is ProviderId {
  return value === 'deepseek' || value === 'alibaba';
}

export function parseSettings(raw: string | null): Settings {
  if (raw === null || raw === '') return DEFAULT_SETTINGS;

  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return DEFAULT_SETTINGS;
  }
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    return DEFAULT_SETTINGS;
  }

  const record = parsed as Record<string, unknown>;

  // Anything that is not 'SG' means UK, including a jurisdiction this build has
  // never heard of. Defaulting to a profile rather than refusing is right here:
  // the alternative is a window that will not open because of a stale string.
  const jurisdiction: Jurisdiction = record['jurisdiction'] === 'SG' ? 'SG' : 'UK';

  const stored = record['model'];
  const model: ModelChoice =
    typeof stored === 'object' &&
    stored !== null &&
    isProviderId((stored as Record<string, unknown>)['provider']) &&
    typeof (stored as Record<string, unknown>)['model'] === 'string' &&
    (stored as Record<string, unknown>)['model'] !== ''
      ? {
          provider: (stored as Record<string, unknown>)['provider'] as ProviderId,
          model: (stored as Record<string, unknown>)['model'] as string,
        }
      : DEFAULT_SETTINGS.model;

  return { jurisdiction, model };
}
