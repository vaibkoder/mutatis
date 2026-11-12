import { createContext, useCallback, useContext, useEffect, useMemo, useState } from 'react';
import { DEFAULT_SETTINGS, parseSettings, type Jurisdiction } from '../renderer/src/settings/parse.js';

/**
 * Settings for the single-file build: the jurisdiction and nothing else.
 *
 * The shared compare hook (compare/useComparisons.ts) imports `useSettings` from
 * renderer/src/settings/SettingsProvider.tsx, which talks to window.mutatis
 * and so cannot run in a browser. vite.web.config.ts redirects that import to
 * THIS module, so both the hook and the web settings dialog read one context.
 * Only the fields the web build uses are exposed; there is no model, no key.
 *
 * Stored under its own key so a file opened from file:// never reads (or
 * overwrites) anything an Electron profile wrote.
 */

export type { Jurisdiction };

const STORAGE_KEY = 'mutatis.web.settings.v1';

interface WebSettingsValue {
  readonly jurisdiction: Jurisdiction;
  setJurisdiction: (value: Jurisdiction) => void;
}

const SettingsContext = createContext<WebSettingsValue | null>(null);

function read(): Jurisdiction {
  try {
    return parseSettings(window.localStorage.getItem(STORAGE_KEY)).jurisdiction;
  } catch {
    return DEFAULT_SETTINGS.jurisdiction;
  }
}

export function SettingsProvider({ children }: { children: React.ReactNode }) {
  const [jurisdiction, setJurisdictionState] = useState<Jurisdiction>(read);

  useEffect(() => {
    try {
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify({ jurisdiction }));
    } catch {
      /* a preference that cannot be saved is not worth a crash */
    }
  }, [jurisdiction]);

  const setJurisdiction = useCallback((next: Jurisdiction) => setJurisdictionState(next), []);

  const value = useMemo(() => ({ jurisdiction, setJurisdiction }), [jurisdiction, setJurisdiction]);
  return <SettingsContext value={value}>{children}</SettingsContext>;
}

export function useSettings(): WebSettingsValue {
  const value = useContext(SettingsContext);
  if (!value) throw new Error('useSettings must be used inside <SettingsProvider>');
  return value;
}
