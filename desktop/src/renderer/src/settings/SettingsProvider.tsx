import { createContext, useCallback, useContext, useEffect, useMemo, useState } from 'react';
import type { KeyStatus, ModelChoice, ProviderId, ProvidersSnapshot } from '@common/ipc.js';
import { DEFAULT_SETTINGS, parseSettings, type Jurisdiction, type Settings } from './parse.js';

/**
 * App settings that are not secrets.
 *
 * Keys are NOT here and never will be — they live in the main process behind
 * safeStorage. What lives here is the ordinary preference state the renderer is
 * allowed to own: which model to talk to, and which jurisdiction profile the
 * compare engine runs under.
 *
 * Persisted to localStorage, every access wrapped, because a renderer that
 * cannot read a preference must still start. `app://` is a secure context so
 * storage is reliable here — but a cleared profile or a private-mode-like
 * restriction must degrade to defaults, not to a blank window.
 */

export type { Jurisdiction, Settings };
export { DEFAULT_SETTINGS };

const STORAGE_KEY = 'mutatis.settings.v1';

interface SettingsContextValue extends Settings {
  setJurisdiction: (value: Jurisdiction) => void;
  setModel: (value: ModelChoice) => void;

  /**
   * The provider registry and which keys exist, held HERE rather than fetched
   * per component. Two copies would disagree the moment settings stores a key
   * and the chat panel is still showing its gate.
   */
  readonly snapshot: ProvidersSnapshot | null;
  readonly keys: KeyStatus | null;
  readonly loadFailed: boolean;
  refreshKeys: () => Promise<void>;
  /** Store a key, or remove it by passing an empty string. */
  saveKey: (provider: ProviderId, key: string) => Promise<{ ok: boolean; error?: string }>;
}

const SettingsContext = createContext<SettingsContextValue | null>(null);

function read(): Settings {
  try {
    return parseSettings(window.localStorage.getItem(STORAGE_KEY));
  } catch {
    // Reading localStorage can THROW, not just return null, when site data is
    // blocked. A preference we cannot read must not stop the app starting.
    return DEFAULT_SETTINGS;
  }
}

function write(settings: Settings): void {
  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
  } catch {
    // A preference that cannot be saved is an annoyance. Throwing here would
    // make it a crash.
  }
}

export function SettingsProvider({ children }: { children: React.ReactNode }) {
  const [settings, setSettings] = useState<Settings>(read);
  const [snapshot, setSnapshot] = useState<ProvidersSnapshot | null>(null);
  const [keys, setKeys] = useState<KeyStatus | null>(null);
  const [loadFailed, setLoadFailed] = useState(false);

  useEffect(() => {
    write(settings);
  }, [settings]);

  useEffect(() => {
    let live = true;
    void (async () => {
      const [providers, status] = await Promise.all([
        window.mutatis.providers(),
        window.mutatis.secrets.status(),
      ]);
      if (!live) return;
      if (!providers || !status) {
        setLoadFailed(true);
        return;
      }
      setSnapshot(providers);
      setKeys(status);
    })();
    return () => {
      live = false;
    };
  }, []);

  const refreshKeys = useCallback(async () => {
    const status = await window.mutatis.secrets.status();
    if (status) setKeys(status);
  }, []);

  const saveKey = useCallback(
    async (provider: ProviderId, key: string) => {
      const result = await window.mutatis.secrets.set(provider, key);
      await refreshKeys();
      return result.ok
        ? { ok: true as const }
        : { ok: false as const, error: result.error ?? 'The key could not be saved.' };
    },
    [refreshKeys],
  );

  const setJurisdiction = useCallback((jurisdiction: Jurisdiction) => {
    setSettings((current) => ({ ...current, jurisdiction }));
  }, []);

  const setModel = useCallback((model: ModelChoice) => {
    setSettings((current) => ({ ...current, model }));
  }, []);

  const value = useMemo<SettingsContextValue>(
    () => ({
      ...settings,
      setJurisdiction,
      setModel,
      snapshot,
      keys,
      loadFailed,
      refreshKeys,
      saveKey,
    }),
    [settings, setJurisdiction, setModel, snapshot, keys, loadFailed, refreshKeys, saveKey],
  );

  return <SettingsContext value={value}>{children}</SettingsContext>;
}

export function useSettings(): SettingsContextValue {
  const value = useContext(SettingsContext);
  if (!value) throw new Error('useSettings must be used inside <SettingsProvider>');
  return value;
}
