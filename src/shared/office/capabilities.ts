/**
 * Capabilities module — the single source of truth for which Office JS requirement
 * sets are available in the current host. Feature modules check here; they never
 * throw when a capability is absent; they degrade gracefully.
 * (spec/ARCHITECTURE.md §Cross-Platform Matrix)
 */

export interface CapabilityMap {
  wordApi15: boolean;  // footnotes, TOC field insert
  wordApi16: boolean;  // getTrackedChanges().acceptAll/rejectAll
  ooxmlLoop: boolean;  // getFileAsync + insertFileFromBase64
  /** WordApi 1.3 — listItem.listString, Range.paragraphs, Body.search (feature 04). */
  wordApi13: boolean;
  /** True on a desktop host (Windows/Mac), false on Word Online / web / unknown (feature 04b). */
  isDesktop: boolean;
}

let _caps: CapabilityMap | null = null;

/** Detect and cache capability flags. Call once after Office.onReady. */
export function detectCapabilities(): CapabilityMap {
  if (_caps) return _caps;

  const isSupported = (name: string, version: string): boolean => {
    try {
      return typeof Office !== 'undefined' &&
        Office.context.requirements.isSetSupported(name, version);
    } catch {
      return false;
    }
  };

  _caps = {
    wordApi13: isSupported('WordApi', '1.3'),
    wordApi15: isSupported('WordApi', '1.5'),
    wordApi16: isSupported('WordApi', '1.6'),
    // OOXML loop is available when the typed API is available; degraded on constrained hosts.
    ooxmlLoop: isSupported('WordApi', '1.1'),
    isDesktop: detectIsDesktop(),
  };

  return _caps;
}

/**
 * True on a desktop host (Word for Windows/Mac), false on Word Online / web /
 * unknown. Guarded: returns false when `Office.context.platform` or the
 * `PlatformType` enum is unavailable (e.g. unit tests, constrained hosts).
 */
function detectIsDesktop(): boolean {
  try {
    if (typeof Office === 'undefined') return false;
    const platform = Office.context?.platform;
    const T = Office.PlatformType;
    if (!platform || !T) return false;
    return platform === T.PC || platform === T.Mac;
  } catch {
    return false;
  }
}

/** Get the cached capability map. Returns all-false if called before detectCapabilities(). */
export function getCapabilities(): CapabilityMap {
  return _caps ?? {
    wordApi13: false, wordApi15: false, wordApi16: false, ooxmlLoop: false, isDesktop: false,
  };
}

/** True if a named capability is available. */
export function isCapable(key: keyof CapabilityMap): boolean {
  return getCapabilities()[key];
}

/** Reset the cache (for testing). */
export function _resetCapabilities(): void {
  _caps = null;
}
