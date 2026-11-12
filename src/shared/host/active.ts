/**
 * Active-host registry — the injection point for the current `HostPort`.
 *
 * Each app target installs its adapter once at startup (`setActiveHost(...)`),
 * and host-neutral UI/feature code reaches the host via `getActiveHost()`
 * rather than importing an adapter directly. This keeps the core free of any
 * compile-time dependency on a specific host (and free of `Office.*`/`Word.*`).
 *
 * (spec/features/12-standalone-implementation.md §3.)
 */
import type { HostPort } from './types.js';

let _host: HostPort | null = null;

/** Install the active host adapter. Called once at app startup. */
export function setActiveHost(host: HostPort): void {
  _host = host;
}

/**
 * Get the active host adapter.
 * @throws if no host has been installed yet (programming error — the app
 * target must call `setActiveHost` before any feature/UI code runs).
 */
export function getActiveHost(): HostPort {
  if (!_host) {
    throw new Error(
      'No active HostPort installed. The app target must call setActiveHost() at startup.',
    );
  }
  return _host;
}

/** Whether a host has been installed (lets UI degrade instead of throwing). */
export function hasActiveHost(): boolean {
  return _host !== null;
}

/** Reset the registry (for testing). */
export function _resetActiveHost(): void {
  _host = null;
}
