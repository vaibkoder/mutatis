/**
 * host-web — storage provider (ADR-0022, spec/features/12 §3 phase 2).
 *
 * The web target has no `Office.context.document.settings`, so document-scoped
 * settings are backed by the same client-side localStorage store as machine-local
 * preferences. Both are `createLocalStore()` (host-free); nothing leaves the
 * machine — the confidentiality invariant holds by construction (ADR-0001/0003).
 *
 * A later phase may give `docSettings` a per-document-keyed namespace; for the
 * factual-consistency PoC a single localStorage-backed store is sufficient.
 */
import { createLocalStore } from '@shared/storage/index.js';
import type { StorageProvider } from '@shared/storage/index.js';

/** Build the web `StorageProvider` (localStorage for both stores). */
export function createWebStorageProvider(): StorageProvider {
  return {
    local: createLocalStore(),
    docSettings: createLocalStore(),
  };
}
