/**
 * Feature 06 — Numbering: scheme storage + JSON export / import (M5).
 *
 * Spec §7.3:
 *  - Machine library → localStorage via the Storage abstraction.
 *  - Document-pinned (optional) → document.settings.
 *  - JSON export: serialise one or all user schemes to a downloadable Blob.
 *    No document content — only numbering definitions.
 *  - JSON import: validate schemaVersion and structure; collision handling;
 *    reject malformed JSON, never partial-import.
 *
 * Built-in preset ids are RESERVED — a user scheme importing with a built-in id
 * gets a fresh uuid-like id (timestamp-based; deterministic in tests if seeded).
 */

import type { NumberingScheme } from './types.js';
import type { CollisionResolution } from './types.js';
import { validateScheme, BUILTIN_IDS, SCHEMA_VERSION, duplicateScheme } from './schema.js';
import { getStorageProvider } from '@shared/storage/index.js';

/** Read a Blob as text using FileReader (works in jsdom + browsers; blob.text() not universal). */
function readBlobText(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = () => reject(reader.error ?? new Error('FileReader failed'));
    reader.readAsText(blob);
  });
}

const LIBRARY_KEY_PREFIX = 'numbering.scheme.';
const DOC_KEY_PREFIX = 'numbering.scheme.doc.';

// ---------------------------------------------------------------------------
// Key helpers
// ---------------------------------------------------------------------------

function libraryKey(id: string): string {
  return `${LIBRARY_KEY_PREFIX}${id}`;
}

function docKey(id: string): string {
  return `${DOC_KEY_PREFIX}${id}`;
}

/** Generate a timestamp-based id for imported/duplicated schemes. */
function newId(): string {
  return `user-${Date.now()}-${Math.floor(Math.random() * 10000)}`;
}

// ---------------------------------------------------------------------------
// Save / load (machine library)
// ---------------------------------------------------------------------------

/** Save a scheme to the machine-local library. Overwrites by id. */
export async function saveScheme(scheme: NumberingScheme): Promise<void> {
  const store = getStorageProvider().local;
  await store.set(libraryKey(scheme.id), scheme as unknown as Record<string, unknown>);
}

/** Load a scheme by id from the machine-local library. Returns undefined if absent. */
export async function loadScheme(id: string): Promise<NumberingScheme | undefined> {
  const store = getStorageProvider().local;
  const raw = await store.get(libraryKey(id));
  if (!raw) return undefined;
  return raw as unknown as NumberingScheme;
}

/** Load all user schemes from the machine-local library (excludes built-in presets). */
export async function listUserSchemes(): Promise<NumberingScheme[]> {
  const store = getStorageProvider().local;
  const keys = await store.list(LIBRARY_KEY_PREFIX);
  const schemes: NumberingScheme[] = [];
  for (const key of keys) {
    const raw = await store.get(key);
    if (raw) schemes.push(raw as unknown as NumberingScheme);
  }
  return schemes;
}

/** Delete a scheme from the machine-local library. No-op if absent. */
export async function deleteScheme(id: string): Promise<void> {
  const store = getStorageProvider().local;
  await store.set(libraryKey(id), undefined as unknown as Record<string, unknown>);
}

// ---------------------------------------------------------------------------
// Document-pinned schemes (document.settings)
// ---------------------------------------------------------------------------

/** Pin a scheme to the current document (stored in document settings). */
export async function pinSchemeToDocument(scheme: NumberingScheme): Promise<void> {
  const store = getStorageProvider().docSettings;
  await store.set(docKey(scheme.id), scheme as unknown as Record<string, unknown>);
}

// ---------------------------------------------------------------------------
// JSON export
// ---------------------------------------------------------------------------

export interface ExportEnvelope {
  schemaVersion: number;
  exportedAt: string;
  schemes: NumberingScheme[];
}

/**
 * Export one or all user schemes as a downloadable JSON Blob.
 * Built-in preset content is included if explicitly listed — they are readable.
 * The export carries ONLY numbering definitions — NO document content.
 */
export function exportSchemes(schemes: readonly NumberingScheme[]): Blob {
  const envelope: ExportEnvelope = {
    schemaVersion: SCHEMA_VERSION,
    exportedAt: new Date().toISOString(),
    schemes: schemes.map(s => ({ ...s })),
  };
  const json = JSON.stringify(envelope, null, 2);
  return new Blob([json], { type: 'application/json' });
}

// ---------------------------------------------------------------------------
// JSON import
// ---------------------------------------------------------------------------

export interface ImportResult {
  imported: NumberingScheme[];
  replaced: string[];   // ids of existing schemes that were overwritten
  errors: string[];
}

/**
 * Import schemes from a JSON Blob.
 *
 * Collision handling (spec §7.3):
 *  - Incoming id matches a BUILT-IN → assign a fresh id (never import over a built-in).
 *  - Incoming id matches a USER LIBRARY entry → apply `resolution` (replace | keepBoth).
 *  - No collision → import as-is.
 *
 * Rejects malformed JSON with an error list; never partially imports.
 * All-or-nothing: if ANY scheme fails validation, the entire import is aborted.
 */
export async function importSchemes(
  blob: Blob,
  resolution: CollisionResolution = 'replace',
  builtInPresets: readonly NumberingScheme[] = [],
): Promise<ImportResult> {
  const errors: string[] = [];
  let envelope: ExportEnvelope;

  // Parse JSON (use FileReader for jsdom/browser compat — blob.text() not universal).
  let text: string;
  try {
    text = await readBlobText(blob);
  } catch {
    return { imported: [], replaced: [], errors: ['Could not read the file.'] };
  }

  try {
    envelope = JSON.parse(text) as ExportEnvelope;
  } catch {
    return { imported: [], replaced: [], errors: ['File is not valid JSON.'] };
  }

  // Schema version gate.
  if (typeof envelope.schemaVersion !== 'number') {
    return { imported: [], replaced: [], errors: ['Missing or invalid schemaVersion in file.'] };
  }
  if (envelope.schemaVersion > SCHEMA_VERSION) {
    return {
      imported: [],
      replaced: [],
      errors: [`File uses schemaVersion ${envelope.schemaVersion}; this add-in supports up to ${SCHEMA_VERSION}.`],
    };
  }

  if (!Array.isArray(envelope.schemes) || envelope.schemes.length === 0) {
    return { imported: [], replaced: [], errors: ['File contains no schemes.'] };
  }

  // Validate all schemes BEFORE saving anything (all-or-nothing).
  const candidates: NumberingScheme[] = [];
  for (const raw of envelope.schemes) {
    const scheme = raw as NumberingScheme;
    const validation = validateScheme(scheme);
    if (!validation.valid) {
      errors.push(`Scheme "${scheme.name ?? '?'}" is invalid: ${validation.errors.join('; ')}`);
    } else {
      candidates.push(scheme);
    }
  }
  if (errors.length > 0) {
    return { imported: [], replaced: [], errors };
  }

  // Resolve collisions and save.
  const imported: NumberingScheme[] = [];
  const replaced: string[] = [];
  const builtInIds = new Set(BUILTIN_IDS);
  for (const bi of builtInPresets) builtInIds.add(bi.id);

  for (const scheme of candidates) {
    // Ensure imported schemes are never marked builtIn.
    const base: NumberingScheme = { ...scheme };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (base as any).builtIn = false;
    let finalScheme = base;

    // Built-in collision → fresh id.
    if (builtInIds.has(scheme.id)) {
      finalScheme = duplicateScheme(finalScheme, newId());
    } else {
      // Check user library collision.
      const existing = await loadScheme(scheme.id);
      if (existing) {
        if (resolution === 'keepBoth') {
          finalScheme = duplicateScheme(finalScheme, newId());
        } else {
          replaced.push(scheme.id);
        }
      }
    }

    await saveScheme(finalScheme);
    imported.push(finalScheme);
  }

  return { imported, replaced, errors };
}
