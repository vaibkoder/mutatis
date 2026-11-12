/**
 * Shared Data Sets — C7.
 *
 * Save / load / push data sets via the Storage abstraction.
 * Personal sets go to localStorage; shared sets require a configured backend.
 *
 * Fuzzy matching (spec §C.5.4):
 *   - Exact match: normalised(dataSetKey) === normalised(schemaTag).
 *   - Fuzzy: Levenshtein-like similarity; warn before apply, never silent.
 *   - Unmatched keys: listed for manual mapping.
 *   - NEVER auto-apply a fuzzy match without user confirmation.
 */

import { getStorageProvider } from '@shared/storage/index.js';
import { fillRepeatingGroup } from './repeating.js';
import { fillQuestionnaire } from './fill.js';
import type {
  SharedDataSet, TemplateSchema, FuzzyMatchWarning, FuzzyPushPreview,
} from './types.js';
import { DATASET_KEY_PREFIX } from './types.js';

// ---------------------------------------------------------------------------
// Normalisation for fuzzy matching
// ---------------------------------------------------------------------------

/**
 * Normalise a key for comparison:
 *   CLIENT_NAME → client name
 *   clientName  → client name
 *   "client name" → client name
 */
function normaliseKey(key: string): string {
  return key
    .replace(/_/g, ' ')                         // snake_case → spaces
    .replace(/([a-z])([A-Z])/g, '$1 $2')        // camelCase → spaces
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

/**
 * Compute a simple character-overlap similarity score 0–1.
 * Not a full Levenshtein but sufficient for tag-name matching.
 */
function similarityScore(a: string, b: string): number {
  if (a === b) return 1;
  const shorter = a.length < b.length ? a : b;
  const longer  = a.length < b.length ? b : a;
  if (longer.length === 0) return 1;
  let matches = 0;
  for (const ch of shorter) {
    if (longer.includes(ch)) matches++;
  }
  // Penalise length difference.
  const overlap = matches / longer.length;
  const lenPenalty = 1 - Math.abs(a.length - b.length) / (a.length + b.length + 1);
  return (overlap + lenPenalty) / 2;
}

/** Minimum similarity threshold to propose a fuzzy match (0–1). */
const FUZZY_THRESHOLD = 0.6;

// ---------------------------------------------------------------------------
// Fuzzy preview (C-T9)
// ---------------------------------------------------------------------------

/**
 * Build a preview of what a scalar data set push would do.
 * Returns exact matches, fuzzy candidates (warn before apply), and unmatched keys.
 * NEVER applies anything — the caller must show warnings and get user confirmation.
 */
export function buildFuzzyPreview(
  dataSetValues: Record<string, string>,
  schema: TemplateSchema,
): FuzzyPushPreview {
  const schemaTagNorms = schema.variables.map(v => ({
    tag: v.tag,
    norm: normaliseKey(v.tag),
  }));

  const exact: Array<{ tag: string; value: string }> = [];
  const fuzzy: FuzzyMatchWarning[] = [];
  const unmatched: string[] = [];

  for (const [dsKey, value] of Object.entries(dataSetValues)) {
    const dsNorm = normaliseKey(dsKey);

    // 1. Exact match (normalised).
    const exactMatch = schemaTagNorms.find(s => s.norm === dsNorm);
    if (exactMatch) {
      exact.push({ tag: exactMatch.tag, value });
      continue;
    }

    // 2. Fuzzy match — find the best candidate above threshold.
    let bestScore = 0;
    let bestTag = '';
    for (const { tag, norm } of schemaTagNorms) {
      const score = similarityScore(dsNorm, norm);
      if (score > bestScore) {
        bestScore = score;
        bestTag = tag;
      }
    }

    if (bestScore >= FUZZY_THRESHOLD && bestTag) {
      fuzzy.push({ dataSetKey: dsKey, schemaTag: bestTag, score: bestScore });
    } else {
      unmatched.push(dsKey);
    }
  }

  return { exact, fuzzy, unmatched };
}

// ---------------------------------------------------------------------------
// Save / load / list / delete
// ---------------------------------------------------------------------------

/** Save a data set to the appropriate store. */
export async function saveDataSet(ds: SharedDataSet): Promise<void> {
  const store =
    ds.scope === 'personal'
      ? getStorageProvider().local
      : getStorageProvider().docSettings; // shared backend is stubbed in v1

  const key = `${DATASET_KEY_PREFIX}${ds.dataSetId}`;
  await store.set(key, ds as unknown);
}

/** Load a data set by id and scope. */
export async function loadDataSet(
  dataSetId: string,
  scope: 'personal' | 'shared',
): Promise<SharedDataSet | undefined> {
  const store =
    scope === 'personal'
      ? getStorageProvider().local
      : getStorageProvider().docSettings;

  const key = `${DATASET_KEY_PREFIX}${dataSetId}`;
  return store.get(key) as Promise<SharedDataSet | undefined>;
}

/** List all data set ids in a given scope. */
export async function listDataSets(scope: 'personal' | 'shared'): Promise<string[]> {
  const store =
    scope === 'personal'
      ? getStorageProvider().local
      : getStorageProvider().docSettings;

  const keys = await store.list(DATASET_KEY_PREFIX);
  return keys.map(k => k.slice(DATASET_KEY_PREFIX.length));
}

/** Delete a data set by id and scope. */
export async function deleteDataSet(
  dataSetId: string,
  scope: 'personal' | 'shared',
): Promise<void> {
  const store =
    scope === 'personal'
      ? getStorageProvider().local
      : getStorageProvider().docSettings;

  const key = `${DATASET_KEY_PREFIX}${dataSetId}`;
  await store.delete(key);
}

// ---------------------------------------------------------------------------
// Push a scalar data set
// ---------------------------------------------------------------------------

export interface ScalarPushRequest {
  /** The data set to push. */
  dataSet: SharedDataSet;
  /** The template schema to push into. */
  schema: TemplateSchema;
  /**
   * The confirmed fuzzy matches to apply (user has acknowledged the warnings).
   * Unmatched keys (not in here and not exact) are skipped.
   */
  confirmedFuzzy: FuzzyMatchWarning[];
}

export interface ScalarPushResult {
  written: string[];
  notFound: string[];
  /** Keys skipped because they were unmatched and the user did not confirm. */
  skipped: string[];
}

/**
 * Push a scalar data set into matching template content controls.
 * Exact matches are applied directly; fuzzy matches only if confirmed.
 */
export async function pushScalarDataSet(req: ScalarPushRequest): Promise<ScalarPushResult> {
  if (req.dataSet.kind !== 'scalar' || !req.dataSet.values) {
    return { written: [], notFound: [], skipped: [] };
  }

  const preview = buildFuzzyPreview(req.dataSet.values, req.schema);

  // Build the confirmed fuzzy map.
  const confirmedFuzzyTags = new Set(req.confirmedFuzzy.map(f => f.dataSetKey));

  const fillItems = [];

  // Exact matches — always apply.
  for (const { tag, value } of preview.exact) {
    const variable = req.schema.variables.find(v => v.tag === tag);
    fillItems.push({ tag, type: variable?.type ?? 'text', rawValue: value });
  }

  // Confirmed fuzzy matches.
  for (const warning of preview.fuzzy) {
    if (confirmedFuzzyTags.has(warning.dataSetKey)) {
      const variable = req.schema.variables.find(v => v.tag === warning.schemaTag);
      const value = req.dataSet.values?.[warning.dataSetKey] ?? '';
      fillItems.push({ tag: warning.schemaTag, type: variable?.type ?? 'text', rawValue: value });
    }
  }

  const fillResult = await fillQuestionnaire(fillItems);

  const skipped = preview.unmatched.concat(
    preview.fuzzy
      .filter(w => !confirmedFuzzyTags.has(w.dataSetKey))
      .map(w => w.dataSetKey),
  );

  return {
    written: fillResult.written,
    notFound: fillResult.notFound,
    skipped,
  };
}

// ---------------------------------------------------------------------------
// Push a repeating data set
// ---------------------------------------------------------------------------

export interface RepeatingPushResult {
  ok: boolean;
  capError: string | undefined;
  written: string[];
  notFound: string[];
  hiddenSlots: string[];
  errors: string[];
}

/**
 * Push a repeating data set into the matching fixed-slot group.
 */
export async function pushRepeatingDataSet(
  dataSet: SharedDataSet,
  schema: TemplateSchema,
): Promise<RepeatingPushResult> {
  if (dataSet.kind !== 'repeating' || !dataSet.rows) {
    return {
      ok: false,
      capError: 'Data set is not a repeating set or has no rows.',
      written: [], notFound: [], hiddenSlots: [], errors: [],
    };
  }

  if (!dataSet.groupRef) {
    return {
      ok: false,
      capError: 'Repeating data set has no groupRef.',
      written: [], notFound: [], hiddenSlots: [], errors: [],
    };
  }

  const group = schema.repeating.find(g => g.group === dataSet.groupRef);
  if (!group) {
    return {
      ok: false,
      capError: `No repeating group "${dataSet.groupRef}" found in template schema.`,
      written: [], notFound: [], hiddenSlots: [], errors: [],
    };
  }

  const result = await fillRepeatingGroup(group, dataSet.rows);
  return {
    ok: result.ok,
    capError: result.capError ?? undefined,
    written: result.written,
    notFound: result.notFound,
    hiddenSlots: result.hiddenSlots,
    errors: result.errors,
  };
}
