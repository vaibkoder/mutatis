/**
 * Resolving `compare.batch.*` (ADR-0037; spec §7).
 *
 * The keys travel on `CompareConfig`, so a batch setting is reached the same
 * way every other Compare setting is. The **defaults do not live there**:
 * `shared/config` must not depend on a feature, so the named constants stay in
 * `./constants.ts` and this module is the seam that joins the two.
 *
 * An unset key means "use the constant", never "use zero" — which is why every
 * field is resolved through `??` rather than spread. A spread would let an
 * explicit `poolSize: 0` and an absent `poolSize` become indistinguishable, and
 * for `poolSize` in particular `0` is a meaningful value (auto-size).
 *
 * Pure — no Office JS, no DOM, no network.
 */
import type { CompareConfig } from '@shared/config/index.js';
import {
  PAIRING_AUTO_PAIR_MARGIN,
  PAIRING_AUTO_PAIR_SCORE,
  PAIRING_MIN_SKETCH_TOKENS,
  PAIRING_NO_CONTENT_HEADROOM,
  PAIRING_NO_CONTENT_MATCH_SCORE,
  PAIRING_NO_MATCH_SCORE,
  PAIRING_TIE_EPSILON,
  PAIRING_WEIGHT_CONTENT,
  PAIRING_WEIGHT_FILENAME,
} from './constants.js';
import { MAX_BATCH_VARIANTS, MAX_BUNDLE_DOCUMENTS } from './limits.js';

/** Every batch setting, resolved. No optional fields — that is the point. */
export interface ResolvedBatchConfig {
  maxVariants: number;
  maxBundleDocuments: number;
  /** `0` = auto-size from `hardwareConcurrency`. */
  poolSize: number;
  pairing: {
    weightContent: number;
    weightFilename: number;
    tieEpsilon: number;
    autoPairScore: number;
    autoPairMargin: number;
    noMatchScore: number;
    /** Filename bar for the second, content-free assignment phase (§6.4). */
    noContentMatchScore: number;
    minSketchTokens: number;
  };
}

/**
 * A count: a non-negative integer, or the default.
 *
 * A `NaN` or a negative reaching the caps would not throw — it would silently
 * refuse every batch, or accept every batch, depending on which comparison it
 * landed in. Config comes from a profile a human edits, so it is validated at
 * the seam rather than trusted.
 */
function count(value: number | undefined, fallback: number): number {
  if (value === undefined) return fallback;
  return Number.isFinite(value) && value >= 0 ? Math.floor(value) : fallback;
}

/** A weight or threshold: a number in 0..1, or the default. */
function unit(value: number | undefined, fallback: number): number {
  if (value === undefined) return fallback;
  return Number.isFinite(value) && value >= 0 && value <= 1 ? value : fallback;
}

/**
 * The pairing settings alone, plus the one value derived from them.
 *
 * Split out because the pairing modules need these and nothing else: passing
 * the whole `ResolvedBatchConfig` down would let a scoring function reach for a
 * pool size, and a type that permits a nonsense dependency eventually gets one.
 */
export type PairingSettings = ResolvedBatchConfig['pairing'] & {
  /**
   * The most a pair with NO content evidence may score (see `score.ts`).
   *
   * Derived from the configured `noMatchScore` rather than configured
   * separately: its whole meaning is "just above the point at which a document
   * would rather go unmatched", so a profile that moves that threshold must
   * move this with it or the relationship the ceiling encodes is broken.
   */
  noContentCeiling: number;
};

/** The pairing settings of a resolved config, with the ceiling derived. */
export function pairingSettingsOf(resolved: ResolvedBatchConfig): PairingSettings {
  return {
    ...resolved.pairing,
    // Clamped: `noMatchScore` is a unit value, so a profile setting it to 1
    // would put the ceiling at 1.05 and the grid would report a pair at 105%.
    // A score is a proportion; nothing downstream is entitled to see one above
    // its own maximum just because a threshold was pushed to the edge.
    noContentCeiling: Math.min(1, resolved.pairing.noMatchScore + PAIRING_NO_CONTENT_HEADROOM),
  };
}

/**
 * Pairing settings from a `CompareConfig`.
 *
 * The pairing modules take this rather than reading the constants directly.
 * `spec/features/17-compare-batch.md` §7 requires `compare.batch.pairing.*` to
 * be resolved through the config layer, and a review found that although they
 * WERE resolved here, nothing consumed the result — every module imported the
 * raw constants, so setting `autoPairScore` in a profile did nothing at all.
 * A structural test in the acceptance sweep now forbids those imports outside
 * this file.
 */
export function resolvePairingSettings(config: CompareConfig): PairingSettings {
  return pairingSettingsOf(resolveBatchConfig(config));
}

/**
 * The settings a profile has said nothing about — the feature's own constants.
 *
 * Used as a default argument by the pairing leaves so a unit test can score a
 * pair without assembling a config. Production paths thread the resolved
 * settings from `proposePairing` instead; the acceptance sweep's import wall is
 * what keeps that true.
 */
export const DEFAULT_PAIRING_SETTINGS: PairingSettings = pairingSettingsOf({
  maxVariants: MAX_BATCH_VARIANTS,
  maxBundleDocuments: MAX_BUNDLE_DOCUMENTS,
  poolSize: 0,
  pairing: {
    weightContent: PAIRING_WEIGHT_CONTENT,
    weightFilename: PAIRING_WEIGHT_FILENAME,
    tieEpsilon: PAIRING_TIE_EPSILON,
    autoPairScore: PAIRING_AUTO_PAIR_SCORE,
    autoPairMargin: PAIRING_AUTO_PAIR_MARGIN,
    noMatchScore: PAIRING_NO_MATCH_SCORE,
    noContentMatchScore: PAIRING_NO_CONTENT_MATCH_SCORE,
    minSketchTokens: PAIRING_MIN_SKETCH_TOKENS,
  },
});

/** Resolve `compare.batch.*` against the feature's own constants. */
export function resolveBatchConfig(config: CompareConfig): ResolvedBatchConfig {
  const batch = config.batch;
  const pairing = batch?.pairing;
  return {
    maxVariants: count(batch?.maxVariants, MAX_BATCH_VARIANTS),
    maxBundleDocuments: count(batch?.maxBundleDocuments, MAX_BUNDLE_DOCUMENTS),
    poolSize: count(batch?.poolSize, 0),
    pairing: {
      weightContent: unit(pairing?.weightContent, PAIRING_WEIGHT_CONTENT),
      weightFilename: unit(pairing?.weightFilename, PAIRING_WEIGHT_FILENAME),
      tieEpsilon: unit(pairing?.tieEpsilon, PAIRING_TIE_EPSILON),
      autoPairScore: unit(pairing?.autoPairScore, PAIRING_AUTO_PAIR_SCORE),
      autoPairMargin: unit(pairing?.autoPairMargin, PAIRING_AUTO_PAIR_MARGIN),
      noMatchScore: unit(pairing?.noMatchScore, PAIRING_NO_MATCH_SCORE),
      noContentMatchScore: unit(pairing?.noContentMatchScore, PAIRING_NO_CONTENT_MATCH_SCORE),
      minSketchTokens: count(pairing?.minSketchTokens, PAIRING_MIN_SKETCH_TOKENS),
    },
  };
}
