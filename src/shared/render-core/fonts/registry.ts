/**
 * Font registry — requested-family → OFL metric-clone mapping + face resolution.
 * Host-free, DOM-free. Font BYTES are injected by the caller (the worker
 * bootstrap / a vendored-asset module in Phase 2), so the engine stays pure and
 * unit-testable; this module never fetches or imports binary assets itself.
 *
 * Clone map (ADR-0024 §Decision / brief §4.2): the clones are advance-width
 * identical to the MS originals (proven in Spike A: Liberation Sans == Arial to
 * 0.000%), so line breaks computed on the clone match the original font.
 */
import type { EffectiveRPr } from '../types.js';
import { readLineMetrics, type LineMetrics } from './sfnt-metrics.js';

/** Normalised requested family → vendored clone family. */
export const CLONE_MAP: ReadonlyMap<string, string> = new Map([
  ['times new roman', 'Liberation Serif'],
  ['times', 'Liberation Serif'],
  ['timesnewroman', 'Liberation Serif'],
  ['arial', 'Liberation Sans'],
  ['helvetica', 'Liberation Sans'],
  ['calibri', 'Carlito'],
  ['cambria', 'Caladea'],
  ['courier new', 'Liberation Mono'],
  ['courier', 'Liberation Mono'],
]);

/** Families with NO metric clone — substitute visually, flag as non-metric. */
const NO_CLONE_VISUAL: ReadonlyMap<string, string> = new Map([
  ['garamond', 'Liberation Serif'], // EB Garamond is a visual-only match; Liberation Serif keeps layout sane
  ['georgia', 'Liberation Serif'],
]);

/** Default fallbacks when a family is entirely unknown. */
const DEFAULT_SERIF = 'Liberation Serif';
const DEFAULT_SANS = 'Liberation Sans';

/** One injected font file. `family` is the CLONE family name (e.g. 'Liberation Serif'). */
export interface FaceInput {
  family: string;
  bold: boolean;
  italic: boolean;
  /** Raw sfnt bytes (TTF/OTF) — used for shaping, metrics, and PDF subset embedding. */
  bytes: Uint8Array;
  /** True for an embedded .odttf-extracted face (wins over a clone of the same family). */
  embedded?: boolean;
}

export interface ResolvedFace {
  /** The clone/substitute family actually used. */
  family: string;
  /** The original requested family (for diagnostics / non-metric flagging). */
  requested: string;
  bold: boolean;
  italic: boolean;
  isMetricClone: boolean;
  bytes: Uint8Array;
  metrics: LineMetrics;
}

const norm = (s: string): string => s.trim().toLowerCase().replace(/\s+/g, ' ');
const faceKey = (family: string, bold: boolean, italic: boolean): string =>
  `${norm(family)}|${bold ? 'b' : ''}${italic ? 'i' : ''}`;

export interface FontResolver {
  resolve(rpr: EffectiveRPr): ResolvedFace;
  /** All distinct faces that have been resolved (for PDF embedding planning). */
  used(): ResolvedFace[];
}

/**
 * Build a resolver over the injected faces. Embedded faces (from .odttf) take
 * precedence over clones of the same family.
 */
export function createFontResolver(faces: readonly FaceInput[]): FontResolver {
  // Index faces by (family, style). Embedded wins over clone on collision.
  const byKey = new Map<string, FaceInput>();
  const families = new Set<string>();
  for (const f of faces) {
    const k = faceKey(f.family, f.bold, f.italic);
    const prev = byKey.get(k);
    if (!prev || (f.embedded && !prev.embedded)) byKey.set(k, f);
    families.add(norm(f.family));
  }
  const metricsCache = new Map<FaceInput, LineMetrics>();
  const usedSet = new Map<string, ResolvedFace>();

  function pickFamily(requested: string): { family: string; isMetricClone: boolean } {
    const n = norm(requested);
    // Embedded/exact family present? use it (metric-true by definition).
    if (families.has(n)) return { family: requested, isMetricClone: true };
    const clone = CLONE_MAP.get(n);
    if (clone && families.has(norm(clone))) return { family: clone, isMetricClone: true };
    const visual = NO_CLONE_VISUAL.get(n);
    if (visual && families.has(norm(visual))) return { family: visual, isMetricClone: false };
    // Heuristic serif/sans fallback by name.
    const looksSans = /(arial|helvetica|sans|calibri|verdana|tahoma|segoe)/.test(n);
    const fam = looksSans ? DEFAULT_SANS : DEFAULT_SERIF;
    return { family: families.has(norm(fam)) ? fam : (faces[0]?.family ?? fam), isMetricClone: false };
  }

  function findFace(family: string, bold: boolean, italic: boolean): FaceInput | undefined {
    return (
      byKey.get(faceKey(family, bold, italic)) ??
      byKey.get(faceKey(family, bold, false)) ??
      byKey.get(faceKey(family, false, italic)) ??
      byKey.get(faceKey(family, false, false))
    );
  }

  return {
    resolve(rpr: EffectiveRPr): ResolvedFace {
      const { family, isMetricClone } = pickFamily(rpr.fontFamily);
      const face = findFace(family, rpr.bold, rpr.italic) ?? faces[0];
      if (!face) throw new Error('FontResolver: no faces provided');
      let metrics = metricsCache.get(face);
      if (!metrics) {
        metrics = readLineMetrics(face.bytes);
        metricsCache.set(face, metrics);
      }
      const resolved: ResolvedFace = {
        family: face.family,
        requested: rpr.fontFamily,
        bold: face.bold,
        italic: face.italic,
        isMetricClone,
        bytes: face.bytes,
        metrics,
      };
      usedSet.set(faceKey(face.family, face.bold, face.italic), resolved);
      return resolved;
    },
    used(): ResolvedFace[] {
      return [...usedSet.values()];
    },
  };
}
