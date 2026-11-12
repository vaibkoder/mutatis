import type { Severity } from '@shared/findings/types.js';
import type { SectionNode } from '@shared/section-index/types.js';
import type { TextAnchor } from '@shared/doc-model/anchor.js';

/** Unit types that can be cross-referenced. Mirrors SectionKind + 'paragraph'. */
export type UnitType =
  | 'clause'
  | 'schedule'
  | 'paragraph'
  | 'annex'
  | 'appendix'
  | 'part'
  | 'recital';

/** Classification of a resolved cross-reference. */
export type RefClass = 'ok' | 'notFound' | 'ambiguous' | 'externalMaybe';

/** A prose reference match produced by the reference scanner. */
export interface ReferenceMatch {
  matchId: string;
  /** Exact matched prose, e.g. "Clause 4.2", "paragraph 3 of Schedule 2". */
  rawText: string;
  anchor: TextAnchor;
  paragraphId: string;
  unitType: UnitType;
  /** Normalised key, e.g. "clause:4.2" — resolved against the Section Index. */
  normalizedKey: string;
  /** Expanded keys for ranges/lists, e.g. "Schedules 2 to 5" → [schedule:2..schedule:5]. */
  expandedKeys: string[];
  classification: RefClass;
  /** Targets for Ambiguous finding — the >1 matching nodes. Also seeds the picker. */
  candidateTargets: SectionNode[];
  /** True when the match is backed by a REF/PAGEREF Word field. */
  isFieldBacked: boolean;
  fieldId?: string;
  /** True when not field-backed — "manually typed, may be out of date". */
  staticBadge: boolean;
  severity: Severity;
}

/** Verdict from Feature B three-layer heuristic. */
export type NumberingVerdict = 'dynamic' | 'static' | 'noNumber' | 'indeterminate';

/** Per-paragraph result from Feature B. */
export interface StaticNumberingResult {
  paragraphId: string;
  leadingNumberText: string | null;
  hasListItem: boolean;
  listStringIfAny: string | null;
  /** Only set when layer (c) OOXML probe was run. */
  ooxmlVerdict?: 'tRun' | 'listOrField' | 'mixed';
  verdict: NumberingVerdict;
  confidence: 'high' | 'medium';
  severity: Severity;
}
