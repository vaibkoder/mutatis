import type { SectionNode } from '@shared/section-index/types.js';
import type { TextAnchor } from '@shared/doc-model/anchor.js';

export type { TextAnchor };

// ── Cross-reference descriptor ───────────────────────────────────────────────

/** Section kind mapped to the Section Index key grammar. */
export interface RefDescriptor {
  kind: 'clause' | 'schedule' | 'schedule-paragraph' | 'recital' | 'annexure' | 'para-only';
  /** Top-level number/letter: "14.2" | "2" | "(B)". */
  primary: string;
  /** For schedule-paragraph: the paragraph number within the schedule. */
  secondary?: string;
}

// ── Classification ────────────────────────────────────────────────────────────

export type LookupClassificationType =
  | 'CROSS_REF'        // plain Clause/Schedule/Recital/Annexure
  | 'SCHED_PARA'       // compound "paragraph N of Schedule M" — tested first
  | 'EXTERNAL'         // "Clause X of the Facility Agreement"
  | 'DEFINED_TERM'     // key exists in the dictionary
  | 'NEITHER';         // idle

export interface LookupClassification {
  type: LookupClassificationType;
  /** Set for CROSS_REF and SCHED_PARA. */
  descriptor?: RefDescriptor;
  /** Set for DEFINED_TERM: normalised lookup key. */
  termKey?: string;
  /** Set for EXTERNAL: captured document name. */
  docName?: string;
  /** The candidate phrase that was classified. */
  rawText: string;
}

// ── Defined-terms dictionary ──────────────────────────────────────────────────

export interface DefinedTerm {
  /** Normalised lookup key (lowercased, NFC, collapsed ws). */
  termKey: string;
  /** Surface form as defined, e.g. "Completion Date". */
  display: string;
  /** Extracted definition text trimmed to previewMaxChars. */
  previewText: string;
  defKind: 'inline' | 'by-cross-ref' | 'external';
  /** Set when defKind === "by-cross-ref": raw cross-ref to follow one hop, e.g. "Clause 6.1". */
  forwardRef?: string;
  /** Index into DocModel.allParagraphs at build time. */
  sourceParagraphIndex: number;
  /** For Jump-to navigation via resolveAnchor. */
  anchor: TextAnchor;
}

export interface TermsDictionary {
  /** array > 1 entry → term defined multiple times → ambiguous. */
  byKey: Map<string, DefinedTerm[]>;
  /** True between a DocModel invalidation and the debounced rebuild finishing. */
  dirty: boolean;
  /**
   * Longest-match-first alternation over every defined surface form, built once
   * at dictionary build. Used to locate the defined term *under a bare cursor*
   * within the sentence window (spec §4.2 / T3), when the caret sits inside the
   * term rather than the whole term being selected. `null` when no terms exist.
   * Optional so hand-built test dictionaries need not supply it.
   */
  termMatcher?: RegExp | null;
}

// ── Lookup result union ───────────────────────────────────────────────────────

/** A cross-reference that resolved to exactly one Section Index node. */
export interface LookupResolved {
  type: 'resolved';
  label: string;
  previewText: string;
  node: SectionNode;
}

/** A defined term that resolves to a unique inline/external definition. */
export interface LookupDefinition {
  type: 'definition';
  term: DefinedTerm;
}

/** A term defined by cross-reference: show definition sentence + optionally the resolved clause. */
export interface LookupDefByXref {
  type: 'def-by-xref';
  term: DefinedTerm;
  /** The resolved forward-ref clause, if the one-hop Section Index lookup succeeded. */
  resolvedClause?: LookupResolved;
}

/** An external-document reference (Office JS cannot read the other file). */
export interface LookupExternal {
  type: 'external';
  docName: string;
}

/** A pick-list entry for the ambiguous state. */
export type AmbiguousCandidate =
  | { kind: 'section'; label: string; containerBadge: string; previewText: string; node: SectionNode }
  | { kind: 'term';    label: string; containerBadge: string; previewText: string; term: DefinedTerm };

/** Multiple possible targets — user picks from a list. Best-guess is expanded first. */
export interface LookupAmbiguous {
  type: 'ambiguous';
  candidates: AmbiguousCandidate[];
  /** Index into candidates that the proximity heuristic selected as best guess. */
  bestGuessIndex: number;
}

/** Cross-ref or term-shaped phrase with 0 matches. */
export interface LookupNoDefinition {
  type: 'no-definition';
  phrase: string;
}

/** Selection classified as NEITHER. Shown as idle state. */
export interface LookupIdle {
  type: 'idle';
}

/**
 * A lookup that completed while the dictionary was dirty.
 * Wraps the stale result; the panel shows it with an "updating definitions…" chip.
 */
export interface LookupUpdating {
  type: 'updating';
  stale: Exclude<LookupResult, LookupUpdating>;
}

export type LookupResult =
  | LookupResolved
  | LookupDefinition
  | LookupDefByXref
  | LookupExternal
  | LookupAmbiguous
  | LookupNoDefinition
  | LookupIdle
  | LookupUpdating;
