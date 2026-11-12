import type { TextAnchor } from '@shared/doc-model/anchor.js';

export type { TextAnchor };

/** Where a definition was found. */
export type DefinitionSource = 'prose' | 'table' | 'footnote';

/**
 * Recognition shape that produced a definition (ADR-0013):
 *  A       — quoted term + verb
 *  A-prime — quoted term + colon
 *  B       — inline parenthetical
 *  C       — unquoted candidate + mandatory verb
 *  table   — table structure (no verb)
 */
export type DefinitionShape = 'A' | 'A-prime' | 'B' | 'C' | 'table';

export interface DefinitionRecord {
  source: DefinitionSource;
  /** The matched verb, or the structural markers 'cross-reference'/'cross-document', or null (colon/table). */
  verb: string | 'cross-reference' | 'cross-document' | null;
  shape: DefinitionShape;
  /** "has the meaning given in Clause X" — same document. */
  isCrossRef: boolean;
  /** "as defined in the [Other Agreement]" — external; suppress, never resolve. */
  isCrossDoc: boolean;
  /** e.g. "Clause 14.2" or the named other agreement. */
  crossRefTarget: string | null;
  /** Anchor to the term's quoted/candidate span, for jump-to-navigation. */
  anchor: TextAnchor;
  /** Best-effort; never load-bearing (ParaNode exposes no run styling). */
  styling: { bold: boolean; italic: boolean };
  /** Delimiter actually used around the term. */
  rawQuote: 'double' | 'single' | 'none';
  /** e.g. "for the purposes of Clause 7" (local redefinition) — spec §8. */
  contextScope: string | null;
}

export interface DefinitionEntry {
  /** Normalised lowercase key. */
  key: string;
  /** Surface form as first authoritatively defined, e.g. "Completion Date". */
  canonical: string;
  /** >1 substantive record ⇒ defined-multiple-times candidate. */
  definitions: DefinitionRecord[];
}

export type DefinitionsMap = Map<string, DefinitionEntry>;

export interface UsageRecord {
  /** Exact surface text matched, e.g. "Completion Date" / "completion date". */
  surface: string;
  /** Surface form begins lowercase. */
  isLowercase: boolean;
  /** Immediately preceded by the/a/an/this/that/such/its/their/… (for smart mode). */
  precededByDeterminer: boolean;
  /** First token of a sentence (capitalisation false-positive guard). */
  isSentenceInitial: boolean;
  anchor: TextAnchor;
  inFootnote: boolean;
  /** TRUE for the occurrence inside its own definition — EXCLUDED from "used" counts. */
  isDefinitionOccurrence: boolean;
}

export type UsageIndex = Map<string, UsageRecord[]>;
