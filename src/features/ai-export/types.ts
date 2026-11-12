/**
 * AI-readable export schemas (feature 11 — ADR-0021, spec §4).
 *
 * These are the shapes the two serialisers (JSON + Markdown) consume. They are
 * a *projection* of the checkers' in-memory records — named-field picks only,
 * never a raw dump of the source TypeScript objects (rule F9).
 */

/** Which checker contributed data this run (spec §4.1). */
export type CheckId =
  | 'defined-terms'
  | 'cross-reference'
  | 'factual-consistency'
  | 'compare'
  | 'party-role';

/** Output format. JSON is canonical; Markdown is the paste-into-chat rendering (F7). */
export type ExportFormat = 'json' | 'markdown';

/** Section keys that may appear on the envelope, one per included export. */
export type ExportSectionKey =
  | 'defined_terms'
  | 'section_outline'
  | 'change_set'
  | 'deal_facts'
  | 'cast_list'
  | 'briefing_pack';

// ---------------------------------------------------------------------------
// Top-level envelope (spec §4.1)
// ---------------------------------------------------------------------------

export interface AiExportEnvelope {
  /** Format version sentinel. */
  _mutatis: '1.0';
  /** ISO-8601 generation timestamp. */
  generated: string;
  /** Filename, or "Untitled" if unavailable. */
  document: string;
  /** From the Config layer. */
  jurisdiction: 'SG' | 'EW' | 'unknown';
  /** Which checkers contributed data this run. */
  checks_run: CheckId[];
  /** What was NOT covered (rule F6) — always populated. */
  blind_spots: string[];

  // One key per included export (all optional; present when contributed).
  defined_terms?: DefinedTermsExport;
  section_outline?: SectionOutlineExport;
  change_set?: ChangeSetExport;
  deal_facts?: DealFactsExport;
  cast_list?: CastListExport;
  /** Only present on the composite briefing pack. */
  preamble?: BriefingPreamble;
}

// ---------------------------------------------------------------------------
// #1 Defined-terms glossary (§4.2)
// ---------------------------------------------------------------------------

export interface DefinedTermsExport {
  glossary: TermEntry[];
}

export interface TermEntry {
  /** Canonical surface form, e.g. "Completion Date". */
  term: string;
  /** Full resolved definition text; cross-refs carry the referent inline (F2). */
  definition: string;
  /** Human citation, e.g. "Clause 1.1". */
  defined_at: string;
  source: 'prose' | 'table' | 'footnote';
  /** Omit key when false. */
  defined_multiple_times?: true;
  /** Observed surface variants that differ from `term` (F5). Omit key when empty. */
  variants?: string[];
}

// ---------------------------------------------------------------------------
// #2 Section outline / clause tree (§4.3)
// ---------------------------------------------------------------------------

export interface SectionOutlineExport {
  clauses: ClauseNode[];
}

export interface ClauseNode {
  /** SectionNode.displayLabel, e.g. "Clause 1.2.3" or "Schedule 2". */
  number: string;
  /** Heading text, or "" if unnumbered/unavailable. */
  heading: string;
  /** 1-based nesting depth. */
  level: number;
  /** SectionNode.kind. */
  kind: string;
  /** Nested children; omit key when empty. */
  children?: ClauseNode[];
}

// ---------------------------------------------------------------------------
// #3 Change set with before/after (§4.4)
// ---------------------------------------------------------------------------

export interface ChangeSetExport {
  baseline: string;
  revised: string;
  changes: ChangeEntry[];
}

export interface ChangeEntry {
  id: string;
  type: 'insert' | 'delete' | 'replace' | 'move';
  /** Primary resolved class (DiffOp.classification). */
  classification: string;
  /** Human citation (best-effort part/clause label). */
  clause: string;
  /** Present for delete | replace | move. */
  original?: string;
  /** Present for insert | replace | move. */
  revised?: string;
  /** For type:"move" — target clause citation. */
  move_target?: string;
  /** Omit key when false. */
  low_confidence?: true;
}

// ---------------------------------------------------------------------------
// #4 Deal-facts sheet (§4.5)
// ---------------------------------------------------------------------------

export interface DealFactsExport {
  dates: DateEntry[];
  amounts: AmountEntry[];
  placeholders: PlaceholderEntry[];
}

export interface DateEntry {
  raw: string;
  /** ISO date (yyyy-mm-dd), or "" if unparsed/century-ambiguous. */
  iso: string;
  /** Day name, or "" if unknown. */
  day_of_week: string;
  role: string;
  flagged_weekend: boolean;
  clause: string;
}

export interface AmountEntry {
  raw: string;
  /** Figure value, or null if unparseable. */
  value: number | null;
  /** ISO 4217, or "" if unknown. */
  currency: string;
  words_raw?: string;
  words_value?: number;
  /** Omit key when status === "match". */
  mismatch?: 'value' | 'currency' | 'both';
  clause: string;
}

export interface PlaceholderEntry {
  raw: string;
  category: string;
  clause: string;
}

// ---------------------------------------------------------------------------
// #5 Party/role cast list (§4.6)
// ---------------------------------------------------------------------------

export interface CastListExport {
  parties: PartyEntry[];
}

export interface PartyEntry {
  /** Canonical role surface form, e.g. "Purchaser". */
  role: string;
  /** Observed entity names associated with this role. */
  entities: string[];
  /** Clause citation for the introduction; omit if not formally introduced. */
  introduced_at?: string;
  /** Body occurrences (excluding introduction). */
  body_uses: number;
  /** Present if the party-role checker flagged this entry. */
  finding?: 'used-not-introduced' | 'inconsistent-label';
}

// ---------------------------------------------------------------------------
// #14 Briefing pack (§4.7)
// ---------------------------------------------------------------------------

export interface BriefingPreamble {
  /** One-sentence mechanically-assembled summary (not AI-generated). */
  summary: string;
  /** Unfilled placeholders + undefined-but-used terms. */
  open_items: OpenItem[];
}

export interface OpenItem {
  type: 'placeholder' | 'undefined-term';
  raw: string;
  clause: string;
}
