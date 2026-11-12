export type JurisdictionId = 'SG' | 'UK';

export interface TerminologyConfig {
  /** Default top-level clause noun. Always "Clause" for SG/UK. */
  topLevelUnit: 'Clause';
  /** Accepted aliases for the top-level unit (recognised, never emitted as default). */
  bodyAliases: readonly string[];
  /** Default annexure word. Always "Schedule" for SG/UK. */
  scheduleWord: 'Schedule';
  /** Accepted aliases for the schedule word. */
  annexureAliases: readonly string[];
  /** Sub-unit inside a Schedule. */
  scheduleSubUnit: 'paragraph';
  /** Other division names (roman-numbered, etc.). */
  otherDivisions: readonly string[];
  /** Recital section heading synonyms. */
  recitalNouns: readonly string[];
  /** How recitals are lettered. */
  recitalStyle: 'letteredUpper';
  /** Whether "Section" is treated as a top-level division. Default false for SG/UK. */
  sectionIsTopLevelDivision: boolean;
  /** Self-reference phrases (never flag as undefined). */
  selfReferenceTerms: readonly string[];
}

/**
 * Exclusion allowlist for sub-check 1 (used-but-not-defined). A capitalised
 * term matching any of these is never flagged as undefined — it does not
 * prevent pickup if the term IS actually defined (CONVENTIONS.md §Defined-Term).
 */
export interface DefinedTermExclusions {
  /** Party-role words: Buyer, Seller, Company, Parties, Director(s), etc. */
  partyRoles: readonly string[];
  /** Conventionally-capitalised words: Board, Business Day, Group, Effective Date, etc. */
  conventional: readonly string[];
  /** Common acronyms that are not bespoke defined terms: VAT, GST, PDPA, IRAS, ACRA, CPF. */
  acronyms: readonly string[];
  /** Bare statutory references treated as proper nouns: "the Act", "the Regulations". */
  statutoryWords: readonly string[];
  /** Geographic names, nationalities, months, days, languages, company forms, and
   *  legal boilerplate words that are never bespoke defined terms. */
  properNouns: readonly string[];
  /** Pattern matching a statutory citation family: `<X> Act/Regulations <Year>`. */
  statutoryCitation: RegExp;
}

export interface DefinedTermConfig {
  /** Default quote style for definitions. */
  quoteStyle: 'double';
  /** Definition verbs recognised (expanded set — ADR-0013). Sorted longest-first at build time. */
  verbs: readonly string[];
  /** Whether to nudge "shall mean" → "means". Off by default. */
  styleNudges: { shallMean: boolean };
  /** Capitalisation mode. */
  capMode: 'smart' | 'strict';
  /** Determiners that precede a defined term (stripped before term extraction). */
  determiners: readonly string[];
  /** Exclusion allowlist for used-but-not-defined (CONVENTIONS.md §Defined-Term + ADR-0013). */
  exclusions: DefinedTermExclusions;
  /** Flag clause-scoped local redefinitions as defined-multiple-times. Off by default (spec §8). */
  flagLocalRedefinitions: boolean;
  /** Recognise unquoted verb-anchored definitions (Shape C — ADR-0013). Default on. */
  recogniseUnquoted: boolean;
  /** Recognise colon-as-verb definitions (Shape A′ — ADR-0013). Default on. */
  recogniseColon: boolean;
}

export interface CurrencyEntry {
  /** ISO 4217 currency code, or 'AMBIGUOUS' / 'DOLLAR_FAMILY'. */
  code: string;
}

export interface CurrencyConfig {
  /** Map from symbol/word → entry. */
  table: Readonly<Record<string, CurrencyEntry>>;
  /** Policy when a bare '$' is found. */
  ambiguousDollarPolicy: 'flag-soft';
  /** Tolerance for figure/word discrepancies (fractional). */
  centTolerance: number;
}

export interface DateConfig {
  /** Day-first (DD Month YYYY) ordering. */
  order: 'little-endian';
  defaultDisplayFormat: 'DD Month YYYY';
  /** Weekend days as JS Date.getDay() values (0=Sun, 6=Sat). */
  weekendDays: readonly [0, 6];
  /** Keywords that indicate an execution or completion date role (case-insensitive scan). */
  executionAnchorKeywords: readonly string[];
  /** Keywords that indicate an effective / "as of" date role — weekend rule not applied as Critical. */
  effectiveDateKeywords: readonly string[];
  /** Whether to flag weekend dates with any roleHint, not only execution/completion. Default false. */
  flagAllWeekendDates: boolean;
  /** Character window around the date to scan for role-hint keywords. Default 120. */
  roleHintWindow: number;
}

export interface PlaceholderConfig {
  /** Unicode bullet characters that form [●]-style placeholders. */
  bulletChars: readonly string[];
  /** Keyword placeholders: [TBC], [DATE], etc. */
  keywords: readonly string[];
  /** Strings never flagged as placeholders. */
  whitelist: readonly string[];
  /** Case-citation reporter abbreviations (suppress placeholder detection). */
  citationReporters: readonly string[];
  /** Anchor text near signature blocks (suppress placeholder detection). */
  signatureAnchors: readonly string[];
  /** Whether to flag bare optional-drafting brackets like [and the Guarantor]. Off by default. */
  flagBareBrackets: boolean;
}

/** Word.ListNumbering enum values used by setLevelNumbering (spec §3.1 API table). */
export type ListNumberingStyle =
  | 'arabic' | 'upperRoman' | 'lowerRoman' | 'upperLetter' | 'lowerLetter' | 'none';

/** Display hint for the level editor — cosmetic only. */
export type LevelLabelStyle =
  | 'decimal-compound' | 'paren-lower-alpha' | 'paren-lower-roman' | 'paren-upper-alpha' | 'plain';

/**
 * A single level within a `NumberingPreset` / `NumberingScheme`.
 * Enriched to spec §4.1 shape (feature 06) so the apply engine has everything
 * it needs for setLevelNumbering + setLevelIndents without extra lookup.
 */
export interface NumberingLevel {
  /** 0-based Word list level (0..8). */
  levelIndex: number;
  /** Maps to Word.ListNumbering enum. */
  listNumbering: ListNumberingStyle;
  /** Word formatString with %1..%9 tokens, e.g. "%1.%2." compound, "(%4)" lettered limb. */
  formatString: string;
  /** Literal text before the number token (usually empty — brackets go in formatString). */
  prefix: string;
  /** Literal text after the number token (usually empty). */
  suffix: string;
  /** Hanging-indent geometry in points. */
  indent: { textIndentPt: number; numberIndentPt: number };
  /** Display hint for the level editor. */
  labelStyle: LevelLabelStyle;
  /**
   * restartAfterLevel: N means "reset this level's counter after any level ≤ N".
   * -1 = never restart (continuous — used for top-level clauses across the document).
   * null = inherit Word default (restart under immediate parent).
   */
  restartAfterLevel: number | null;
  /** Optional alignment hint for setLevelAlignment. Default 'left'. */
  alignment?: 'left' | 'right' | 'center';
}

/**
 * A complete, named legal numbering scheme.
 * Built-in presets live in the JurisdictionProfile and are read-only in the UI
 * ("Duplicate to edit"). User schemes are stored via the Storage abstraction.
 * The `suppressForSgUk` field is config-layer-specific (hides the US preset).
 */
export interface NumberingPreset {
  /** Stable UUID — preserved across export/import; built-in ids are reserved. */
  id: string;
  /** Human label shown in the library, e.g. "UK/SG Commercial". */
  name: string;
  /** "SG-UK" | "US" | "neutral" — drives default selection + label strings. */
  jurisdiction: 'SG-UK' | 'US' | 'neutral';
  /** Optional longer description for library tooltip. */
  description?: string;
  /** True for shipped presets (read-only in UI; user must Duplicate to edit). */
  builtIn?: boolean;
  /** Unit noun for cross-ref labelling: "clause" (operative) or "paragraph" (schedules). */
  unitLabel?: string;
  /** Schema version for forward-compat on import. */
  schemaVersion: number;
  /** Whether this preset is hidden when the active profile is SG/UK. */
  suppressForSgUk: boolean;
  /** Levels in index order. Must be contiguous from 0. */
  levels: readonly NumberingLevel[];
}

/**
 * Feature-06-specific numbering config on the JurisdictionProfile.
 * Optional; code-side defaults are returned by `getNumberingConfig()`.
 */
export interface NumberingConfig {
  /** ID of the preset to pre-select when the Numbering tab opens. */
  defaultSchemeId: string;
  /** Highlight colour applied to cross-reference candidates after apply/promote/demote. */
  reviewHighlightColour: string;
  /**
   * Default whitespace rule for the Numbering upgrade's assisted removal
   * (feature 06 §11.1, ADR-0023 constraint 3). The user may override this once,
   * globally, in the upgrade UI (persisted via Storage). `auto` == strip-to-flush.
   */
  whitespaceOverride: NumberingWhitespaceOverride;
}

/**
 * Global whitespace override for assisted static-number removal (ADR-0023 #3).
 * Canonical home for the union — feature code imports it from here (shared is the
 * carrier; one definition).
 */
export type NumberingWhitespaceOverride =
  | 'auto'
  | 'collapse-to-one-space'
  | 'strip-to-flush'
  | 'keep-one-tab';

export interface EntityFieldConfig {
  /** Company registration number validators by jurisdiction. */
  registrationNumbers: Readonly<Record<string, RegExp>>;
  /** Allowed entity suffixes (no US forms). */
  suffixes: readonly string[];
}

export interface XrefUnitsConfig {
  /** Default top-level clause noun. Always "Clause" for SG/UK. */
  primaryBody: 'Clause';
  /** Accepted aliases recognised in prose (never emitted as default replacement). */
  bodyAliases: readonly string[];
  /** Default annexure noun. Always "Schedule" for SG/UK. */
  annexurePrimary: 'Schedule';
  /** Accepted aliases for the schedule word (e.g. "Exhibit" recognised, never emitted). */
  annexureAliases: readonly string[];
  /** Sub-unit noun inside a Schedule/Annex/Appendix. Always "paragraph" for SG/UK. */
  scheduleSubUnit: 'paragraph';
  /** Other structural division nouns (Annex, Annexure, Appendix, Part). */
  otherDivisions: readonly string[];
  /** Recital section heading synonyms (Recital, Background). */
  recitalNouns: readonly string[];
  /** How recitals are lettered: always uppercase for SG/UK. */
  recitalStyle: 'letteredUpper';
}

export interface XrefNumberingConfig {
  bodyStyle: 'decimal';
  subClauseStyle: readonly string[];
}

export interface XrefRangesConfig {
  /** Range connectors: "to", em-dash, en-dash, hyphen. */
  connectors: readonly string[];
  /** List connectors: comma, "and", "&". */
  listConnectors: readonly string[];
  /** Inclusive marker consumed and ignored. */
  inclusiveMarker: string;
}

export interface XrefSectionConfig {
  /** When true, "Section N" is a top-level division (LMA/finance house style). Default false. */
  isTopLevelDivision: boolean;
}

export interface XrefStylesConfig {
  headingStyles: readonly string[];
  scheduleHeading: readonly string[];
  /** TOC paragraph styles — excluded from both index-building and reference-scanning. */
  tocStyles: readonly string[];
}

export interface CrossRefConfig {
  units: XrefUnitsConfig;
  numbering: XrefNumberingConfig;
  ranges: XrefRangesConfig;
  /** Self-reference phrases never matched as cross-references (spec §6.7). */
  selfReferenceTerms: readonly string[];
  section: XrefSectionConfig;
  styles: XrefStylesConfig;
}

/**
 * Conceptual severity label for static-numbering findings (spec §5.2.5 / §7).
 * Mapped to the shared Finding `Severity` union at emit time: low → 'info',
 * medium → 'warning' (the Finding model has no 'low'/'medium' — CONVENTIONS.md
 * §Findings; see TerminologyConfig note "the store normalises").
 */
export type StaticSeverityLabel = 'low' | 'medium';

export interface StaticNumberingConfig {
  /** Severity for static numbering on a heading-styled paragraph. */
  headingSeverity: StaticSeverityLabel;
  /** Severity for static numbering on a body paragraph. */
  bodySeverity: StaticSeverityLabel;
  /** Cap on getOoxml() calls per run (layer c); default 200. */
  maxOoxmlProbes: number;
}

export interface LookupDefinitionConfig {
  /** Accept smart (“/”) or straight (") quotes around terms. */
  quoteStyle: 'smart' | 'straight';
  /** Verbs that introduce a definition (used for preview display). */
  introVerbs: readonly string[];
}

/**
 * Interactive Reference Lookup config (feature 04). Optional on the profile;
 * all fields have code-side defaults via `getLookupConfig(profile)`.
 */
export interface LookupConfig {
  /** Trailing debounce on DocumentSelectionChanged before running a lookup. Default 400. */
  debounceMs: number;
  /** Debounce on dictionary rebuild after a DocModel invalidation event. Default 1000. */
  dictRebuildDebounceMs: number;
  /** Max chars of the sentence window read for bare-cursor classification. Default 240. */
  sentenceWindowChars: number;
  /** Max chars of preview text shown in a collapsed/single panel. Default 600. */
  previewMaxChars: number;
  /** Max chars stored/shown for an expanded definition card ("fuller view"). Default 4000. */
  expandedPreviewMaxChars: number;
  definition: LookupDefinitionConfig;
}

/** Unified jurisdiction profile — all checkers read from here; nothing is hardcoded. */
/**
 * Party-Role Consistency checker (feature 09 — ADR-0016). Optional on the
 * profile: a profile that omits it is handled by the feature's code-side
 * defaults (`resolvePartyRoleConfig`), so older/partial profiles still work.
 * The role vocabulary and entity suffixes are NOT restated here — the feature
 * consumes `definedTerms.exclusions.partyRoles` and `entityFields.suffixes`.
 */
export interface PartyRoleConsistencyConfig {
  /** PR1 recurrence gate: a role must be used at least this many times in the
   *  body before an undefined-role finding fires (conservative — cf. ADR-0015). */
  minUsesForUndefined: number;
  /** Severity for PR1 (used-but-never-introduced) findings. */
  pr1Severity: 'error' | 'warning' | 'info';
  /** Severity for PR2 (same-party-two-labels) findings. */
  pr2Severity: 'error' | 'warning' | 'info';
  /** Collective role words too generic to flag (e.g. "Party", "Parties"). */
  excludeCollectiveRoles: readonly string[];
}

/** Flatten policy applied to pre-existing tracked changes before diffing. */
export type FlattenPolicy = 'acceptAll' | 'rejectAll';

/** Footnote comparison mode. */
export type FootnoteCompareMode = 'compare' | 'ignore';

/**
 * Feature-08 Compare config on the JurisdictionProfile.
 * Optional; code-side defaults are returned by `getCompareConfig()`.
 */
export interface CompareConfig {
  /** Author string for w:author on every emitted revision. */
  author: string;
  /** Reference tokens for cross-reference renumber detection & cascade collapse. */
  referenceTokens: readonly string[];
  /**
   * The jurisdiction's noun for a top-level body division — "Clause" for SG and
   * England & Wales (ADR-0002). Used where Compare has to NAME the unit rather
   * than recognise it: the §6.1 numbering-cascade heading, "Clause numbering
   * (level 2) shifted (+1)".
   *
   * **Not a second home for that noun.** `getCompareConfig` fills this from the
   * profile's `xref.units.primaryBody`, which `spec/CONVENTIONS.md` makes the
   * single source of truth; setting it here overrides that deliberately, for a
   * profile that wants Compare to differ from the cross-reference layer.
   *
   * Separate from `referenceTokens` because that is an unordered DETECTION set
   * — which nouns to recognise in prose — and naming the primary unit by taking
   * its first element is a coincidence, not a rule.
   */
  primaryUnit: string;
  /** Minimum group size to trigger renumbering-cascade collapse (default 3). */
  renumberThreshold: number;
  /**
   * Change-complexity threshold (ADR-0026). A matched `replace` paragraph pair
   * whose Sørensen–Dice token similarity is BELOW this value is demoted to a
   * clean whole-paragraph delete + insert instead of an inline word diff — the
   * the conventional "too dissimilar → show as insert/delete" readability rule.
   * Range [0,1]; default 0.5 (demote only when a majority of content differs).
   * `0` disables the pass. Never suppresses: a demoted pair shows MORE markup
   * (whole old para deleted + whole new para inserted), never less (§5.3).
   */
  replaceSimilarityThreshold: number;
  /**
   * Contiguity escape from the demotion above (ADR-0041). A `replace` pair that
   * FAILS the similarity test is still kept for inline word-diffing when this
   * share or more of the SHORTER paragraph's content tokens survive as one
   * unbroken run in the other.
   *
   * Sørensen–Dice is a set measure and therefore punishes length asymmetry —
   * which is exactly backwards for the commonest heavy edit in drafting, a
   * large contiguous DELETION. Cutting two sentences out of a three-sentence
   * clause scores ~0.36 and demotes, yet its inline rendering is a single
   * strike-through with everything else untouched: the most readable form
   * there is, and what commercial comparison tools conventionally produce. Contiguity is what separates a
   * clean cut from the "equal islands in a sea of ins/del" salad the demotion
   * exists to prevent.
   *
   * Range [0,1]; default 0.8. `1` requires a perfect contiguous survivor;
   * values > 1 disable the escape. Never suppresses either way — this only
   * chooses between two renderings of the same text (§5.3).
   */
  replaceContiguityThreshold: number;
  /**
   * Bidirectional en-US→en-GB spelling variant map (lowercase keys, lowercase values).
   * e.g. `{ 'organize': 'organise', 'color': 'colour', … }`
   * Detection is case-insensitive; changes are always surfaced as `substantiveText`,
   * never `formatting`. (spec §7.1, CONVENTIONS.md §en-GB Spelling)
   */
  spellingVariants: Readonly<Record<string, string>>;
  /** Exclude headers from comparison (§9.7). Default false — compared (ADR-0047). */
  excludeHeaders: boolean;
  /** Exclude footers from comparison (§9.7). Default false — compared (ADR-0047). */
  excludeFooters: boolean;
  /** Whether to compare or ignore footnote content (§9.8). */
  footnotes: FootnoteCompareMode;
  /** How to flatten pre-existing tracked changes on the Original side before diffing. */
  flattenOriginal: FlattenPolicy;
  /** How to flatten pre-existing tracked changes on the Revised side before diffing. */
  flattenRevised: FlattenPolicy;
  /**
   * Compare at scale (feature 17 — ADR-0037). Optional and partial.
   *
   * The KEYS live here so batch settings resolve through the existing
   * `CompareConfig` the way every other Compare setting does. The DEFAULTS
   * deliberately do not: they are named constants in
   * `features/compare/batch/constants.ts`, resolved by
   * `features/compare/batch/config.ts`, because `shared/` must not depend on a
   * feature. An unset key here means "use the constant", never "use zero".
   */
  batch?: CompareBatchConfig;
}

/** Pairing weights and thresholds (spec 17 §6). All optional. */
export interface CompareBatchPairingConfig {
  weightContent?: number;
  weightFilename?: number;
  tieEpsilon?: number;
  autoPairScore?: number;
  autoPairMargin?: number;
  noMatchScore?: number;
  /** Filename bar for the content-free second assignment phase (spec 17 §6.4). */
  noContentMatchScore?: number;
  minSketchTokens?: number;
}

/** `compare.batch.*` (spec 17 §7). All optional; defaults are feature constants. */
export interface CompareBatchConfig {
  maxVariants?: number;
  maxBundleDocuments?: number;
  /** `0` = auto-size from `hardwareConcurrency`. */
  poolSize?: number;
  pairing?: CompareBatchPairingConfig;
}

export interface JurisdictionProfile {
  id: JurisdictionId;
  terminology: TerminologyConfig;
  currency: CurrencyConfig;
  dateLocale: DateConfig;
  definedTerms: DefinedTermConfig;
  placeholderVocab: PlaceholderConfig;
  numberingPresets: readonly NumberingPreset[];
  entityFields: EntityFieldConfig;
  spellingLocale: 'en-GB';
  xref: CrossRefConfig;
  staticNumbering: StaticNumberingConfig;
  /** Party-Role Consistency config (feature 09 — ADR-0016). Optional; defaulted in code. */
  partyRoleConsistency?: PartyRoleConsistencyConfig;
  /** Interactive Reference Lookup config (feature 04). Optional; defaulted in code. */
  lookup?: LookupConfig;
  /** Numbering tab config (feature 06). Optional; defaulted in code via getNumberingConfig(). */
  numbering?: NumberingConfig;
  /** Compare (feature 08). Optional; defaulted in code via getCompareConfig(). */
  compare?: CompareConfig;
}
