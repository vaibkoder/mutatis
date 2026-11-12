/**
 * Template feature — core types.
 *
 * Schemas for template variables (C.4.1), shared data sets (C.4.2),
 * and fixed-slot repeating groups (C.4.3).
 *
 * Persistence shape:  document.settings  key  "mutatis.templateSchema"
 *                     document.settings  key  "mutatis.hiddenSections"   (conditional cache)
 *                     localStorage       key  "mutatis.datasets.*"         (shared data sets)
 */

// ---------------------------------------------------------------------------
// C.4.1 — Template variable schema
// ---------------------------------------------------------------------------

/** All widget types rendered in the questionnaire. */
export type VariableType =
  | 'text'
  | 'longtext'
  | 'number'
  | 'date'           // rendered DD Month YYYY, never month-first
  | 'select'
  | 'entityType'     // jurisdiction-aware dropdown (SG/UK entity suffixes)
  | 'regNumber'      // validated UK_CH or SG_UEN from config
  | 'address'        // multi-line
  | 'boolean';       // checkbox — also drives conditionals

/** Optional validation constraints on a variable. */
export interface VariableValidation {
  /** Maximum character length (text/longtext). */
  maxLen?: number;
  /** Regex pattern as a string (serialisable). */
  pattern?: string;
  /** Minimum numeric value (number type). */
  min?: number;
  /** Maximum numeric value (number type). */
  max?: number;
}

/** A single template variable — maps 1:1 to a tagged PlainText content control. */
export interface TemplateVariable {
  /** Machine name, UPPER_SNAKE.  Must be unique within the document. */
  tag: string;
  /** Human label shown in the questionnaire (== content-control .title). */
  label: string;
  /** Widget type. */
  type: VariableType;
  /** Whether an answer is required before submission. */
  required: boolean;
  /** Optional validation constraints. */
  validation?: VariableValidation;
  /** Default display value (pre-filled in the form). */
  default?: string;
  /** Only for type "date" — always "DD MMMM YYYY" per spec. */
  format?: 'DD MMMM YYYY';
  /** Only for type "regNumber" — which scheme's rules to apply. */
  regScheme?: 'UK_CH' | 'SG_UEN';
  /** Only for type "select" or "entityType" — allowed values. */
  options?: string[];
}

// ---------------------------------------------------------------------------
// C.4.1 — Conditional schema
// ---------------------------------------------------------------------------

/** Condition expression for show/hide — compared to the current answers map. */
export type ShowWhenCondition =
  | { var: string; equals: string | boolean | number }
  | { var: string; notEquals: string | boolean | number }
  | { and: ShowWhenCondition[] }
  | { or: ShowWhenCondition[] };

/** What to do when the condition becomes false. */
export type OnHidePolicy =
  | 'non-destructive'      // cache OOXML, clear content, restore on re-show
  | 'destructive-confirm'; // confirm modal → control.delete(false), irreversible

/** A conditional section: one tagged CC whose visibility depends on answers. */
export interface ConditionalSection {
  /** Tag matching the section content control's .tag property. */
  tag: string;
  /** Human label shown in the toggles panel. */
  label: string;
  /** Condition that must be true to show the section. */
  showWhen: ShowWhenCondition;
  /** Default non-destructive; destructive-confirm requires explicit user sign-off. */
  onHide: OnHidePolicy;
}

// ---------------------------------------------------------------------------
// C.4.3 — Fixed-slot repeating group schema
// ---------------------------------------------------------------------------

/** Minimum slot count. */
export const SLOT_COUNT_DEFAULT = 4;
/** Maximum slot count — hard cap per spec §C.4.3. */
export const SLOT_COUNT_MAX = 6;

/** One tag pattern in a repeating slot definition. */
export interface SlotTagPattern {
  /** Field name in a data-set row (e.g. "name"). */
  field: string;
  /** Tag pattern with `{i}` placeholder, e.g. "DIRECTOR_{i}_NAME". */
  tagPattern: string;
}

/** A repeating group definition (fixed slots, no OOXML cloning). */
export interface RepeatingGroup {
  /** Group identifier, UPPER_SNAKE. */
  group: string;
  /**
   * Number of pre-placed slots in the template document.
   * Default SLOT_COUNT_DEFAULT (4); max SLOT_COUNT_MAX (6).
   * Enforced by `guardSlotCount`.
   */
  slotCount: number;
  /** One entry per field type within a slot. */
  slotTags: SlotTagPattern[];
  /** Hide unused slots non-destructively after fill. */
  hideUnusedSlots: boolean;
}

/** Guard: clamp or throw if slotCount violates the spec constraint. */
export function guardSlotCount(n: number): number {
  if (!Number.isInteger(n) || n < 1) {
    throw new RangeError(`slotCount must be a positive integer; got ${n}`);
  }
  if (n > SLOT_COUNT_MAX) {
    throw new RangeError(
      `slotCount ${n} exceeds the fixed-slot cap of ${SLOT_COUNT_MAX}. ` +
      'Use a template with more slots or reduce the row count.',
    );
  }
  return n;
}

// ---------------------------------------------------------------------------
// C.4.1 — Top-level template schema (stored in docSettings.templateSchema)
// ---------------------------------------------------------------------------

/** Active jurisdiction, drives validators, entity-type options, execution blocks. */
export type TemplateJurisdiction = 'EW' | 'SG';

/** The complete template schema stored per-document in document.settings. */
export interface TemplateSchema {
  /** Stable id for this template, e.g. "tmpl_msa_ew_v2". */
  templateId: string;
  /** Jurisdiction profile — drives all jurisdiction-aware defaults. */
  jurisdictionProfile: TemplateJurisdiction;
  /** Scalar variable definitions (each maps to a PlainText content control). */
  variables: TemplateVariable[];
  /** Conditional section definitions. */
  conditionals: ConditionalSection[];
  /** Fixed-slot repeating group definitions. */
  repeating: RepeatingGroup[];
}

/** docSettings key used to persist the template schema. */
export const TEMPLATE_SCHEMA_KEY = 'mutatis.templateSchema';

/** docSettings key for the hidden-sections OOXML cache. */
export const HIDDEN_SECTIONS_KEY = 'mutatis.hiddenSections';

/** Map of conditional tag → cached OOXML for non-destructive restore. */
export type HiddenSectionsCache = Record<string, string>;

// ---------------------------------------------------------------------------
// C.4.2 — Shared Data Set schema
// ---------------------------------------------------------------------------

/** Data set kind. */
export type DataSetKind = 'scalar' | 'repeating';

/** Scope: personal (localStorage) or shared (backend, if configured). */
export type DataSetScope = 'personal' | 'shared';

/** A row in a repeating data set. */
export type DataSetRow = Record<string, string>;

/**
 * A shared data set.
 * Stored under key "mutatis.datasets.<dataSetId>" in the local or shared store.
 */
export interface SharedDataSet {
  /** Stable id, e.g. "ds_acme_directors". */
  dataSetId: string;
  /** Human label. */
  label: string;
  /** personal = localStorage; shared = configured backend. */
  scope: DataSetScope;
  /** Target jurisdiction. */
  jurisdiction: TemplateJurisdiction;
  /** Scalar or repeating. */
  kind: DataSetKind;
  /** For repeating sets: the group identifier this applies to. */
  groupRef?: string;
  /** For repeating sets: the row data. */
  rows?: DataSetRow[];
  /** For scalar sets: tag → value map. */
  values?: Record<string, string>;
}

/** Storage key prefix for data sets. */
export const DATASET_KEY_PREFIX = 'mutatis.datasets.';

// ---------------------------------------------------------------------------
// C.5 — Answers map (in-memory only, not persisted)
// ---------------------------------------------------------------------------

/** In-memory answers map during a fill session: tag → raw answer value. */
export type AnswersMap = Record<string, string | boolean | number>;

// ---------------------------------------------------------------------------
// Fuzzy-match warning (C.5.4 / C-T9)
// ---------------------------------------------------------------------------

/** A candidate fuzzy match between a data-set key and a schema tag. */
export interface FuzzyMatchWarning {
  /** The key name as it appears in the data set. */
  dataSetKey: string;
  /** The matched schema tag (UPPER_SNAKE). */
  schemaTag: string;
  /** The normalised similarity score 0–1 (higher = better). */
  score: number;
}

/** Result of a fuzzy push operation before the user confirms. */
export interface FuzzyPushPreview {
  /** Exact matches (no warning needed). */
  exact: Array<{ tag: string; value: string }>;
  /** Fuzzy matches that need user confirmation before applying. */
  fuzzy: FuzzyMatchWarning[];
  /** Keys that could not be matched to any schema tag. */
  unmatched: string[];
}
