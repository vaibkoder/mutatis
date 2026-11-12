/**
 * Feature 07-A — Table of Contents: type definitions.
 *
 * Types are config-schema-forward: the runtime representation mirrors the JSON
 * stored in docSettings and firm config, so serialise/deserialise is trivial.
 * All label strings must be sourced from the Localisation/Config layer or the
 * spec default config; no US-isms ("Exhibit", "Section" as primary, month-first
 * dates) may appear here.
 *
 * Spec references: spec/features/07-toc-precedent-templates.md §A.4 / §A.4.1.
 */

// ---------------------------------------------------------------------------
// TOC profile — firm- or document-level config
// ---------------------------------------------------------------------------

/** Tab-leader style for TOC entries (not a field switch — applied via TOC styles). */
export type TabLeader = 'dots' | 'dashes' | 'none';

/** Per-style → TOC level mapping. When non-empty, supersedes \o for those styles. */
export interface StyleLevelEntry {
  /** Named Word paragraph style, e.g. "Heading 1" or "Clause Heading". */
  style: string;
  /** TOC level (1–9). */
  level: number;
}

/** Switch settings portion of a TOC profile. */
export interface TocSwitches {
  /** Level range string, e.g. "1-3". Used to emit \o "1-3" when styleLevelMap is empty. */
  levels: string;
  /** Emit \h (hyperlink entries). Default: true. */
  useHyperlinks: boolean;
  /** Emit \n "<levels>" to hide page numbers. Default: false. */
  hidePageNumbers: boolean;
  /** Right-align page numbers (via TOC styles, not a field switch). Default: true. */
  rightAlignPageNo: boolean;
  /** Preserve tab and leader from the TOC field (informational, not a switch). */
  preserveTabAndLeader: boolean;
}

/** A named TOC profile, carrying switch settings and an optional style-level map. */
export interface TocProfile {
  /** Stable identifier; the built-in profile uses "house-default". */
  id: string;
  /** Human-readable label, displayed in the profile selector. */
  label: string;
  switches: TocSwitches;
  /**
   * When non-empty, emits \t "Style,Level;…" INSTEAD OF \o for these styles.
   * Firms remap to their bespoke heading style names here.
   */
  styleLevelMap: StyleLevelEntry[];
  /** Tab leader (dots | dashes | none); applied to TOC 1–9 styles, not the field code. */
  tabLeader: TabLeader;
}

// ---------------------------------------------------------------------------
// Index of Schedules config
// ---------------------------------------------------------------------------

/** Where to insert the Index of Schedules relative to the main TOC. */
export type SchedulePlacement = 'afterMainToc' | 'atCursor' | 'beforeFirstSchedule';

/** Config for generating an Index of Schedules. */
export interface ScheduleIndexConfig {
  enabled: boolean;
  /** Label used in the UI and as a heading; must never be "Table of Authorities". */
  label: string;
  /** Named paragraph styles that identify Schedule headings. */
  scheduleStyles: string[];
  /** TOC level to assign these entries in the \t switch. */
  level: number;
  /** True → its own TOC field with \t (Strategy 1); false → HTML snapshot (Strategy 2). */
  separateField: boolean;
  placement: SchedulePlacement;
}

// ---------------------------------------------------------------------------
// TOC registry — persisted to docSettings
// ---------------------------------------------------------------------------

/** A registry entry describing one inserted TOC field. */
export interface TocRegistryEntry {
  /** Stable Word field id (field.id), used to locate the field for refresh/go-to. */
  fieldId: string;
  /** Optional bookmark name for scoped TOCs (e.g. "TOC_SCOPE_part1_1"). */
  bookmark?: string;
  /** Human-readable label shown in the registered-TOC list. */
  scopeLabel: string;
}

/** The full registry array stored under docSettings key "toc.registry". */
export type TocRegistry = TocRegistryEntry[];

// ---------------------------------------------------------------------------
// Build-switches options
// ---------------------------------------------------------------------------

/** Options passed to buildTocSwitches alongside the profile. */
export interface BuildSwitchesOptions {
  /** When provided, appends \b <bookmark> to the switch string (scoped TOC). */
  bookmark?: string;
}

// ---------------------------------------------------------------------------
// Strategy 2 snapshot item
// ---------------------------------------------------------------------------

/** One entry in a Strategy-2 Index of Schedules snapshot. */
export interface ScheduleSnapshotEntry {
  paragraphId: string;
  text: string;
  /** Bookmark name inserted at the paragraph (for HTML link targets). */
  bookmark: string;
}
