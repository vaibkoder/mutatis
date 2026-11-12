/** Finding severity levels. Features may label these as critical/medium/low in their own UI; the store normalises. */
export type Severity = 'error' | 'warning' | 'info';

/**
 * Host-neutral location token (ADR-0022). Carries only data — a paragraph id, an
 * optional character span into that paragraph's normalised text, and a label.
 * Navigation is performed by the active host via `HostPort.reveal(location)`,
 * never by holding a `Word.Range`. This is structurally a `DocLocation`
 * (`@shared/host`); kept as its own name so checkers depend on `@shared/findings`
 * without importing the host port. Keep the two shapes in sync.
 */
export interface FindingLocation {
  /** Paragraph ID from the DocModel. May be absent for document-level findings. */
  paraId?: string;
  /** Inclusive start offset of the offending span into the paragraph's text. */
  charStart?: number;
  /** Exclusive end offset of the offending span into the paragraph's text. */
  charEnd?: number;
  /** Human-readable location label (e.g. "Clause 4.2(a)"). */
  label: string;
}

export interface SuggestedAction {
  /** Button label (e.g. "Accept fix"). Keep short. */
  label: string;
  /** Apply the fix. Called inside a Word.run context. */
  apply(ctx: Word.RequestContext): Promise<void>;
}

export interface Finding {
  /** Stable unique identifier for deduplication. */
  id: string;
  /** Grouping key — e.g. 'cross-reference', 'defined-terms/used-not-defined'. */
  category: string;
  /** Hard discriminator keeping distinct features ungrouped in the UI. */
  featureId?: string;
  severity: Severity;
  title: string;
  detail?: string;
  location: FindingLocation;
  suggestedAction?: SuggestedAction;
  /**
   * Optional excerpt from the source/reference being compared against.
   * When present, the location label becomes a toggle button that expands
   * this text inline below the finding row.
   */
  sourceSnippet?: { label: string; text: string };
  /**
   * All locations of this issue in document order, for multi-occurrence
   * navigation. Present (length > 1) when the same issue recurs — e.g. an
   * undefined term used in several places. `location` equals `occurrences[0]`
   * by convention. The UI shows an occurrence counter + Next control when set.
   */
  occurrences?: FindingLocation[];
}

/** Tab identifiers that map findings to feature tabs. */
export type TabId = 'proof' | 'review' | 'repair' | 'numbering' | 'content' | 'compare' | 'settings';

/** Map from feature category prefix → tab. */
export const CATEGORY_TAB_MAP: Readonly<Record<string, TabId>> = {
  'defined-terms': 'proof',
  'cross-reference': 'proof',
  'static-numbering': 'proof',
  'factual-consistency': 'proof',
  'party-role': 'proof',
  'interactive-lookup': 'review',
  'repair': 'repair',
  'numbering': 'numbering',
  'content': 'content',
  'compare': 'compare',
};

export function findingToTab(finding: Finding): TabId {
  for (const [prefix, tab] of Object.entries(CATEGORY_TAB_MAP)) {
    if (finding.category.startsWith(prefix)) return tab;
  }
  return 'proof';
}
