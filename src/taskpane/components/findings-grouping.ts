/**
 * Host-neutral findings presentation helpers — category labels, feature ordering,
 * and the split that groups a category→findings map into the four display
 * features (cross-reference, static-numbering, defined-terms, other).
 *
 * Single source of truth shared by the add-in's `FindingsList` (reads from the
 * global findings store) and the standalone web app's `WebFindings` (works off
 * the pure `analyse*` returns). One home for the labels/order,
 * no duplicate tables. Contains NO `Office.*`/`Word.*` token, so it imports
 * cleanly into the host-free web bundle.
 */
import type { Finding } from '@shared/findings/index.js';

export const CATEGORY_LABELS: Partial<Record<string, string>> = {
  'defined-terms/used-not-defined': 'Used but not defined',
  'defined-terms/defined-not-used': 'Defined but not used',
  'defined-terms/defined-multiple': 'Defined more than once',
  'defined-terms/capitalisation': 'Capitalisation inconsistency',
  'cross-reference/not-found': 'Broken (not found)',
  'cross-reference/ambiguous': 'Ambiguous',
  'cross-reference/external-maybe': 'May be in a separate file',
  'cross-reference/static': 'Manually typed reference (may be out of date)',
  'static-numbering/manually-typed': 'Manually-typed clause numbers',
  'factual-consistency/weekend-date': 'Weekend execution dates',
  'factual-consistency/amount-mismatch': 'Amount/word mismatches',
  'factual-consistency/currency-mismatch': 'Currency disagreements',
  'factual-consistency/ambiguous-currency': 'Ambiguous currency (check)',
  'factual-consistency/century-ambiguous': 'Century-ambiguous dates (check)',
  'factual-consistency/placeholder': 'Unfilled placeholders',
  'party-role/undefined-role': 'Party-role used but never introduced',
  'party-role/inconsistent-label': 'Same party, two labels',
};

/** Categories belonging to Feature 01 (defined terms). Ordered per spec §6. */
export const DEFINED_TERMS_CATEGORY_ORDER = [
  'defined-terms/used-not-defined',
  'defined-terms/defined-not-used',
  'defined-terms/defined-multiple',
  'defined-terms/capitalisation',
] as const;

/** Categories belonging to Feature A (cross-reference). Ordered per spec §8.1. */
export const XREF_CATEGORY_ORDER = [
  'cross-reference/not-found',
  'cross-reference/ambiguous',
  'cross-reference/external-maybe',
  'cross-reference/static',
] as const;

export function categoryLabel(category: string): string {
  return CATEGORY_LABELS[category] ?? category;
}

/** A category sub-group ready to render: its category id and its findings. */
export interface CategoryGroup {
  cat: string;
  findings: Finding[];
}

/** Group a flat findings list into a category→findings map, preserving order. */
export function groupByCategory(findings: ReadonlyArray<Finding>): Map<string, Finding[]> {
  const out = new Map<string, Finding[]>();
  for (const f of findings) {
    const bucket = out.get(f.category);
    if (bucket) bucket.push(f);
    else out.set(f.category, [f]);
  }
  return out;
}

/**
 * Split a category→findings map into the four display features. Cross-reference
 * (Feature A) and defined-terms (Feature 01) sub-groups are returned in their
 * spec order with empty categories dropped; static-numbering (Feature B) is a flat
 * list; everything else (factual-consistency, party-role…) falls into `other`.
 */
export function splitFindingsByFeature(grouped: Map<string, Finding[]>): {
  xrefOrdered: CategoryGroup[];
  staticNum: Finding[];
  definedTermsOrdered: CategoryGroup[];
  other: CategoryGroup[];
} {
  const xref = new Map<string, Finding[]>();
  const staticNum: Finding[] = [];
  const defined = new Map<string, Finding[]>();
  const other: CategoryGroup[] = [];

  for (const [cat, findings] of grouped.entries()) {
    if (cat.startsWith('cross-reference/')) {
      xref.set(cat, findings);
    } else if (cat.startsWith('static-numbering/')) {
      staticNum.push(...findings);
    } else if (cat.startsWith('defined-terms/')) {
      defined.set(cat, findings);
    } else {
      other.push({ cat, findings });
    }
  }

  const ordered = (order: ReadonlyArray<string>, map: Map<string, Finding[]>): CategoryGroup[] =>
    order
      .map(cat => ({ cat, findings: map.get(cat) ?? [] }))
      .filter(({ findings }) => findings.length > 0);

  return {
    xrefOrdered: ordered(XREF_CATEGORY_ORDER, xref),
    staticNum,
    definedTermsOrdered: ordered(DEFINED_TERMS_CATEGORY_ORDER, defined),
    other,
  };
}
