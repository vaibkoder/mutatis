/**
 * Feature 08 — Compare: build the leading summary-annotation block (C1).
 *
 * Pure OOXML authoring. The block is injected OUTSIDE the revision stream
 * (before the body, via insertOoxml in the driver — §5.5) so that accepting or
 * rejecting tracked changes never disturbs it, and the user can delete the
 * whole block in one action.
 *
 * Spec: spec/features/08-compare.md §5.5, §8.2.
 */
import type { ChangeClass, ChangeSummary } from '../model.js';

/** Human-readable labels for each ChangeClass (en-GB, SG/UK terminology). */
const CLASS_LABELS: Record<ChangeClass, string> = {
  structural: 'Structural',
  obligationLanguage: 'Obligation language',
  numericalValue: 'Numerical value',
  definedTerm: 'Defined-term',
  crossRefRenumber: 'Cross-reference renumbering',
  substantiveText: 'Substantive text',
  whitespacePunct: 'Whitespace / punctuation',
  formatting: 'Formatting-only',
};

/** Priority order for display (top = most significant). */
const CLASS_ORDER: ChangeClass[] = [
  'structural',
  'obligationLanguage',
  'numericalValue',
  'definedTerm',
  'crossRefRenumber',
  'substantiveText',
  'whitespacePunct',
  'formatting',
];

function escapeText(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/** One plain-text paragraph, optionally bold. */
function para(text: string, bold = false): string {
  const rPr = bold ? '<w:rPr><w:b/></w:rPr>' : '';
  return `<w:p><w:r>${rPr}<w:t xml:space="preserve">${escapeText(text)}</w:t></w:r></w:p>`;
}

/**
 * Build the summary-annotation block as an OOXML string suitable for
 * insertOoxml (a self-contained sequence of `<w:p>` paragraphs).
 *
 * Contents (§8.2): total change count, per-category breakdown, dedicated moved
 * and en-GB/US spelling lines, renumbering-cascade annotations, and a count of
 * quarantined comments.
 *
 * Returned as a bare run of paragraphs; the caller wraps it for insertOoxml.
 */
export function buildSummaryAnnotations(summary: ChangeSummary): string {
  const lines: string[] = [];

  lines.push(para('Compare summary', true));
  lines.push(para(`Total changes: ${summary.total}`));

  // Per-category counts, in priority order, omitting zero buckets for brevity
  // but always keeping the block auditable.
  for (const cls of CLASS_ORDER) {
    const count = summary.byCategory[cls] ?? 0;
    if (count > 0) {
      lines.push(para(`  ${CLASS_LABELS[cls]}: ${count}`));
    }
  }

  // Dedicated lines (§8.2).
  lines.push(para(`Moved paragraphs: ${summary.movedParagraphs}`));
  lines.push(para(`en-GB/en-US spelling changes: ${summary.spellingVariantCount}`));

  const quarantined =
    summary.quarantinedCommentsOriginal + summary.quarantinedCommentsRevised;
  if (quarantined > 0) {
    lines.push(
      para(
        `Comments excluded from compare: ${quarantined} ` +
          `(Original ${summary.quarantinedCommentsOriginal}, Revised ${summary.quarantinedCommentsRevised})`,
      ),
    );
  }

  // Renumbering cascades — one collapsed line each (§6.1).
  if (summary.renumberingCascades.length > 0) {
    lines.push(para('Renumbering cascades:', true));
    for (const cascade of summary.renumberingCascades) {
      const sign = cascade.shift > 0 ? `+${cascade.shift}` : `${cascade.shift}`;
      const examples = cascade.exampleRefs.length > 0
        ? ` — example: ${cascade.exampleRefs.join('; ')}`
        : '';
      lines.push(
        para(
          `  ${cascade.tokenKind} references renumbered (${sign}): ` +
            `${cascade.affectedCount} reference${cascade.affectedCount === 1 ? '' : 's'}${examples}`,
        ),
      );
    }
  }

  return lines.join('');
}
