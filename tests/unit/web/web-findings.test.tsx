/**
 * apps/web — WebFindings view.
 *
 * Groups a flat findings list into the four display features (cross-reference,
 * static-numbering, defined-terms, factual-consistency) reusing the shared
 * findings-grouping helpers + the host-neutral FindingRow, and renders the two
 * banners. No Office anywhere.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { render } from 'preact';
import { act } from 'preact/test-utils';
import type { Finding } from '@shared/findings/index.js';
import type { DefinedTermsAnalysis } from '@features/defined-terms/index.js';
import type { CrossReferenceAnalysis } from '@features/cross-reference/index.js';
import { WebFindings } from '../../../src/web/WebFindings.js';

let container: HTMLDivElement;

beforeEach(() => {
  container = document.createElement('div');
  document.body.appendChild(container);
});

afterEach(() => {
  render(null, container);
  container.remove();
});

function finding(id: string, category: string, featureId: string, title: string): Finding {
  return {
    id,
    category,
    featureId,
    severity: 'warning',
    title,
    location: { label: 'p0' },
  } as Finding;
}

const FINDINGS: Finding[] = [
  finding('x1', 'cross-reference/not-found', 'xref', 'Not found: "Clause 9"'),
  finding('s1', 'static-numbering/manually-typed', 'staticNumbering', 'Manually typed: "4.2"'),
  finding('d1', 'defined-terms/defined-not-used', 'definedTerms', 'Defined but not used: "Widget"'),
  finding('f1', 'factual-consistency/placeholder', 'factual', 'Unfilled placeholder: "[●]"'),
];

const DEFINED = { noDefinitions: false, crossRefCount: 2, crossDocCount: 0 } as DefinedTermsAnalysis;
const XREF = { fieldsApiUnavailable: true } as CrossReferenceAnalysis;

describe('WebFindings', () => {
  it('renders every feature group with the shared labels + banners', () => {
    act(() => {
      render(<WebFindings findings={FINDINGS} defined={DEFINED} xref={XREF} />, container);
    });
    const text = container.textContent ?? '';

    // The four feature groups.
    expect(text).toContain('Cross-references');
    expect(text).toContain('Manually-typed clause numbers');
    expect(text).toContain('Defined terms');
    // Factual-consistency category label (from the shared CATEGORY_LABELS).
    expect(text).toContain('Unfilled placeholders');

    // Banners.
    expect(text).toContain('defined by cross-reference within this document');
    expect(text).toContain('REF-field detection is not available outside Word');

    // One FindingRow per finding.
    expect(container.querySelectorAll('.dp-finding-row')).toHaveLength(4);
  });

  it('shows a no-issues message and no banners when empty', () => {
    act(() => {
      render(<WebFindings findings={[]} defined={null} xref={null} />, container);
    });
    expect(container.textContent).toContain('No issues found');
    expect(container.querySelector('.dp-capability-banner')).toBeNull();
    expect(container.querySelectorAll('.dp-finding-row')).toHaveLength(0);
  });
});
