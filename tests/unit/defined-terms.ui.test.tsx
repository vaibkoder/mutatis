import { describe, it, expect, beforeEach, vi } from 'vitest';
import { render } from 'preact';
import { act } from 'preact/test-utils';
import { FindingsList } from '@taskpane/shell/FindingsList.js';
import { addFinding, clearFindings } from '@shared/findings/index.js';
import { DocModel } from '@shared/doc-model/index.js';

// DocModel.load() calls Word.run which is unavailable in jsdom; stub it out.
vi.spyOn(DocModel, 'load').mockResolvedValue(undefined);

function mount(tabId: 'proof' | 'review' = 'proof') {
  const div = document.createElement('div');
  act(() => { render(<FindingsList tabId={tabId} />, div); });
  return div;
}

function click(el: Element | null) {
  act(() => { el?.dispatchEvent(new MouseEvent('click', { bubbles: true })); });
}

function loc(label: string) {
  return { label };
}

describe('Defined-Terms UI wiring (T10)', () => {
  beforeEach(() => clearFindings());

  it('renders the corrected capitalisation label (capitalisation, not capitalisation-drift)', () => {
    addFinding({
      id: 'defined-terms/capitalisation/p0-0-5',
      category: 'defined-terms/capitalisation',
      severity: 'info',
      title: '"property" is lowercase but "Property" is a defined term',
      location: { label: 'p0' },
    });
    const div = mount();
    expect(div.textContent).toContain('Capitalisation inconsistency');
  });

  it('source snippet toggle: clicking location button expands and collapses the snippet panel', () => {
    addFinding({
      id: 'defined-terms/capitalisation/p1-0-5',
      category: 'defined-terms/capitalisation',
      severity: 'info',
      title: '"property" is lowercase but "Property" is a defined term',
      location: { label: 'Property' },
      sourceSnippet: { label: 'Property', text: '"Property" means the land and buildings.' },
    });
    const div = mount();

    // No snippet panel before toggle.
    expect(div.querySelector('.dp-source-snippet')).toBeNull();

    // Click the location toggle button.
    click(div.querySelector('.dp-finding-row__location--src'));
    expect(div.querySelector('.dp-source-snippet')).toBeTruthy();
    expect(div.querySelector('.dp-source-snippet__text')?.textContent).toContain('"Property" means');

    // Click again — panel collapses.
    click(div.querySelector('.dp-finding-row__location--src'));
    expect(div.querySelector('.dp-source-snippet')).toBeNull();
  });

  it('shows a smart/strict capitalisation toggle on the proof tab that flips on click', () => {
    const div = mount();
    const toggle = div.querySelector('button.dp-cap-toggle');
    expect(toggle).toBeTruthy();
    expect(toggle?.textContent).toContain('Smart');
    click(toggle);
    expect(div.querySelector('button.dp-cap-toggle')?.textContent).toContain('Strict');
  });

  it('does not render the toggle on non-proof tabs', () => {
    const div = mount('review');
    expect(div.querySelector('button.dp-cap-toggle')).toBeNull();
  });

  it('running Proof checks renders the defined-terms banner', async () => {
    const div = mount();
    expect(div.querySelector('.dp-defined-terms-banner')).toBeNull();
    // Dispatch click — starts the async handleRunProof.
    div.querySelector('button.dp-proof-actions__run')?.dispatchEvent(
      new MouseEvent('click', { bubbles: true }),
    );
    // Flush all pending microtasks (allows DocModel.load mock + setAnalysis to run),
    // then flush Preact state updates via act.
    await new Promise<void>(resolve => setTimeout(resolve, 0));
    await act(async () => {});
    // With an empty in-memory document the banner reports "No definitions detected".
    const banner = div.querySelector('.dp-defined-terms-banner');
    expect(banner).toBeTruthy();
    expect(banner?.textContent).toContain('No definitions detected');
  });
});

describe('Defined-Terms occurrence counter + Next + sort', () => {
  beforeEach(() => clearFindings());

  it('renders a count chip for a multi-occurrence finding', () => {
    addFinding({
      id: 'defined-terms/used-not-defined/a',
      category: 'defined-terms/used-not-defined',
      severity: 'error',
      title: '"Disclosure Letter" is used but never defined',
      location: loc('p1'),
      occurrences: [loc('p1'), loc('p2'), loc('p3')],
    });
    const div = mount();
    expect(div.querySelector('.dp-finding-row__count')?.textContent).toContain('3×');
  });

  it('renders a Next control whose position advances on click', () => {
    addFinding({
      id: 'defined-terms/used-not-defined/b',
      category: 'defined-terms/used-not-defined',
      severity: 'error',
      title: '"Disclosure Letter" is used but never defined',
      location: loc('p1'),
      occurrences: [loc('p1'), loc('p2'), loc('p3')],
    });
    const div = mount();
    const next = div.querySelector('.dp-finding-row__next');
    expect(next).toBeTruthy();
    expect(next?.textContent).toContain('1/3');
    click(next);
    expect(div.querySelector('.dp-finding-row__next')?.textContent).toContain('2/3');
  });

  it('does not render Next for a single-occurrence finding', () => {
    addFinding({
      id: 'defined-terms/used-not-defined/c',
      category: 'defined-terms/used-not-defined',
      severity: 'error',
      title: '"Solo Term" is used but never defined',
      location: loc('p1'),
    });
    const div = mount();
    expect(div.querySelector('.dp-finding-row__next')).toBeNull();
  });

  it('does not render the filter bar when there are no findings', () => {
    const div = mount();
    expect(div.querySelector('.dp-findings-search')).toBeNull();
  });

  it('renders the filter bar once findings exist and narrows results by query', () => {
    addFinding({
      id: 'defined-terms/used-not-defined/disclosure',
      category: 'defined-terms/used-not-defined',
      severity: 'error',
      title: '"Disclosure Letter" is used but never defined',
      location: loc('p1'),
    });
    addFinding({
      id: 'defined-terms/used-not-defined/completion',
      category: 'defined-terms/used-not-defined',
      severity: 'error',
      title: '"Completion Date" is used but never defined',
      location: loc('p2'),
    });
    const div = mount();
    const input = div.querySelector('.dp-findings-search__input') as HTMLInputElement | null;
    expect(input).toBeTruthy();
    // Both findings visible before filtering.
    expect(div.textContent).toContain('Disclosure Letter');
    expect(div.textContent).toContain('Completion Date');

    // Type a query that only matches one finding.
    act(() => {
      input!.value = 'completion';
      input!.dispatchEvent(new Event('input', { bubbles: true }));
    });
    expect(div.textContent).toContain('Completion Date');
    expect(div.textContent).not.toContain('Disclosure Letter');
    expect(div.querySelector('.dp-findings-search__count')?.textContent).toContain('1 of 2');
  });

  it('matches clause references by location label and shows a no-match message', () => {
    addFinding({
      id: 'cross-reference/not-found/x',
      category: 'cross-reference/not-found',
      severity: 'error',
      title: 'Reference target not found',
      location: loc('Clause 4.2'),
    });
    const div = mount();
    const input = div.querySelector('.dp-findings-search__input') as HTMLInputElement;

    // Matches on the clause-reference location label.
    act(() => { input.value = '4.2'; input.dispatchEvent(new Event('input', { bubbles: true })); });
    expect(div.textContent).toContain('Reference target not found');

    // A query matching nothing shows the no-match message, not "No issues found".
    act(() => { input.value = 'zzz-nope'; input.dispatchEvent(new Event('input', { bubbles: true })); });
    expect(div.textContent).toContain('No findings match');
    expect(div.textContent).not.toContain('No issues found');
  });

  it('filter also matches on detail text and occurrence labels', () => {
    addFinding({
      id: 'defined-terms/used-not-defined/detail-match',
      category: 'defined-terms/used-not-defined',
      severity: 'error',
      title: '"Alpha" is used but never defined',
      detail: 'first appears in the indemnity wording',
      location: loc('p1'),
    });
    addFinding({
      id: 'defined-terms/used-not-defined/occ-match',
      category: 'defined-terms/used-not-defined',
      severity: 'error',
      title: '"Beta" is used but never defined',
      location: loc('p1'),
      occurrences: [loc('p1'), loc('Clause 9.3')],
    });
    const div = mount();
    const input = div.querySelector('.dp-findings-search__input') as HTMLInputElement;

    // Matches on detail text only (no title/label contains "indemnity").
    act(() => { input.value = 'indemnity'; input.dispatchEvent(new Event('input', { bubbles: true })); });
    expect(div.textContent).toContain('Alpha');
    expect(div.textContent).not.toContain('Beta');

    // Matches on a secondary occurrence label ("Clause 9.3"), not the primary one.
    act(() => { input.value = '9.3'; input.dispatchEvent(new Event('input', { bubbles: true })); });
    expect(div.textContent).toContain('Beta');
    expect(div.textContent).not.toContain('Alpha');
  });

  it('frequency sort toggle reorders findings most-used first', () => {
    addFinding({
      id: 'defined-terms/used-not-defined/low',
      category: 'defined-terms/used-not-defined',
      severity: 'error',
      title: '"Alpha" is used but never defined',
      location: loc('p1'),
      occurrences: [loc('p1'), loc('p2')],
    });
    addFinding({
      id: 'defined-terms/used-not-defined/high',
      category: 'defined-terms/used-not-defined',
      severity: 'error',
      title: '"Beta" is used but never defined',
      location: loc('p1'),
      occurrences: [loc('p1'), loc('p2'), loc('p3'), loc('p4')],
    });
    const div = mount();
    const toggle = div.querySelector('button.dp-freq-sort');
    expect(toggle).toBeTruthy();
    click(toggle);
    const titles = Array.from(div.querySelectorAll('.dp-finding-row__title')).map(e => e.textContent);
    const alphaIdx = titles.findIndex(t => t?.includes('Alpha'));
    const betaIdx = titles.findIndex(t => t?.includes('Beta'));
    expect(betaIdx).toBeLessThan(alphaIdx); // Beta (4×) before Alpha (2×)
  });
});
