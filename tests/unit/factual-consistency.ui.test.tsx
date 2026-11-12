import { describe, it, expect, beforeEach } from 'vitest';
import { render } from 'preact';
import { FindingsList } from '@taskpane/shell/FindingsList.js';
import { addFinding, clearFindings } from '@shared/findings/index.js';

function mount() {
  const div = document.createElement('div');
  render(<FindingsList tabId="proof" />, div);
  return div;
}

describe('T8 UI — Proof tab action bar and finding rows', () => {
  beforeEach(() => clearFindings());

  it('renders "Run Proof checks" button on the proof tab', () => {
    const div = mount();
    const btn = div.querySelector('button.dp-proof-actions__run');
    expect(btn).toBeTruthy();
    expect(btn?.textContent).toContain('Run Proof checks');
  });

  it('does NOT render run button on other tabs', () => {
    const div = document.createElement('div');
    render(<FindingsList tabId="review" />, div);
    expect(div.querySelector('button.dp-proof-actions__run')).toBeNull();
  });

  it('ambiguous-currency finding gets check class, not info class', () => {
    addFinding({
      id: 'factual-consistency/ambiguous-currency/p0-10',
      category: 'factual-consistency/ambiguous-currency',
      severity: 'info',
      title: "Bare '$' — currency not specified",
      location: { label: 'p0' },
    });
    const div = mount();
    const row = div.querySelector('.dp-finding-row--check');
    expect(row).toBeTruthy();
    expect(div.querySelector('.dp-finding-row--info')).toBeNull();
  });

  it('error finding keeps error class (not check)', () => {
    addFinding({
      id: 'factual-consistency/amount-mismatch/p0-0',
      category: 'factual-consistency/amount-mismatch',
      severity: 'error',
      title: 'Amount mismatch',
      location: { label: 'p0' },
    });
    const div = mount();
    expect(div.querySelector('.dp-finding-row--error')).toBeTruthy();
    expect(div.querySelector('.dp-finding-row--check')).toBeNull();
  });

  it('AC-1: suggested-fix button is present but does not auto-apply on render', () => {
    const applyFn = () => Promise.resolve();
    addFinding({
      id: 'factual-consistency/amount-mismatch/p0-0',
      category: 'factual-consistency/amount-mismatch',
      severity: 'error',
      title: 'Amount mismatch',
      location: { label: 'p0' },
      suggestedAction: { label: 'View suggested corrections', apply: applyFn },
    });
    const div = mount();
    // Button exists but apply() has NOT been called (no auto-edit)
    const actionBtn = div.querySelector('.dp-finding-row__action');
    expect(actionBtn).toBeTruthy();
    expect(actionBtn?.textContent).toContain('View suggested corrections');
  });

  it('AC-2: ambiguous-currency finding has info severity, not error', () => {
    addFinding({
      id: 'factual-consistency/ambiguous-currency/p0-5',
      category: 'factual-consistency/ambiguous-currency',
      severity: 'info',
      title: "Bare '$' — currency not specified",
      location: { label: 'p0' },
    });
    const div = mount();
    // The finding row should NOT have the --error class
    expect(div.querySelector('.dp-finding-row--error')).toBeNull();
    // It should have the --check class (amber)
    expect(div.querySelector('.dp-finding-row--check')).toBeTruthy();
  });

  it('category label shows correct human-readable text', () => {
    addFinding({
      id: 'factual-consistency/placeholder/p0-0-TBC',
      category: 'factual-consistency/placeholder',
      severity: 'info',
      title: 'Unfinished placeholder [TBC]',
      location: { label: 'p0' },
    });
    const div = mount();
    expect(div.textContent).toContain('Unfilled placeholders');
  });
});
