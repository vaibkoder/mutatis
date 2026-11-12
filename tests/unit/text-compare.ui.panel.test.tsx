/**
 * Feature 12 (Compare — additional sources) T4: Text-vs-text panel UI.
 * Spec: spec/features/10-compare-additional-sources.md §4.1 · ADR-0018.
 *
 * Renders against a STUB text driver (no real pipeline). Asserts: the mode toggle
 * switches to the text panel; two text areas appear; the Open-redline button is
 * gated on redlineReady and calls openRedline() only on click; the summary +
 * filtered change list reuse the feature-08 components.
 */
import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { render } from 'preact';
import { act } from 'preact/test-utils';

import { ComparePanel } from '@features/compare/ComparePanel.js';
import {
  makeStubDriverState,
  makeStubTextDriverState,
  type CompareDriverState,
  type TextCompareDriverState,
} from '@features/compare/ui-types.js';
import type { ChangeSummary, DiffOp } from '@features/compare/model.js';

let container: HTMLDivElement;
let driver: CompareDriverState;
let textDriver: TextCompareDriverState;

function emptySummary(total: number): ChangeSummary {
  return {
    total,
    byCategory: {
      structural: 0, obligationLanguage: 0, numericalValue: 0, definedTerm: 0,
      crossRefRenumber: 0, substantiveText: 0, whitespacePunct: 0, formatting: 0,
    },
    renumberingCascades: [],
    movedParagraphs: 0,
    partChanges: { headers: 0, footers: 0, footnotes: 0 },
    styleDefinitionNote: null,
    spellingVariantCount: 0,
    quarantinedCommentsOriginal: 0,
    quarantinedCommentsRevised: 0,
    filtersActive: { hideFormattingOnly: true, collapseRenumbering: true },
  };
}

function op(partial: Partial<DiffOp> & Pick<DiffOp, 'id' | 'type' | 'classification'>): DiffOp {
  return {
    origRange: null, revRange: null, partId: 'document',
    classes: [partial.classification], author: 'Mutatis Compare',
    date: '2026-06-08T10:00:00Z', ...partial,
  };
}

beforeEach(() => {
  container = document.createElement('div');
  document.body.appendChild(container);
  driver = makeStubDriverState();
  textDriver = makeStubTextDriverState();
});

afterEach(() => {
  document.body.removeChild(container);
});

function mount() {
  act(() => {
    render(<ComparePanel driver={driver} textDriver={textDriver} />, container);
  });
  return container;
}

function switchToTextMode() {
  const btns = [...container.querySelectorAll('.dp-compare__mode-btn')] as HTMLButtonElement[];
  const textBtn = btns.find(b => /pasted text/i.test(b.textContent ?? ''))!;
  act(() => { textBtn.click(); });
}

describe('ComparePanel — mode toggle', () => {
  it('defaults to Word-documents mode (slots visible, no text areas)', () => {
    const div = mount();
    expect(div.textContent).toContain('Original (base)');
    expect(div.querySelector('.dp-compare__text-input')).toBeFalsy();
  });

  it('switching to Pasted-text reveals two text areas', () => {
    mount();
    switchToTextMode();
    const areas = container.querySelectorAll('.dp-compare__text-input');
    expect(areas.length).toBe(2);
  });
});

describe('TextComparePanel — input + run', () => {
  it('typing into a text area updates the driver signal', () => {
    mount();
    switchToTextMode();
    const area = container.querySelector('.dp-compare__text-input') as HTMLTextAreaElement;
    act(() => {
      area.value = 'Hello clause.';
      area.dispatchEvent(new Event('input', { bubbles: true }));
    });
    expect(textDriver.originalText.value).toBe('Hello clause.');
  });

  it('Compare button calls runCompare', () => {
    const spy = vi.spyOn(textDriver, 'runCompare');
    mount();
    switchToTextMode();
    const btn = container.querySelector('.dp-compare__run-btn') as HTMLButtonElement;
    act(() => { btn.click(); });
    expect(spy).toHaveBeenCalledTimes(1);
  });
});

describe('TextComparePanel — results + click-to-open', () => {
  beforeEach(() => {
    textDriver.status.value = 'done';
    textDriver.redlineReady.value = true;
    const summary = emptySummary(2);
    summary.byCategory.substantiveText = 1;
    summary.byCategory.formatting = 1;
    textDriver.summary.value = summary;
    textDriver.ops.value = [
      op({ id: 'a', type: 'replace', classification: 'substantiveText',
        wordDiffs: [{ op: 'del', text: 'shall' }, { op: 'ins', text: 'may' }] }),
      op({ id: 'b', type: 'replace', classification: 'formatting' }),
    ];
  });

  it('renders the reused summary total and an Open-redline button', () => {
    mount();
    switchToTextMode();
    expect(container.textContent).toContain('Total changes:');
    expect(container.querySelector('.dp-compare__open-redline-btn')).toBeTruthy();
  });

  it('the redline opens only on click (openRedline not called during render)', () => {
    const spy = vi.spyOn(textDriver, 'openRedline');
    mount();
    switchToTextMode();
    // Not called just by rendering the results.
    expect(spy).not.toHaveBeenCalled();
    const btn = container.querySelector('.dp-compare__open-redline-btn') as HTMLButtonElement;
    act(() => { btn.click(); });
    expect(spy).toHaveBeenCalledTimes(1);
  });

  it('Open-redline button is disabled until redlineReady is true', () => {
    textDriver.redlineReady.value = false;
    mount();
    switchToTextMode();
    const btn = container.querySelector('.dp-compare__open-redline-btn') as HTMLButtonElement;
    expect(btn.disabled).toBe(true);
  });

  it('reuses the change list with default filters (formatting row hidden, total intact)', () => {
    mount();
    switchToTextMode();
    const rows = container.querySelectorAll('.dp-compare__change-row');
    // 1 substantive shown; 1 formatting hidden by default filter.
    expect(rows.length).toBe(1);
    // ADR-0008: summary total still reflects both changes.
    const total = container.querySelector('.dp-compare__summary-total-count');
    expect(total?.textContent).toBe('2');
  });
});
