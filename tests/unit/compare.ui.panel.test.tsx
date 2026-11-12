/**
 * Feature 08 — Compare C6: ComparePanel + ChangeList UI.
 * Acceptance T2/T3-UI: renders summary + list; a filter toggle changes only
 * list visibility (summary total unchanged — ADR-0008 guard); cascade rows
 * collapse + expand.
 */
import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { render } from 'preact';
import { act } from 'preact/test-utils';
import { ComparePanel } from '@features/compare/ComparePanel.js';
import { makeStubDriverState, type CompareDriverState } from '@features/compare/ui-types.js';
import { applyFilters } from '@features/compare/components/ChangeList.js';
import { setExportFormat } from '@features/ai-export/index.js';
import type { ChangeSummary, DiffOp, ParaRecord } from '@features/compare/model.js';

let container: HTMLDivElement;
let driver: CompareDriverState;

function op(partial: Partial<DiffOp> & Pick<DiffOp, 'id' | 'type' | 'classification'>): DiffOp {
  return {
    origRange: null,
    revRange: null,
    partId: 'document',
    classes: [partial.classification],
    author: 'Mutatis Compare',
    date: '2026-06-08T10:00:00Z',
    ...partial,
  };
}

function emptySummary(total: number): ChangeSummary {
  return {
    total,
    byCategory: {
      structural: 0,
      obligationLanguage: 0,
      numericalValue: 0,
      definedTerm: 0,
      crossRefRenumber: 0,
      substantiveText: 0,
      whitespacePunct: 0,
      formatting: 0,
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

beforeEach(() => {
  container = document.createElement('div');
  document.body.appendChild(container);
  driver = makeStubDriverState();
});

afterEach(() => {
  document.body.removeChild(container);
});

function mount() {
  act(() => {
    render(<ComparePanel driver={driver} />, container);
  });
  return container;
}

describe('ComparePanel — smoke', () => {
  it('renders without error with the stub driver', () => {
    expect(mount()).toBeTruthy();
  });

  it('shows both source slots with direction', () => {
    const div = mount();
    expect(div.textContent).toContain('Original (base)');
    expect(div.textContent).toContain('Revised (compared)');
  });
});

describe('ComparePanel — Run button gating', () => {
  it('disables Run Compare when status is idle and no sources are ready', () => {
    const div = mount();
    const btn = div.querySelector('.dp-compare__run-btn') as HTMLButtonElement;
    expect(btn).toBeTruthy();
    expect(btn.disabled).toBe(true);
  });

  it('enables Run Compare when both sources are ready', () => {
    driver.originalReady.value = true;
    driver.revisedReady.value = true;
    const div = mount();
    const btn = div.querySelector('.dp-compare__run-btn') as HTMLButtonElement;
    expect(btn.disabled).toBe(false);
  });

  it('shows Cancel only while running', () => {
    driver.status.value = 'running';
    const div = mount();
    expect(div.querySelector('.dp-compare__cancel-btn')).toBeTruthy();
  });
});

describe('ComparePanel — results render', () => {
  beforeEach(() => {
    driver.status.value = 'done';
    const summary = emptySummary(3);
    summary.byCategory.substantiveText = 2;
    summary.byCategory.formatting = 1;
    driver.summary.value = summary;
    driver.ops.value = [
      op({ id: 'a', type: 'replace', classification: 'substantiveText', wordDiffs: [
        { op: 'eq', text: 'The ' }, { op: 'ins', text: 'amended ' }, { op: 'eq', text: 'clause' },
      ] }),
      op({ id: 'b', type: 'replace', classification: 'substantiveText', wordDiffs: [
        { op: 'del', text: 'old' }, { op: 'ins', text: 'new' },
      ] }),
      op({ id: 'c', type: 'replace', classification: 'formatting' }),
    ];
  });

  it('renders the summary total and category counts', () => {
    const div = mount();
    expect(div.textContent).toContain('Total changes:');
    expect(div.textContent).toContain('3');
  });

  it('hides the formatting-only row by default but keeps the summary total', () => {
    const div = mount();
    const rows = div.querySelectorAll('.dp-compare__change-row');
    // 2 substantive shown, 1 formatting hidden by default filter.
    expect(rows.length).toBe(2);
    // Summary total still reflects ALL 3 (ADR-0008: never reduced by a filter).
    const total = div.querySelector('.dp-compare__summary-total-count');
    expect(total?.textContent).toBe('3');
  });

  it('toggling the formatting filter OFF reveals the hidden row; total unchanged', () => {
    mount();
    act(() => {
      driver.hideFormattingOnly.value = false;
      render(<ComparePanel driver={driver} />, container);
    });
    const rows = container.querySelectorAll('.dp-compare__change-row');
    expect(rows.length).toBe(3); // formatting row now visible
    const total = container.querySelector('.dp-compare__summary-total-count');
    expect(total?.textContent).toBe('3'); // total never changes
  });
});

describe('ComparePanel — cascade rows collapse + expand', () => {
  beforeEach(() => {
    driver.status.value = 'done';
    const summary = emptySummary(4);
    summary.byCategory.crossRefRenumber = 3;
    summary.byCategory.substantiveText = 1;
    summary.renumberingCascades = [
      { id: 'casc1', tokenKind: 'Clause', shift: 1, affectedCount: 3, exampleRefs: ['Clause 4.2 → 4.3'] },
    ];
    driver.summary.value = summary;
    driver.ops.value = [
      op({ id: 's1', type: 'replace', classification: 'substantiveText' }),
      op({ id: 'r1', type: 'replace', classification: 'crossRefRenumber', collapsedInto: 'casc1' }),
      op({ id: 'r2', type: 'replace', classification: 'crossRefRenumber', collapsedInto: 'casc1' }),
      op({ id: 'r3', type: 'replace', classification: 'crossRefRenumber', collapsedInto: 'casc1' }),
    ];
  });

  it('renders cascade members collapsed (one header, members hidden)', () => {
    const div = mount();
    // The three collapsed members are NOT top-level rows.
    const rows = div.querySelectorAll('.dp-compare__change-list > .dp-compare__change-row');
    expect(rows.length).toBe(1); // only the substantive op
    // A cascade header exists.
    expect(div.querySelector('.dp-compare__cascade-header')).toBeTruthy();
    // Members are not yet rendered.
    expect(div.querySelector('.dp-compare__cascade-members')).toBeFalsy();
  });

  it('clicking the cascade header expands its members', () => {
    const div = mount();
    const header = div.querySelector('.dp-compare__cascade-header') as HTMLButtonElement;
    act(() => {
      header.click();
    });
    const members = container.querySelector('.dp-compare__cascade-members');
    expect(members).toBeTruthy();
    expect(members!.querySelectorAll('.dp-compare__change-row').length).toBe(3);
  });
});

describe('ComparePanel — change-set Copy-for-AI wiring (feature 11)', () => {
  const writeText = vi.fn((_text: string) => Promise.resolve());
  const para = (index: number, normText: string): ParaRecord =>
    ({ index, partId: 'document', rawXml: '', normText } as ParaRecord);

  beforeEach(() => {
    Object.defineProperty(navigator, 'clipboard', { value: { writeText }, configurable: true });
    writeText.mockClear();
    setExportFormat('json'); // assert against the structured envelope
    driver.status.value = 'done';
    driver.summary.value = emptySummary(2);
    // Pure delete/insert ops — these carry NO wordDiffs in production (those are
    // 'replace'-only). The before/after text must come from the para arrays the
    // panel threads into buildChangeSet; without that wiring it serialises empty.
    driver.originalParas.value = [para(0, 'Kept intro.'), para(1, 'The first clause, since removed.')];
    driver.revisedParas.value = [para(0, 'Kept intro.'), para(1, 'A brand-new inserted clause.')];
    driver.ops.value = [
      op({ id: 'd1', type: 'delete', classification: 'substantiveText', origRange: { start: 1, end: 1 } }),
      op({ id: 'i1', type: 'insert', classification: 'substantiveText', revRange: { start: 1, end: 1 } }),
    ];
  });

  afterEach(() => { setExportFormat('markdown'); });

  it('copies real before/after paragraph text for pure delete/insert (no wordDiffs)', () => {
    mount();
    const btn = container.querySelector('.dp-ai-export__copy-btn') as HTMLButtonElement;
    expect(btn).toBeTruthy();
    act(() => { btn.click(); });

    expect(writeText).toHaveBeenCalledTimes(1);
    const env = JSON.parse(writeText.mock.calls[0]![0]);
    expect(env._mutatis).toBe('1.0');
    const del = env.change_set.changes.find((c: { type: string }) => c.type === 'delete');
    const ins = env.change_set.changes.find((c: { type: string }) => c.type === 'insert');
    expect(del.original).toBe('The first clause, since removed.');
    expect(del.revised).toBeUndefined();
    expect(ins.revised).toBe('A brand-new inserted clause.');
    expect(ins.original).toBeUndefined();
  });
});

describe('applyFilters — pure ADR-0008 guard', () => {
  const ops: DiffOp[] = [
    op({ id: 'sub', type: 'replace', classification: 'substantiveText' }),
    op({ id: 'fmt', type: 'replace', classification: 'formatting' }),
    // A punctuation change as the classifier really emits it: substantive
 //   primary, `whitespacePunct` secondary. `classification` is never the class
 //   itself (ADR-0036), so a fixture claiming it would pin a shape that cannot
 //   occur — which is how the removed "hide whitespace/punctuation" filter kept
 //   looking covered while filtering nothing.
    op({
      id: 'ws', type: 'replace',
      classification: 'substantiveText', classes: ['substantiveText', 'whitespacePunct'],
    }),
    op({ id: 'cell', type: 'replace', classification: 'formatting', partId: 'tableCell:r1c1' }),
    op({ id: 'spell', type: 'replace', classification: 'substantiveText', spellingVariant: { usSpelling: 'organize', gbSpelling: 'organise' } }),
  ];

  it('never mutates the input array', () => {
    const before = [...ops];
    applyFilters(ops, { hideFormattingOnly: true, collapseRenumbering: true });
    expect(ops).toEqual(before);
  });

  it('hides formatting-only when enabled', () => {
    const out = applyFilters(ops, { hideFormattingOnly: true, collapseRenumbering: true });
    expect(out.find(o => o.id === 'fmt')).toBeUndefined();
  });

  it('keeps a punctuation change visible — no filter subtracts it (ADR-0036)', () => {
    // In SG/E&W drafting a comma governs the scope of a proviso, so a change a
    // reviewer must be able to see. `whitespacePunct` is carried as a secondary
    // class for reporting; nothing filters on it.
    const out = applyFilters(ops, { hideFormattingOnly: true, collapseRenumbering: true });
    expect(out.find(o => o.id === 'ws')).toBeDefined();
  });

  it('always keeps table-cell and spelling-variant items visible (§7.1/§7.3)', () => {
    const out = applyFilters(ops, { hideFormattingOnly: true, collapseRenumbering: true });
    expect(out.find(o => o.id === 'cell')).toBeDefined(); // never filtered despite 'formatting'
    expect(out.find(o => o.id === 'spell')).toBeDefined();
  });
});
