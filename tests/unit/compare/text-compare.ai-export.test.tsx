/**
 * Feature 11 / ADR-0021 — "Copy for AI" change-set export on the Text-vs-text
 * Compare panel (parity with the Word-vs-Word panel). Covers the shared builder
 * (`buildChangeSet`) and that the button renders in the text-mode results.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { render } from 'preact';
import { act } from 'preact/test-utils';

import { ComparePanel } from '@features/compare/ComparePanel.js';
import {
  makeStubDriverState,
  makeStubTextDriverState,
  type CompareDriverState,
  type TextCompareDriverState,
} from '@features/compare/ui-types.js';
import { buildChangeSet } from '@features/compare/change-set-export.js';
import type { ChangeSummary, DiffOp, ParaRecord } from '@features/compare/model.js';

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
    classes: [partial.classification], author: 'Mutatis',
    date: '2026-06-19T10:00:00Z', ...partial,
  };
}

function para(index: number, normText: string): ParaRecord {
  return {
    index, partId: 'document', rawXml: `<w:p><w:r><w:t>${normText}</w:t></w:r></w:p>`,
    normText, normTextAggressive: normText.toLowerCase(), numberingPrefix: '',
    hash: `h${index}`, styleId: null, isTableCell: false, isSignatureBlock: false,
    isHeaderFooter: false, isFootnote: false,
  };
}

describe('buildChangeSet (shared Compare → AI export)', () => {
  it('returns null when there are no ops', () => {
    expect(buildChangeSet([], [], [])).toBeNull();
  });

  it('builds a compare envelope carrying the before/after of a change', () => {
    const ops: DiffOp[] = [
      op({ id: 'a', type: 'replace', classification: 'obligationLanguage',
        origRange: { start: 0, end: 0 }, revRange: { start: 0, end: 0 },
        wordDiffs: [{ op: 'del', text: 'shall' }, { op: 'ins', text: 'may' }] }),
    ];
    const env = buildChangeSet(ops, [para(0, 'The Seller shall deliver.')], [para(0, 'The Seller may deliver.')]);
    expect(env).not.toBeNull();
    const json = JSON.stringify(env);
    expect(json).toContain('shall'); // before
    expect(json).toContain('may');   // after
  });
});

describe('TextComparePanel — Copy for AI button', () => {
  let container: HTMLDivElement;
  let driver: CompareDriverState;
  let textDriver: TextCompareDriverState;

  beforeEach(() => {
    container = document.createElement('div');
    document.body.appendChild(container);
    driver = makeStubDriverState();
    textDriver = makeStubTextDriverState();
    textDriver.status.value = 'done';
    textDriver.redlineReady.value = true;
    textDriver.summary.value = emptySummary(1);
    textDriver.ops.value = [
      op({ id: 'a', type: 'replace', classification: 'substantiveText',
        origRange: { start: 0, end: 0 }, revRange: { start: 0, end: 0 },
        wordDiffs: [{ op: 'del', text: 'old' }, { op: 'ins', text: 'new' }] }),
    ];
    textDriver.originalParas.value = [para(0, 'old value')];
    textDriver.revisedParas.value = [para(0, 'new value')];
  });

  afterEach(() => {
    document.body.removeChild(container);
  });

  it('renders a "Copy for AI" button in the text-mode results', () => {
    act(() => {
      render(<ComparePanel driver={driver} textDriver={textDriver} />, container);
    });
    const textBtn = [...container.querySelectorAll('.dp-compare__mode-btn')]
      .find((b) => /pasted text/i.test(b.textContent ?? '')) as HTMLButtonElement;
    act(() => { textBtn.click(); });

    const copyBtn = container.querySelector('.dp-ai-export__copy-btn');
    expect(copyBtn).toBeTruthy();
    expect(copyBtn?.textContent).toContain('Copy for AI');
  });
});
