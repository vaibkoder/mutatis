/**
 * Feature 08 — Compare · classifier taxonomy tests (Task B1).
 *
 * Covers T2 (formatting-only), T5 (obligation language + priority), and the
 * multi-class rule. Fixtures are hand-built DiffOps — no dependency on the diff
 * or preprocess pipelines.
 */

import { describe, it, expect } from 'vitest';
import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import type { CompareConfig } from '@shared/config/index.js';
import type { DiffOp, ParaRecord, WordDiff } from '../../src/features/compare/model.js';
import { classifyOp, classifyAll } from '../../src/features/compare/classify/classifier.js';
import { classifyClasses } from '../../src/features/compare/classify/taxonomy.js';

const CONFIG: CompareConfig = getCompareConfig(SG_PROFILE);

// ---------------------------------------------------------------------------
// Fixture helpers
// ---------------------------------------------------------------------------

function para(overrides: Partial<ParaRecord> = {}): ParaRecord {
  return {
    index: 0,
    partId: 'document',
    rawXml: '<w:p/>',
    normText: '',
    normTextAggressive: '',
    numberingPrefix: '',
    hash: '0',
    styleId: null,
    isTableCell: false,
    isSignatureBlock: false,
    isHeaderFooter: false,
    isFootnote: false,
    ...overrides,
  };
}

function replaceOp(wordDiffs: WordDiff[], overrides: Partial<DiffOp> = {}): DiffOp {
  return {
    id: 'op-1',
    type: 'replace',
    origRange: { start: 0, end: 0 },
    revRange: { start: 0, end: 0 },
    partId: 'document',
    wordDiffs,
    classification: 'substantiveText',
    classes: ['substantiveText'],
    author: 'Test',
    date: '2026-06-08T00:00:00Z',
    ...overrides,
  };
}

// ---------------------------------------------------------------------------
// T2 — formatting-only
// ---------------------------------------------------------------------------

describe('taxonomy — T2 formatting-only', () => {
  it('a style-only delta (no text change) classifies as formatting and is still present', () => {
    // No non-equal WordDiff text → zero text delta → formatting-only.
    const op = replaceOp([{ op: 'eq', text: 'The heading' }], {
      origRange: { start: 0, end: 0 },
      revRange: { start: 0, end: 0 },
    });
    const orig = [para({ normText: 'The heading', normTextAggressive: 'the heading' })];
    const rev = [para({ normText: 'The heading', normTextAggressive: 'the heading', styleId: 'Heading1' })];

    const out = classifyOp(op, orig, rev, CONFIG);
    expect(out.classification).toBe('formatting');
    expect(out.classes).toEqual(['formatting']);
    // ADR-0008: the op is never dropped.
    expect(out.id).toBe('op-1');
  });

  it('an op with no wordDiffs at all is formatting-only, not dropped', () => {
    const op = replaceOp([], {});
    const out = classifyOp(op, [para()], [para()], CONFIG);
    expect(out.classification).toBe('formatting');
    expect(out.classes).toContain('formatting');
  });
});

// ---------------------------------------------------------------------------
// T5 — obligation language + priority
// ---------------------------------------------------------------------------

describe('taxonomy — T5 obligation language', () => {
  it('shall → may classifies as obligationLanguage (primary)', () => {
    const op = replaceOp([
      { op: 'del', text: 'shall' },
      { op: 'ins', text: 'may' },
    ]);
    const orig = [para({ normText: 'The tenant shall pay' })];
    const rev = [para({ normText: 'The tenant may pay' })];

    const out = classifyOp(op, orig, rev, CONFIG);
    expect(out.classification).toBe('obligationLanguage');
    expect(out.classes[0]).toBe('obligationLanguage');
    // It is also substantive text (multi-class), but obligation wins primary.
    expect(out.classes).toContain('substantiveText');
  });

  it('detects "shall not" as a whole modal phrase, case-insensitively', () => {
    const op = replaceOp([{ op: 'del', text: 'Shall Not' }, { op: 'ins', text: 'must' }]);
    const out = classifyOp(op, [para()], [para()], CONFIG);
    expect(out.classes).toContain('obligationLanguage');
  });

  it('does not flag a substring like "marshall" as a modal', () => {
    const op = replaceOp([{ op: 'ins', text: 'marshall the evidence' }]);
    const out = classifyOp(op, [para()], [para()], CONFIG);
    expect(out.classes).not.toContain('obligationLanguage');
    expect(out.classification).toBe('substantiveText');
  });
});

// ---------------------------------------------------------------------------
// Priority — obligation (2) beats numerical (3)
// ---------------------------------------------------------------------------

describe('taxonomy — priority resolution', () => {
  it('an op changing both a modal AND an amount picks obligationLanguage as primary', () => {
    const op = replaceOp([
      { op: 'del', text: 'shall' },
      { op: 'ins', text: 'may' },
      { op: 'del', text: '£40,000', atomic: 'amount' },
      { op: 'ins', text: '£50,000', atomic: 'amount' },
    ]);
    const out = classifyOp(op, [para()], [para()], CONFIG);
    expect(out.classification).toBe('obligationLanguage');
    // Both categories are recorded in classes (multi-class), priority-ordered.
    expect(out.classes).toContain('obligationLanguage');
    expect(out.classes).toContain('numericalValue');
    expect(out.classes.indexOf('obligationLanguage')).toBeLessThan(
      out.classes.indexOf('numericalValue'),
    );
  });

  it('structural (insert) outranks everything else', () => {
    const op = replaceOp([{ op: 'ins', text: 'A new Clause shall apply' }], {
      type: 'insert',
      origRange: null,
      revRange: { start: 0, end: 0 },
    });
    const out = classifyOp(op, [], [para({ normText: 'A new Clause shall apply' })], CONFIG);
    expect(out.classification).toBe('structural');
    expect(out.classes).toContain('obligationLanguage');
  });

  it('classes are returned in canonical priority order', () => {
    const op = replaceOp([
      { op: 'del', text: 'shall' },
      { op: 'ins', text: 'may' },
      { op: 'del', text: '£40,000', atomic: 'amount' },
      { op: 'ins', text: '£50,000', atomic: 'amount' },
    ]);
    // Give the two sides distinct aggressive-normalised text so this is not
    // also a whitespacePunct case (which would legitimately add an 8th class).
    const orig = [para({ normTextAggressive: 'tenant shall pay 40000' })];
    const rev = [para({ normTextAggressive: 'tenant may pay 50000' })];
    const classes = classifyClasses(op, orig, rev);
    // obligationLanguage (2) before numericalValue (3) before substantiveText (6).
    expect(classes).toEqual(['obligationLanguage', 'numericalValue', 'substantiveText']);
  });
});

// ---------------------------------------------------------------------------
// numericalValue / definedTerm signals
// ---------------------------------------------------------------------------

describe('taxonomy — numerical & defined-term signals', () => {
  it('an amount atomic change classifies as numericalValue', () => {
    const op = replaceOp([
      { op: 'del', text: '£40,000', atomic: 'amount' },
      { op: 'ins', text: '£50,000', atomic: 'amount' },
    ]);
    const out = classifyOp(op, [para()], [para()], CONFIG);
    expect(out.classification).toBe('numericalValue');
  });

  it('a definedTerm atomic change classifies as definedTerm', () => {
    const op = replaceOp([
      { op: 'del', text: '"Permitted Encumbrance"', atomic: 'definedTerm' },
      { op: 'ins', text: '"Permitted Security"', atomic: 'definedTerm' },
    ]);
    const out = classifyOp(op, [para()], [para()], CONFIG);
    expect(out.classification).toBe('definedTerm');
  });

  it('a Title-Case multi-word term inside the CHANGE flags definedTerm', () => {
    const op = replaceOp([
      { op: 'del', text: 'Permitted Encumbrance' },
      { op: 'ins', text: 'Permitted Security' },
    ]);
    const orig = [para({ normText: 'Each Permitted Encumbrance applies here' })];
    const rev = [para({ normText: 'Each Permitted Security applies here' })];
    expect(classifyOp(op, orig, rev, CONFIG).classification).toBe('definedTerm');
  });

  it('a term merely MENTIONED in the paragraph does not flag definedTerm', () => {
    // §6 row 4 says "change TOUCHES a ... term". Testing the paragraph instead
    // made almost every edit a defined-term change: `TITLE_CASE_MULTIWORD`
    // matches a sentence-initial capital plus a capitalised word, so `The
    // Vendor shall ...` — how most clauses in an SG/E&W contract open —
    // matched, whatever the edit actually was.
    const op = replaceOp([{ op: 'ins', text: ' extra words' }]);
    const orig = [para({ normText: 'The Permitted Encumbrance applies here' })];
    const rev = [para({ normText: 'The Permitted Encumbrance applies here extra words' })];
    const out = classifyOp(op, orig, rev, CONFIG);
    expect(out.classes).not.toContain('definedTerm');
    expect(out.classification).toBe('substantiveText');
  });

  it('any edit to the paragraph that DEFINES a term flags definedTerm', () => {
    // §6 row 4's second limb. Here the paragraph-level test is right: an edit
    // anywhere in a definition is an edit to the definition.
    const definition = '"Losses" means all losses, damages and costs.';
    const op = replaceOp([
      { op: 'del', text: 'and costs' },
      { op: 'ins', text: 'costs and expenses' },
    ]);
    const orig = [para({ normText: definition })];
    const rev = [para({ normText: '"Losses" means all losses, damages costs and expenses.' })];
    expect(classifyOp(op, orig, rev, CONFIG).classification).toBe('definedTerm');
  });

  it('a pure renumbering classifies the same whatever prose surrounds it', () => {
    // The defect this replaced: row 4 outranks row 5, so a paragraph-wide
    // defined-term test stopped a pure renumbering ever classifying as
    // `crossRefRenumber` — and `computeCascades` only collapses ops that do.
    // Identical edits must not classify differently because of words they never
    // touched.
    const renumber: WordDiff[] = [
      { op: 'eq', text: 'under Clause ' },
      { op: 'del', text: '4.2', atomic: 'clauseRef' },
      { op: 'ins', text: '4.3', atomic: 'clauseRef' },
      { op: 'eq', text: ' above' },
    ];
    const prose = [
      'subject to the limits under Clause 4.2 above',
      'The Vendor shall comply with the limits under Clause 4.2 above',
      'each Permitted Encumbrance is subject to the limits under Clause 4.2 above',
    ];
    for (const text of prose) {
      const out = classifyOp(
        replaceOp(renumber),
        [para({ normText: text })],
        [para({ normText: text.replace('4.2', '4.3') })],
        CONFIG,
      );
      expect(out.classification, text).toBe('crossRefRenumber');
    }
  });
});

// ---------------------------------------------------------------------------
// whitespacePunct
// ---------------------------------------------------------------------------

describe('taxonomy — whitespacePunct', () => {
  /**
   * `whitespacePunct` is a real class, and it is **never** a primary
   * classification. The proof is two lines of `taxonomy.ts`:
   *
   *   `isWhitespacePunct` returns false unless `hasTextDelta(op)`   (line ~175)
   *   `isSubstantiveText(op)` **is** `hasTextDelta(op)`             (line ~161)
   *
   * so matching `whitespacePunct` (priority 7) implies matching
   * `substantiveText` (priority 6), which always wins. Nothing can read
   * `op.classification === 'whitespacePunct'` and see a hit — which is why the
   * "hide whitespace/punctuation" filter that did exactly that never filtered
   * anything, and was removed rather than rewired (ADR-0036).
   *
   * Rewiring it to read `op.classes` was rejected: an op carries BOTH classes,
   * so hiding on the secondary would hide a change that is also
   * `substantiveText`. In SG/E&W drafting a comma governs the scope of a
   * proviso, so that is a change a reviewer must see.
   */
  it('records the class on a punctuation-only delta, but never as the primary', () => {
    const op = replaceOp([
      { op: 'del', text: 'co-operate' },
      { op: 'ins', text: 'cooperate' },
    ]);
    const orig = [para({ normText: 'parties co-operate', normTextAggressive: 'parties cooperate' })];
    const rev = [para({ normText: 'parties cooperate', normTextAggressive: 'parties cooperate' })];
    const out = classifyOp(op, orig, rev, CONFIG);
    expect(out.classes).toContain('whitespacePunct');
    expect(out.classification).toBe('substantiveText');
    expect(out.classification).not.toBe('whitespacePunct');
  });

  it('no op that matches whitespacePunct can ever be classified as it', () => {
    // Every shape that reaches the class: a punctuation swap, a punctuation
    // deletion, an insertion, and a multi-record range. If any of these could
    // make `whitespacePunct` primary, a filter reading `classification` would
    // have something to hide — none can.
    const shapes: Array<{ label: string; diffs: WordDiff[] }> = [
      { label: 'swap', diffs: [{ op: 'del', text: 'co-operate' }, { op: 'ins', text: 'cooperate' }] },
      { label: 'comma dropped', diffs: [{ op: 'del', text: ',' }] },
      { label: 'semicolon added', diffs: [{ op: 'ins', text: ';' }] },
    ];
    for (const { label, diffs } of shapes) {
      const op = replaceOp(diffs);
      // Aggressive normalisation identical on both sides — the class's own
      // precondition — so this is the most favourable possible input for it.
      const orig = [para({ normText: 'a, b', normTextAggressive: 'a b' })];
      const rev = [para({ normText: 'a b', normTextAggressive: 'a b' })];
      const out = classifyOp(op, orig, rev, CONFIG);
      expect(out.classification, label).not.toBe('whitespacePunct');
    }
  });
});

// ---------------------------------------------------------------------------
// classifyAll — purity & length preservation (ADR-0008)
// ---------------------------------------------------------------------------

describe('classifyAll — purity and never-drop', () => {
  it('returns a new array of the same length without mutating inputs', () => {
    const ops: DiffOp[] = [
      replaceOp([{ op: 'ins', text: 'hello' }], { id: 'a' }),
      replaceOp([], { id: 'b' }),
      replaceOp([{ op: 'del', text: 'shall' }, { op: 'ins', text: 'may' }], { id: 'c' }),
    ];
    const snapshot = JSON.parse(JSON.stringify(ops));
    const out = classifyAll(ops, [para()], [para()], CONFIG);

    expect(out).toHaveLength(3);
    expect(out).not.toBe(ops);
    // Inputs unmutated.
    expect(ops).toEqual(snapshot);
    // Every input id survives — never dropped.
    expect(out.map((o) => o.id)).toEqual(['a', 'b', 'c']);
  });
});
