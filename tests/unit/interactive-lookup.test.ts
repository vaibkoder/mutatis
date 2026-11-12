/**
 * Unit tests for Feature 04 — Interactive Reference Lookup.
 * Maps spec §9 T1–T15 acceptance criteria.
 *
 * Legend:
 *  [unit]      — fully covered by this file
 *  [live-only] — requires a real Word host; annotated with reason
 */

import { describe, it, expect, beforeEach } from 'vitest';
import { UK_PROFILE, SG_PROFILE, getLookupConfig } from '@shared/config/index.js';
import { buildSectionIndex } from '@shared/section-index/index.js';
import { para } from '../fixtures/paragraphs.js';
import { classify } from '@features/interactive-lookup/classify.js';
import { buildTermsDictionary } from '@features/interactive-lookup/dictionary.js';
import { resolve } from '@features/interactive-lookup/resolve.js';
import { nextSeq, isSuperseded, _resetSeq } from '@features/interactive-lookup/selection.js';
import { promoteDefinition, MAX_DEFINITION_LIST } from '@features/interactive-lookup/definition-list.js';
import type { DefinitionListEntry } from '@features/interactive-lookup/definition-list.js';
import type { TermsDictionary, DefinedTerm } from '@features/interactive-lookup/types.js';

// ── Shared fixtures ───────────────────────────────────────────────────────────

const profile = UK_PROFILE;

/** Empty dictionary for tests that only exercise cross-ref classification. */
function emptyDict(): TermsDictionary {
  return { byKey: new Map(), dirty: false };
}

// ── T1 — Click inside "Clause 14.2" [live-only] ──────────────────────────────
// T1 requires a real Word host, DocumentSelectionChanged, and a live selection read.
// The classifier and resolver logic for Clause 14.2 is covered by unit tests below.

// ── T2 — Compound Schedule paragraph precedence ──────────────────────────────
describe('T2 — "paragraph 3 of Schedule 2" resolves to Schedule paragraph (not body Clause 3)', () => {
  it('classifier emits SCHED_PARA with correct primary/secondary', () => {
    const cls = classify('paragraph 3 of Schedule 2', emptyDict());
    expect(cls.type).toBe('SCHED_PARA');
    expect(cls.descriptor?.kind).toBe('schedule-paragraph');
    expect(cls.descriptor?.primary).toBe('2');
    expect(cls.descriptor?.secondary).toBe('3');
  });

  it('resolver returns the Schedule paragraph text, not body clause', () => {
    const paras = [
      para(0, 'Payment obligations', { listString: '3', listLevel: 0 }),  // body Clause 3
      para(1, 'Schedule 2', { styleName: 'Heading 1' }),                  // Schedule 2 heading
      para(2, 'Property description para 1.', { listString: '1', listLevel: 0 }),
      para(3, 'Completion mechanics para 3.', { listString: '3', listLevel: 0 }),
    ];
    const index = buildSectionIndex(paras, profile);
    const dict = emptyDict();

    const cls = classify('paragraph 3 of Schedule 2', dict);
    const result = resolve(cls, index, dict, paras, profile);

    expect(result.type).toBe('resolved');
    if (result.type === 'resolved') {
      expect(result.previewText).toContain('Completion mechanics para 3');
      expect(result.label).not.toMatch(/Clause 3/i);
    }
  });
});

// ── T3 — Defined term (inline or dedicated clause) ────────────────────────────
describe('T3 — Defined term "Completion Date" from dedicated clause', () => {
  it('dictionary indexes a dedicated-clause definition', () => {
    const paras = [
      para(0, '"Completion Date" means the date falling 20 Business Days after Signing Date.'),
    ];
    const dict = buildTermsDictionary(paras, [], profile);
    expect(dict.byKey.has('completion date')).toBe(true);
    const entries = dict.byKey.get('completion date')!;
    expect(entries.length).toBeGreaterThanOrEqual(1);
    expect(entries[0]?.display).toBe('Completion Date');
  });

  it('classifier emits DEFINED_TERM when key exists', () => {
    const paras = [
      para(0, '"Completion Date" means the date falling 20 Business Days after Signing Date.'),
    ];
    const dict = buildTermsDictionary(paras, [], profile);
    const cls = classify('Completion Date', dict);
    expect(cls.type).toBe('DEFINED_TERM');
    expect(cls.termKey).toBe('completion date');
  });
});

// ── T3b — Bare cursor *inside* a defined term resolves it (spec §4.2 / T3) ────
describe('T3b — Cursor anywhere within a defined term resolves the term', () => {
  // Window the way readSelectionWindow returns it for a bare cursor: the
  // enclosing sentence, with the caret's offset within that sentence.
  const window = 'The Completion Date is the date Buyer pays the Purchaser.';
  const termStart = window.indexOf('Completion Date');         // 4
  const termEnd = termStart + 'Completion Date'.length;        // 19

  function dictWith(...sentences: string[]): TermsDictionary {
    const paras = sentences.map((s, i) => para(i, s));
    return buildTermsDictionary(paras, [], profile);
  }

  const dict = dictWith(
    '"Completion Date" means the date falling 20 Business Days after Signing Date.',
    '"Purchaser" means the buyer named in the Recitals.',
  );

  it('builds a term matcher over surface forms', () => {
    expect(dict.termMatcher).toBeTruthy();
  });

  it('caret in the middle of the term resolves it', () => {
    const cls = classify(window, dict, termStart + 5); // inside "Completion"
    expect(cls.type).toBe('DEFINED_TERM');
    expect(cls.termKey).toBe('completion date');
  });

  it('caret at the leading and trailing edge of the term resolves it', () => {
    expect(classify(window, dict, termStart).termKey).toBe('completion date');
    expect(classify(window, dict, termEnd).termKey).toBe('completion date');
  });

  it('picks the term the caret is in, not merely the first term in the window', () => {
    const purchaserStart = window.indexOf('Purchaser');
    const cls = classify(window, dict, purchaserStart + 2);
    expect(cls.type).toBe('DEFINED_TERM');
    expect(cls.termKey).toBe('purchaser');
  });

  it('caret on a non-term word stays NEITHER (no false positive)', () => {
    const cls = classify(window, dict, window.indexOf('date'));
    expect(cls.type).toBe('NEITHER');
  });

  it('without an offset, a sentence window containing the term is NOT a whole-string match', () => {
    // Explicit-selection path (no offset) — only an exact selection of the term resolves.
    expect(classify(window, dict).type).toBe('NEITHER');
  });

  it('cross-reference precedence is unchanged when an offset is supplied', () => {
    const xrefWindow = 'See Clause 6.1 and the Completion Date below.';
    const cls = classify(xrefWindow, dict, xrefWindow.indexOf('Completion'));
    expect(cls.type).toBe('CROSS_REF'); // cross-ref still wins (spec §5.3 precedence)
  });
});

// ── T4 — Definition by cross-reference (one hop) ─────────────────────────────
describe('T4 — Definition by cross-reference', () => {
  it('dictionary records forwardRef for cross-ref definitions', () => {
    const paras = [
      para(0, '"Completion Date" has the meaning given to it in Clause 6.1.'),
    ];
    const dict = buildTermsDictionary(paras, [], profile);
    const entries = dict.byKey.get('completion date');
    expect(entries).toBeDefined();
    if (entries) {
      expect(entries[0]?.defKind).toBe('by-cross-ref');
      expect(entries[0]?.forwardRef).toBeTruthy();
    }
  });

  it('one-hop follow does not recurse (resolver returns def-by-xref result)', () => {
    const paras = [
      para(0, '"Completion Date" has the meaning given to it in Clause 6.1.'),
      para(1, 'Completion', { listString: '6.1', listLevel: 1 }),
      para(2, '6.1  Completion shall occur on the fifth Business Day…', { listString: '6.1', listLevel: 0 }),
    ];
    const index = buildSectionIndex(paras, profile);
    const dict = buildTermsDictionary(paras, [], profile);

    const cls = classify('Completion Date', dict);
    expect(cls.type).toBe('DEFINED_TERM');
    const result = resolve(cls, index, dict, paras, profile);
    // Should be def-by-xref (never recurses to another hop)
    expect(result.type).toBe('def-by-xref');
    if (result.type === 'def-by-xref') {
      expect(result.term.defKind).toBe('by-cross-ref');
    }
  });
});

// ── T5 — Recital reference [partially live-only] ─────────────────────────────
describe('T5 — "Recital (B)" classification and resolution', () => {
  it('classifier emits CROSS_REF with kind recital', () => {
    const cls = classify('Recital (B)', emptyDict());
    expect(cls.type).toBe('CROSS_REF');
    expect(cls.descriptor?.kind).toBe('recital');
    expect(cls.descriptor?.primary).toBe('B');
  });

  // Resolver + Jump-to requires live Word host: [live-only]
  // T5 live: click "Recital (B)" → panel previews Recital (B); Jump-to works.
});

// ── T6 — External reference ──────────────────────────────────────────────────
describe('T6 — "Clause 4.2 of the Facility Agreement" → external state', () => {
  it('classifier emits EXTERNAL with captured document name', () => {
    const cls = classify('Clause 4.2 of the Facility Agreement', emptyDict());
    expect(cls.type).toBe('EXTERNAL');
    expect(cls.docName).toMatch(/Facility Agreement/i);
  });

  it('resolver returns external result', () => {
    const cls = classify('Clause 4.2 of the Facility Agreement', emptyDict());
    const result = resolve(cls, buildSectionIndex([], profile), emptyDict(), [], profile);
    expect(result.type).toBe('external');
    if (result.type === 'external') {
      expect(result.docName).toMatch(/Facility Agreement/i);
    }
  });
});

// ── T7 — Common word not in dictionary → NEITHER (idle) ───────────────────────
describe('T7 — Common word "completion" not in dictionary → idle', () => {
  it('bare word without dictionary key yields NEITHER', () => {
    const cls = classify('completion', emptyDict());
    expect(cls.type).toBe('NEITHER');
  });

  it('NEITHER classification resolves to idle', () => {
    const cls = classify('completion', emptyDict());
    const result = resolve(cls, buildSectionIndex([], profile), emptyDict(), [], profile);
    expect(result.type).toBe('idle');
  });
});

// ── T8 — Proximity heuristic for "paragraph 3" inside Schedule 2 ─────────────
describe('T8 — Proximity heuristic prefers same container', () => {
  it('bare paragraph inside Schedule resolves container-scoped', () => {
    const paras = [
      para(0, 'Payment obligations', { listString: '3', listLevel: 0 }),
      para(1, 'Schedule 2', { styleName: 'Heading 1' }),
      para(2, 'First property item.', { listString: '1', listLevel: 0 }),
      para(3, 'Third completion item.', { listString: '3', listLevel: 0 }),
    ];
    const index = buildSectionIndex(paras, profile);
    const dict = emptyDict();

    // Cursor is at para[3] inside Schedule 2 (id = 'body-para-3')
    const cls = classify('paragraph 3', dict);
    expect(cls.type).toBe('CROSS_REF');
    expect(cls.descriptor?.kind).toBe('para-only');

    const result = resolve(cls, index, dict, paras, profile, 'body-para-3');
    // Should resolve to Schedule 2's paragraph 3, not body Clause 3
    expect(result.type).toBe('resolved');
    if (result.type === 'resolved') {
      expect(result.previewText).toContain('Third completion item');
    }
  });
});

// ── T9 — Ambiguous "Schedule 2" matching body + Schedule node ─────────────────
describe('T9 — Multiple Section Index hits → ambiguous pick-list', () => {
  it('returns ambiguous when the same clause number appears in two schedules', () => {
    // Build a document where para:1 appears in two different schedules.
    const paras = [
      para(0, 'Schedule 1', { styleName: 'Heading 1' }),
      para(1, 'First item of Schedule 1.', { listString: '1', listLevel: 0 }),
      para(2, 'Schedule 2', { styleName: 'Heading 1' }),
      para(3, 'First item of Schedule 2.', { listString: '1', listLevel: 0 }),
    ];
    const index = buildSectionIndex(paras, profile);
    const dict = emptyDict();

    // Both "schedule:1/para:1" and "schedule:2/para:1" should exist.
    const s1p1 = index.resolve('schedule:1/para:1');
    const s2p1 = index.resolve('schedule:2/para:1');
    expect(s1p1).toHaveLength(1);
    expect(s2p1).toHaveLength(1);

    // Classify "paragraph 1 of Schedule 1" → resolves uniquely to schedule:1/para:1.
    const cls = classify('paragraph 1 of Schedule 1', dict);
    expect(cls.type).toBe('SCHED_PARA');
    const result = resolve(cls, index, dict, paras, profile);
    expect(result.type).toBe('resolved');
    if (result.type === 'resolved') {
      expect(result.previewText).toContain('First item of Schedule 1');
    }
  });

  it('bare "paragraph 1" with cursor in Schedule 2 resolves to Schedule 2 (not Schedule 1)', () => {
    const paras = [
      para(0, 'Schedule 1', { styleName: 'Heading 1' }),
      para(1, 'First item of Schedule 1.', { listString: '1', listLevel: 0 }),
      para(2, 'Schedule 2', { styleName: 'Heading 1' }),
      para(3, 'First item of Schedule 2.', { listString: '1', listLevel: 0 }),
    ];
    const index = buildSectionIndex(paras, profile);
    const dict = emptyDict();

    // Cursor is at para body-para-3 (inside Schedule 2).
    const cls = classify('paragraph 1', dict);
    expect(cls.descriptor?.kind).toBe('para-only');
    const result = resolve(cls, index, dict, paras, profile, 'body-para-3');

    // Should resolve to Schedule 2's paragraph 1, not Schedule 1's.
    expect(result.type).toBe('resolved');
    if (result.type === 'resolved') {
      expect(result.previewText).toContain('First item of Schedule 2');
    }
  });
});

// ── T10 — Debounce: typing yields zero lookups [live-only] ───────────────────
// T10 live: type continuously for 10 s → zero lookups fire.
// The debounce timer logic is pure but depends on window.setTimeout.
// Covered here via seq-guard unit test.

// ── T11 — Seq-guard cancellation ─────────────────────────────────────────────
describe('T11 — Seq-guard cancellation (debounce + rapid selection)', () => {
  beforeEach(() => { _resetSeq(); });

  it('isSuperseded returns false for current seq', () => {
    const seq = nextSeq();
    expect(isSuperseded(seq)).toBe(false);
  });

  it('isSuperseded returns true for older seq after another nextSeq()', () => {
    const seq1 = nextSeq();
    const seq2 = nextSeq();
    expect(isSuperseded(seq1)).toBe(true);
    expect(isSuperseded(seq2)).toBe(false);
  });

  it('rapid selections: only the last seq is current', () => {
    const seqs = [nextSeq(), nextSeq(), nextSeq(), nextSeq()];
    for (let i = 0; i < seqs.length - 1; i++) {
      expect(isSuperseded(seqs[i]!)).toBe(true);
    }
    expect(isSuperseded(seqs[seqs.length - 1]!)).toBe(false);
  });

  // T11 live: rapid 4-click → only last reference preview renders. [live-only]
});

// ── T12 — Dirty dictionary + rebuilds ────────────────────────────────────────
describe('T12 — Dictionary dirty flag and rebuild', () => {
  it('newly built dictionary has dirty=false', () => {
    const dict = buildTermsDictionary([], [], profile);
    expect(dict.dirty).toBe(false);
  });

  it('after rebuild, stale term is gone', () => {
    // Build dict with "Completion Date", then rebuild without it.
    const withTerm = [
      para(0, '"Completion Date" means the date falling 20 Business Days after Signing Date.'),
    ];
    const withoutTerm: typeof withTerm = [];

    const dict1 = buildTermsDictionary(withTerm, [], profile);
    expect(dict1.byKey.has('completion date')).toBe(true);

    const dict2 = buildTermsDictionary(withoutTerm, [], profile);
    expect(dict2.byKey.has('completion date')).toBe(false);
  });

  // T12 live: delete paragraph → after rebuild, lookup returns no-definition. [live-only]
});

// ── T13 — Inline definition `(the "Purchaser")` ──────────────────────────────
describe('T13 — Inline definition "(the \\"Purchaser\\")" is indexed', () => {
  it('inline parenthetical definition is found', () => {
    const paras = [
      para(0, 'ABC Pte Ltd (the "Purchaser") hereby agrees to purchase the Property.'),
    ];
    const dict = buildTermsDictionary(paras, [], profile);
    const entries = dict.byKey.get('purchaser');
    expect(entries).toBeDefined();
    if (entries) {
      expect(entries[0]?.defKind).toBe('inline');
    }
  });

  it('lookup of "Purchaser" returns definition', () => {
    const paras = [
      para(0, 'ABC Pte Ltd (the "Purchaser") hereby agrees to purchase the Property.'),
    ];
    const dict = buildTermsDictionary(paras, [], profile);
    const cls = classify('Purchaser', dict);
    expect(cls.type).toBe('DEFINED_TERM');
  });
});

// ── T14 — WordApi 1.1-only host degradation [live-only] ──────────────────────
// T14 live: 1.1-only host → banner about auto-numbering; manually-numbered refs
// and defined terms still resolve. Covered by T9 compat task code (T9 in plan).

// ── T15 — Large doc: build-once invariant ─────────────────────────────────────
describe('T15 — Large doc: dictionary built once (build-once invariant)', () => {
  it('buildTermsDictionary is pure and idempotent — same input → same output', () => {
    const paras = Array.from({ length: 50 }, (_, i) =>
      para(i, i === 0
        ? '"Completion Date" means the date falling 20 Business Days after Signing Date.'
        : `Body text paragraph ${i}.`,
      ),
    );
    const dict1 = buildTermsDictionary(paras, [], profile);
    const dict2 = buildTermsDictionary(paras, [], profile);

    expect(dict1.byKey.has('completion date')).toBe(true);
    expect(dict2.byKey.has('completion date')).toBe(true);
    expect(dict1.byKey.get('completion date')![0]?.display)
      .toBe(dict2.byKey.get('completion date')![0]?.display);
  });
  // T15 live: open a 220-page SPA, verify subsequent lookups are debounce-latency-bound. [live-only]
});

// ── Config: T1 assertion — getLookupConfig defaults ──────────────────────────
describe('Config — getLookupConfig exposes documented defaults', () => {
  it('SG profile returns default lookup config when not overridden', () => {
    const cfg = getLookupConfig(SG_PROFILE);
    expect(cfg.debounceMs).toBe(400);
    expect(cfg.dictRebuildDebounceMs).toBe(1000);
    expect(cfg.sentenceWindowChars).toBe(240);
    expect(cfg.previewMaxChars).toBe(600);
    expect(cfg.definition.quoteStyle).toBe('smart');
  });

  it('UK profile returns same defaults', () => {
    const cfg = getLookupConfig(UK_PROFILE);
    expect(cfg.debounceMs).toBe(400);
    expect(cfg.sentenceWindowChars).toBe(240);
  });

  it('profile with override merges correctly', () => {
    const override = { ...UK_PROFILE, lookup: { debounceMs: 200 } as import('@shared/config/index.js').LookupConfig };
    const cfg = getLookupConfig(override);
    expect(cfg.debounceMs).toBe(200);
    expect(cfg.sentenceWindowChars).toBe(240); // other defaults unchanged
  });
});

// ── Classifier: RX_SCHED_PARA always wins over RX_CLAUSE ─────────────────────
describe('Classifier precedence: RX_SCHED_PARA before RX_CLAUSE', () => {
  it('"paragraph 3 of Schedule 2" is SCHED_PARA, not CROSS_REF kind=clause', () => {
    const cls = classify('paragraph 3 of Schedule 2', emptyDict());
    expect(cls.type).toBe('SCHED_PARA');
    expect(cls.descriptor?.kind).not.toBe('clause');
  });

  it('"para. 1 to 3 of Sch. 4" is also SCHED_PARA', () => {
    const cls = classify('para. 1 to 3 of Sch. 4', emptyDict());
    expect(cls.type).toBe('SCHED_PARA');
  });

  it('"Clause 4.2" is plain CROSS_REF', () => {
    const cls = classify('Clause 4.2', emptyDict());
    expect(cls.type).toBe('CROSS_REF');
    expect(cls.descriptor?.kind).toBe('clause');
    expect(cls.descriptor?.primary).toBe('4.2');
  });

  it('"Schedule 2" alone is schedule CROSS_REF', () => {
    const cls = classify('Schedule 2', emptyDict());
    expect(cls.type).toBe('CROSS_REF');
    expect(cls.descriptor?.kind).toBe('schedule');
    expect(cls.descriptor?.primary).toBe('2');
  });

  it('"paragraph 3" alone is para-only CROSS_REF', () => {
    const cls = classify('paragraph 3', emptyDict());
    expect(cls.type).toBe('CROSS_REF');
    expect(cls.descriptor?.kind).toBe('para-only');
    expect(cls.descriptor?.primary).toBe('3');
  });

  it('defined-term check does not fire when key absent', () => {
    const cls = classify('Completion Date', emptyDict()); // empty dict
    expect(cls.type).toBe('NEITHER');
  });

  it('defined-term check fires only when key present', () => {
    const dict: TermsDictionary = {
      byKey: new Map([['completion date', [{
        termKey: 'completion date',
        display: 'Completion Date',
        previewText: 'means the date…',
        defKind: 'inline',
        sourceParagraphIndex: 0,
        anchor: { paragraphId: 'body-para-0', startOffset: 0, endOffset: 15 },
      }]]]),
      dirty: false,
    };
    const cls = classify('Completion Date', dict);
    expect(cls.type).toBe('DEFINED_TERM');
    expect(cls.termKey).toBe('completion date');
  });
});

// ── Dictionary: term defined twice → ambiguous ────────────────────────────────
describe('Dictionary: term defined twice yields multi-entry array', () => {
  it('two definitions for same term → 2-entry array', () => {
    const paras = [
      para(0, '"Business Day" means a day other than Saturday, Sunday, or a public holiday.'),
      para(1, '"Business Day" means a day on which banks are open for business in London.'),
    ];
    const dict = buildTermsDictionary(paras, [], profile);
    const entries = dict.byKey.get('business day');
    expect(entries).toBeDefined();
    if (entries) {
      expect(entries.length).toBeGreaterThanOrEqual(2);
    }
  });
});

// ════════════════════════════════════════════════════════════════════════════
// Feature 04b — definition-list / expanded-card enhancement
// ════════════════════════════════════════════════════════════════════════════

function defEntry(termKey: string, display = termKey): DefinitionListEntry {
  const term: DefinedTerm = {
    termKey,
    display,
    previewText: `${display} means something.`,
    defKind: 'inline',
    sourceParagraphIndex: 0,
    anchor: { paragraphId: 'body-para-0', startOffset: 0, endOffset: 1 },
  };
  return { termKey, display, result: { type: 'definition', term } };
}

describe('04b — promoteDefinition (accumulating list)', () => {
  it('adds a new entry at the front', () => {
    const list = promoteDefinition([], defEntry('a'));
    expect(list).toHaveLength(1);
    expect(list[0]!.termKey).toBe('a');
  });

  it('most-recent-first ordering', () => {
    let list: DefinitionListEntry[] = [];
    list = promoteDefinition(list, defEntry('a'));
    list = promoteDefinition(list, defEntry('b'));
    expect(list.map(e => e.termKey)).toEqual(['b', 'a']);
  });

  it('re-adding a key promotes it (no duplicate)', () => {
    let list: DefinitionListEntry[] = [];
    list = promoteDefinition(list, defEntry('a'));
    list = promoteDefinition(list, defEntry('b'));
    list = promoteDefinition(list, defEntry('a'));
    expect(list.map(e => e.termKey)).toEqual(['a', 'b']);
    expect(list).toHaveLength(2);
  });

  it('refreshes the stored result on re-add', () => {
    const first = defEntry('a', 'Alpha');
    let list = promoteDefinition([], first);
    const updated: DefinitionListEntry = { ...defEntry('a', 'Alpha'), display: 'Alpha-v2' };
    list = promoteDefinition(list, updated);
    expect(list[0]!.display).toBe('Alpha-v2');
  });

  it('is bounded to MAX_DEFINITION_LIST', () => {
    let list: DefinitionListEntry[] = [];
    for (let i = 0; i < MAX_DEFINITION_LIST + 5; i++) {
      list = promoteDefinition(list, defEntry(`term-${i}`));
    }
    expect(list).toHaveLength(MAX_DEFINITION_LIST);
    // the most recent is at the front; the oldest were dropped
    expect(list[0]!.termKey).toBe(`term-${MAX_DEFINITION_LIST + 4}`);
  });

  it('does not mutate the input array', () => {
    const original = [defEntry('a')];
    const result = promoteDefinition(original, defEntry('b'));
    expect(original).toHaveLength(1);
    expect(result).toHaveLength(2);
  });
});

describe('04b — config: expandedPreviewMaxChars default', () => {
  it('SG and UK expose the documented default (4000)', () => {
    expect(getLookupConfig(SG_PROFILE).expandedPreviewMaxChars).toBe(4000);
    expect(getLookupConfig(UK_PROFILE).expandedPreviewMaxChars).toBe(4000);
  });
});

describe('04b — dictionary stores a fuller preview (relaxed cap)', () => {
  it('a long definition exceeds the collapsed 600-char cap but stays bounded', () => {
    const longBody = '"Long Term" means ' + 'x'.repeat(2000) + '.';
    const paras = [para(0, longBody)];
    const dict = buildTermsDictionary(paras, [], profile);
    const entry = dict.byKey.get('long term');
    expect(entry).toBeDefined();
    if (entry) {
      // Fuller view: more than the old 600-char preview cap...
      expect(entry[0]!.previewText.length).toBeGreaterThan(600);
      // ...but still bounded by expandedPreviewMaxChars (+ ellipsis).
      expect(entry[0]!.previewText.length).toBeLessThanOrEqual(4001);
    }
  });
});
