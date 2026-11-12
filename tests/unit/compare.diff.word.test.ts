import { describe, it, expect } from 'vitest';

import { getCompareConfig, UK_PROFILE } from '@shared/config/index.js';
import { wordDiff, tokenize, tokenizeTagged, atomicKindOf } from '@features/compare/diff/index.js';

const CONFIG = getCompareConfig(UK_PROFILE);

describe('tokenize — atomic tokens (§5.3)', () => {
  it('never splits an amount mid-token', () => {
    const toks = tokenize('pay £40,000 now', CONFIG);
    expect(toks).toContain('£40,000');
    // The digits must not be separate tokens.
    expect(toks).not.toContain('40');
    expect(toks).not.toContain('000');
  });

  it('keeps a quoted defined term whole', () => {
    const toks = tokenize('the "Permitted Encumbrance" applies', CONFIG);
    expect(toks).toContain('"Permitted Encumbrance"');
  });

  it('keeps a clause reference VALUE whole, the noun an ordinary word', () => {
    const toks = tokenizeTagged('see Clause 4.2(a) below', CONFIG);
    // The value is indivisible (P5) …
    expect(toks).toContainEqual({ text: '4.2(a)', atomic: 'clauseRef' });
    // … but the noun is not part of it: renumbering marks the number alone (P1).
    expect(toks).toContainEqual({ text: 'Clause' });
    expect(toks.map((t) => t.text).join('')).toBe('see Clause 4.2(a) below');
  });

  it('recognises a lower-case reference noun', () => {
    // "clause 4.2" is the standard E&W/SG form; the match is case-insensitive.
    expect(tokenizeTagged('see clause 4.2 below', CONFIG)).toContainEqual({
      text: '4.2',
      atomic: 'clauseRef',
    });
  });

  it('does not eat a word that merely ends in a reference noun', () => {
    // Case-insensitivity must not turn "subparagraph 3" into "sub" + "paragraph 3".
    expect(tokenize('see subparagraph 3 below', CONFIG)).toContain('subparagraph');
    expect(tokenize('see subparagraph 3 below', CONFIG)).not.toContain('sub');
    // Same guard, the other reference noun: "Subclause" must not lose its "Sub".
    const sub = tokenizeTagged('under Subclause 4.2 hereof', CONFIG);
    expect(sub.map((t) => t.text)).toContain('Subclause');
    expect(sub.map((t) => t.text)).not.toContain('clause');
    // "Subclause" is not one of this profile's `referenceTokens`, so its number
    // is NOT a tagged cross-reference — nothing downstream may stamp it as one.
    expect(sub.find((t) => t.text === '4.2')?.atomic).toBeUndefined();
  });

  it('tags both values of "paragraph N of Schedule M"', () => {
    const toks = tokenizeTagged('under paragraph 2 of Schedule 4 hereof', CONFIG);
    expect(toks).toContainEqual({ text: '2', atomic: 'clauseRef' });
    expect(toks).toContainEqual({ text: '4', atomic: 'clauseRef' });
    expect(toks).toContainEqual({ text: 'of' });
    expect(toks.map((t) => t.text).join('')).toBe('under paragraph 2 of Schedule 4 hereof');
  });

  it('keeps a date whole', () => {
    const toks = tokenize('by 1 January 2026 at noon', CONFIG);
    expect(toks).toContain('1 January 2026');
  });

  it('keeps a duration whole', () => {
    expect(tokenize('within 30 days', CONFIG)).toContain('30 days');
    expect(tokenize('within three (3) Business Days', CONFIG)).toContain(
      'three (3) Business Days',
    );
  });
});

describe('atomicKindOf', () => {
  it('classifies atomic kinds', () => {
    expect(atomicKindOf('£40,000', CONFIG)).toBe('amount');
    expect(atomicKindOf('S$1,250,000', CONFIG)).toBe('amount');
    expect(atomicKindOf('US$2.5m', CONFIG)).toBe('amount');
    expect(atomicKindOf('"Permitted Encumbrance"', CONFIG)).toBe('definedTerm');
    expect(atomicKindOf('Clause 4.2(a)', CONFIG)).toBe('clauseRef');
    expect(atomicKindOf('Schedule 3', CONFIG)).toBe('clauseRef');
    expect(atomicKindOf('1 January 2026', CONFIG)).toBe('date');
    expect(atomicKindOf('7.5%', CONFIG)).toBe('percent');
    expect(atomicKindOf('30 days', CONFIG)).toBe('duration');
  });

  it('returns undefined for ordinary words', () => {
    expect(atomicKindOf('Seller', CONFIG)).toBeUndefined();
    expect(atomicKindOf('shall', CONFIG)).toBeUndefined();
  });
});

describe('wordDiff — T5 (obligation swap + amount atomicity)', () => {
  it('reports shall→may as a single non-atomic ins + del', () => {
    const diffs = wordDiff('The Seller shall pay', 'The Seller may pay', CONFIG);
    const ins = diffs.filter((d) => d.op === 'ins');
    const del = diffs.filter((d) => d.op === 'del');

    expect(ins).toHaveLength(1);
    expect(del).toHaveLength(1);
    expect(ins[0]!.text).toBe('may');
    expect(ins[0]!.atomic).toBeUndefined();
    expect(del[0]!.text).toBe('shall');
    expect(del[0]!.atomic).toBeUndefined();
  });

  it('reports £40,000 → £50,000 as one atomic amount change, not digit-level', () => {
    const diffs = wordDiff('pay £40,000 now', 'pay £50,000 now', CONFIG);
    const ins = diffs.filter((d) => d.op === 'ins');
    const del = diffs.filter((d) => d.op === 'del');

    expect(ins).toHaveLength(1);
    expect(del).toHaveLength(1);
    expect(ins[0]!.text).toBe('£50,000');
    expect(ins[0]!.atomic).toBe('amount');
    expect(del[0]!.text).toBe('£40,000');
    expect(del[0]!.atomic).toBe('amount');
    // No bare-digit tokens leaked into the diff.
    expect(diffs.some((d) => d.text === '4' || d.text === '5')).toBe(false);
  });

  it('coalesces consecutive equal non-atomic tokens', () => {
    const diffs = wordDiff('The Seller shall pay', 'The Seller may pay', CONFIG);
    const eqs = diffs.filter((d) => d.op === 'eq');
    // "The Seller " before and " pay" after should each be single coalesced eq runs.
    expect(eqs.length).toBeLessThanOrEqual(2);
    expect(eqs.some((d) => d.text.includes('The Seller'))).toBe(true);
  });
});

/**
 * The possessive apostrophe is not a quote delimiter (ADR-0029, correction of
 * 2026-09-11; plan § The two word-diff defects, defect B).
 *
 * `'[^']+'` unguarded matches everything between any two apostrophes — and in
 * English legal drafting that means everything between any two POSSESSIVES.
 * §5.2 folds `’ → '` before the tokenizer runs, so every Word document hits it.
 * The resulting span is `definedTerm`, therefore atomic, therefore P5 forbids
 * splitting it, therefore changing one word inside strikes and re-inserts the
 * whole span. The worst live instance restated 355 identical characters to say
 * that a clause had been appended.
 *
 * The 30-case corpus contained seven possessives and never two in one
 * paragraph, so it scored this 0 for its entire life.
 */
describe('tokenize — an apostrophe is a possessive before it is a quote', () => {
  it('keeps a singular possessive inside its word and spans nothing', () => {
    const toks = tokenizeTagged(
      "the Landlord and the Landlord's agent shall be given full notice of the surveys " +
        "commissioned by the Tenant and the Tenant's insurer.",
      CONFIG,
    );
    // Each possessive is ONE ordinary word — `Purchaser's` is one word to a
    // reader, so a rendering that swaps it whole borrows nothing.
    expect(toks.map((t) => t.text)).toContain("Landlord's");
    expect(toks.map((t) => t.text)).toContain("Tenant's");
    // And nothing anywhere is atomic. Without the guard, one 90-character
    // `definedTerm` runs from the first possessive to the second.
    expect(toks.filter((t) => t.atomic !== undefined)).toEqual([]);
  });

  it('keeps a PLURAL possessive whole, apostrophe and all', () => {
    // The plural is the harder half: `Vendors'` ends on the apostrophe, so the
    // closing-quote guard alone cannot see it — the span `' rights and the
    // Purchasers'` ends on an apostrophe followed by a space. Only the opening
    // guard rules it out, and only because the word-run break consults the
    // matcher at its real offset in the text rather than against a slice.
    const toks = tokenizeTagged(
      "The Purchasers' rights and the Vendors' duties survive Completion.",
      CONFIG,
    );
    expect(toks.map((t) => t.text)).toContain("Purchasers'");
    expect(toks.map((t) => t.text)).toContain("Vendors'");
    expect(toks.filter((t) => t.atomic !== undefined)).toEqual([]);
  });

  it('still recognises a genuinely single-quoted defined term', () => {
    // The guard narrows the rule; it does not withdraw it. An opening quote
    // that does not follow a word character, closing before one, is a quote.
    const toks = tokenizeTagged(
      "The Vendor shall deliver the 'Permitted Encumbrance' schedule.",
      CONFIG,
    );
    expect(toks.find((t) => t.atomic === 'definedTerm')?.text).toBe("'Permitted Encumbrance'");
  });

  it('still recognises a double-quoted defined term, which never had the defect', () => {
    const toks = tokenizeTagged(
      'The Vendor shall deliver the "Permitted Encumbrance" schedule.',
      CONFIG,
    );
    expect(toks.find((t) => t.atomic === 'definedTerm')?.text).toBe('"Permitted Encumbrance"');
  });

  it('declines rather than mis-marking when a quote closes on a possessive', () => {
    // `'the Vendor's duty'` — the apostrophe after `Vendor` is a possessive,
    // not the closing quote, so the honest span is the whole phrase and the
    // tokenizer cannot know that. It marks NOTHING atomic rather than marking
    // `'the Vendor'`, which would be a defined term the document does not have.
    const toks = tokenizeTagged("he said 'the Vendor's duty' was clear", CONFIG);
    expect(toks.filter((t) => t.atomic !== undefined)).toEqual([]);
    expect(toks.map((t) => t.text)).toContain("Vendor's");
  });

  it('marks one word, not the span between two possessives', () => {
    // The end-to-end consequence, and the reason this matters: the reader must
    // see the one word that changed. Before the guard, `full` → `reasonable`
    // struck the whole span between the two possessives and re-inserted it.
    const diffs = wordDiff(
      "the Landlord's agent shall be given full notice of the claims made to the Insurer " +
        "and the surveys commissioned by the Tenant's insurer.",
      "the Landlord's agent shall be given reasonable notice of the claims made to the Insurer " +
        "and the surveys commissioned by the Tenant's insurer.",
      CONFIG,
    );
    const del = diffs.filter((d) => d.op === 'del').map((d) => d.text).join('');
    const ins = diffs.filter((d) => d.op === 'ins').map((d) => d.text).join('');
    expect(del).toBe('full');
    expect(ins).toBe('reasonable');
  });
});
