import { describe, it, expect, vi } from 'vitest';
import { resolveAnchor } from '@shared/doc-model/anchor.js';
import type { TextAnchor } from '@shared/doc-model/anchor.js';
import type { ParaNode } from '@shared/doc-model/types.js';

// ---------------------------------------------------------------------------
// resolveAnchor re-resolves anchors in the *caller's* context via a
// document-wide body.search(), picking the hit by global occurrence rank.
// The mocks below model that: a body whose search() returns one Range per
// occurrence of the needle across the document.
// ---------------------------------------------------------------------------

function makeRange(label: string): Word.Range {
  return { _label: label, select: vi.fn(), scrollIntoView: vi.fn() } as unknown as Word.Range;
}

function makeParaNode(id: string, text: string, container: ParaNode['container'] = 'body'): ParaNode {
  return {
    id,
    index: 0,
    text,
    styleName: 'Normal',
    container,
    getRange: () => makeRange(`para-${id}`),
  };
}

/**
 * Build a mock RequestContext whose body.search(needle) returns `hits[needle]`
 * (one Range per document-order occurrence). Tracks the search call.
 */
function makeCtx(hits: Record<string, Word.Range[]>): {
  ctx: Word.RequestContext;
  bodyRange: Word.Range;
  search: ReturnType<typeof vi.fn>;
} {
  const bodyRange = makeRange('body');
  const search = vi.fn((needle: string) => ({
    items: hits[needle] ?? [],
    load: vi.fn(),
  }));
  const ctx = {
    document: { body: { getRange: () => bodyRange, search } },
  } as unknown as Word.RequestContext;
  return { ctx, bodyRange, search };
}

describe('resolveAnchor', () => {
  it('throws if paragraphId is not found', () => {
    const paras: ReadonlyArray<ParaNode> = [makeParaNode('p0', 'hello world')];
    const anchor: TextAnchor = { paragraphId: 'missing', startOffset: 0, endOffset: 5 };
    const { ctx } = makeCtx({});
    expect(() => resolveAnchor(ctx, paras, anchor)).toThrow("paragraph 'missing' not found");
  });

  it('searches for the exact offending substring with matchCase', () => {
    const para = makeParaNode('p0', 'Amount is S$10,000');
    const { ctx, search } = makeCtx({ 'S$10,000': [makeRange('hit')] });
    resolveAnchor(ctx, [para], { paragraphId: 'p0', startOffset: 10, endOffset: 18 });
    expect(search).toHaveBeenCalledWith('S$10,000', { matchCase: true, matchWholeWord: false });
  });

  it('selects the first occurrence when rank=0', () => {
    const para = makeParaNode('p0', 'foo bar foo baz');
    const hit0 = makeRange('h0');
    const hit1 = makeRange('h1');
    const { ctx } = makeCtx({ foo: [hit0, hit1] });
    const resolved = resolveAnchor(ctx, [para], { paragraphId: 'p0', startOffset: 0, endOffset: 3 });
    (resolved as unknown as { select(): void }).select();
    expect(hit0.select).toHaveBeenCalled();
    expect(hit1.select).not.toHaveBeenCalled();
  });

  it('selects the second occurrence when an earlier match exists in the same paragraph', () => {
    const para = makeParaNode('p0', 'foo bar foo baz');
    const hit0 = makeRange('h0');
    const hit1 = makeRange('h1');
    const { ctx } = makeCtx({ foo: [hit0, hit1] });
    // Anchor on the SECOND "foo" (offset 8) → rank 1.
    const resolved = resolveAnchor(ctx, [para], { paragraphId: 'p0', startOffset: 8, endOffset: 11 });
    (resolved as unknown as { select(): void }).select();
    expect(hit0.select).not.toHaveBeenCalled();
    expect(hit1.select).toHaveBeenCalled();
  });

  it('counts occurrences in earlier paragraphs for the global rank', () => {
    // p0 has two "foo"; the anchor targets the single "foo" in p1 → rank 2.
    const p0 = makeParaNode('p0', 'foo and foo again');
    const p1 = makeParaNode('p1', 'foo here');
    const hits = [makeRange('h0'), makeRange('h1'), makeRange('h2')];
    const { ctx } = makeCtx({ foo: hits });
    const resolved = resolveAnchor(ctx, [p0, p1], { paragraphId: 'p1', startOffset: 0, endOffset: 3 });
    (resolved as unknown as { select(): void }).select();
    expect(hits[2]!.select).toHaveBeenCalled();
    expect(hits[0]!.select).not.toHaveBeenCalled();
    expect(hits[1]!.select).not.toHaveBeenCalled();
  });

  it('does not count footnote paragraphs (body.search excludes footnotes)', () => {
    // A footnote before the body para contains the needle, but must not shift rank.
    const fn = makeParaNode('fn0', 'foo in footnote', 'footnote');
    const body = makeParaNode('p0', 'foo in body');
    const hit0 = makeRange('h0');
    const { ctx } = makeCtx({ foo: [hit0] });
    const resolved = resolveAnchor(ctx, [fn, body], { paragraphId: 'p0', startOffset: 0, endOffset: 3 });
    (resolved as unknown as { select(): void }).select();
    expect(hit0.select).toHaveBeenCalled();
  });

  it('is a no-op when the text is no longer present (document edited)', () => {
    const para = makeParaNode('p0', 'some text');
    const { ctx } = makeCtx({}); // search returns no hits
    const resolved = resolveAnchor(ctx, [para], { paragraphId: 'p0', startOffset: 0, endOffset: 4 });
    // Must not throw.
    expect(() => (resolved as unknown as { select(): void }).select()).not.toThrow();
  });

  it('navigates to the paragraph for a zero-length anchor', () => {
    const para = makeParaNode('p0', 'whole paragraph text');
    const paraHit = makeRange('paraHit');
    const { ctx, search } = makeCtx({ 'whole paragraph text': [paraHit] });
    const resolved = resolveAnchor(ctx, [para], { paragraphId: 'p0', startOffset: 5, endOffset: 5 });
    expect(search).toHaveBeenCalledWith('whole paragraph text', { matchCase: true, matchWholeWord: false });
    (resolved as unknown as { select(): void }).select();
    expect(paraHit.select).toHaveBeenCalled();
  });
});
