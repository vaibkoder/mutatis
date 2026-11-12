import { describe, it, expect, vi } from 'vitest';
import { applyMetadata } from '@features/repair/apply/metadata.js';
import type { RepairFinding } from '@features/repair/types.js';

// ── Helpers ──────────────────────────────────────────────────────────────────

function makeBaseFinding(
  id: string,
  action: string,
  overrides: Partial<RepairFinding> = {}
): RepairFinding {
  return {
    id,
    category: 'metadata',
    severity: 'confidentiality',
    label: id,
    destructive: false,
    requiresConfirm: false,
    selectedByDefault: true,
    action,
    result: null,
    ...overrides,
  };
}

/**
 * Build a fake `Word.RequestContext` for the apply-metadata tests.
 *
 * `ctx.sync` is a vi.fn() resolving to undefined so tests can assert call
 * counts when needed.  The document graph is constructed from the opts
 * supplied; all `load()` calls are no-ops (data is already present on the
 * plain objects after the matching sync()).
 */
function makeFakeCtx(opts: {
  /** Initial writable built-in property values (default: all empty strings). */
  builtinProps?: Partial<{
    author: string;
    subject: string;
    title: string;
    keywords: string;
    company: string;
    manager: string;
    comments: string;
    category: string;
  }>;
  customPropItems?: Array<{ delete: ReturnType<typeof vi.fn> }>;
  xmlPartItems?: Array<{ delete: ReturnType<typeof vi.fn> }>;
  commentItems?: Array<{ delete: ReturnType<typeof vi.fn> }>;
  /** If provided, ctx.sync throws this on the first call that maps to the
   *  given action key. Used to simulate a GeneralException. */
  throwOn?: 'clearBuiltinProps' | 'deleteCustomProps' | 'deleteCustomXml' | 'deleteComments';
}): {
  ctx: Word.RequestContext;
  builtinPropsObj: Record<string, string>;
  syncMock: ReturnType<typeof vi.fn>;
} {
  const {
    builtinProps = {},
    customPropItems = [],
    xmlPartItems = [],
    commentItems = [],
    throwOn,
  } = opts;

  // Writable built-in props — we use a plain object so assignments are
  // captured; tests verify by reading the same object afterwards.
  const builtinPropsObj: Record<string, string> = {
    author: '',
    subject: '',
    title: '',
    keywords: '',
    company: '',
    manager: '',
    comments: '',
    category: '',
    ...builtinProps,
  };

  const customProperties = {
    items: customPropItems,
    load: vi.fn(),
  };

  const properties = new Proxy(
    { customProperties, load: vi.fn() },
    {
      get(target, prop) {
        if (prop in target) return (target as Record<string | symbol, unknown>)[prop];
        return builtinPropsObj[prop as string];
      },
      set(_target, prop, value) {
        builtinPropsObj[prop as string] = value as string;
        return true;
      },
    }
  );

  const customXmlParts = {
    items: xmlPartItems,
    load: vi.fn(),
  };

  const commentsCollection = {
    items: commentItems,
    load: vi.fn(),
  };

  const body = {
    getComments: () => commentsCollection,
  };

  const document = {
    properties,
    customXmlParts,
    body,
  };

  // Track which "logical action" the sync belongs to so we can throw
  // selectively.  We use a simple call-count approach aligned with how
  // applyMetadata calls sync: each action issues at most two syncs, but for
  // the throw-on scenario we throw on the very first sync within that action.
  let syncCallCount = 0;
  const throwOnActions: Record<string, number> = {
    clearBuiltinProps: 1,   // first sync for this action
    deleteCustomProps: 2,   // first sync after entering that branch
    deleteCustomXml: 4,     // each action may call sync twice; rough ordering
    deleteComments: 6,
  };

  const syncMock = vi.fn().mockImplementation(() => {
    syncCallCount++;
    if (throwOn) {
      const threshold = throwOnActions[throwOn];
      if (syncCallCount === threshold) {
        const e = new Error('GeneralException');
        e.name = 'GeneralException';
        return Promise.reject(e);
      }
    }
    return Promise.resolve(undefined);
  });

  const ctx = { document, sync: syncMock } as unknown as Word.RequestContext;
  return { ctx, builtinPropsObj, syncMock };
}

// ── Tests ────────────────────────────────────────────────────────────────────

describe('applyMetadata', () => {
  // ── 1. clearBuiltinProps ───────────────────────────────────────────────────
  it('clearBuiltinProps: clears all writable built-in props and returns done', async () => {
    const { ctx, builtinPropsObj } = makeFakeCtx({
      builtinProps: {
        author: 'Jane Smith',
        subject: 'Confidential Matter',
        title: 'Service Agreement',
        keywords: 'law nda',
        company: 'BigLaw LLP',
        manager: 'John Doe',
        comments: 'draft v3',
        category: 'Contracts',
      },
    });

    const finding = makeBaseFinding('metadata.builtinProperties', 'clearBuiltinProps');
    const results = await applyMetadata(ctx, [finding]);

    expect(results).toHaveLength(1);
    expect(results[0].result).toBe('done');

    // Verify each writable prop was set to ""
    expect(builtinPropsObj['author']).toBe('');
    expect(builtinPropsObj['subject']).toBe('');
    expect(builtinPropsObj['title']).toBe('');
    expect(builtinPropsObj['keywords']).toBe('');
    expect(builtinPropsObj['company']).toBe('');
    expect(builtinPropsObj['manager']).toBe('');
    expect(builtinPropsObj['comments']).toBe('');
    expect(builtinPropsObj['category']).toBe('');
  });

  // ── 2. deleteCustomProps ───────────────────────────────────────────────────
  it('deleteCustomProps: calls delete() on each custom property and returns done', async () => {
    const prop1 = { delete: vi.fn() };
    const prop2 = { delete: vi.fn() };
    const { ctx } = makeFakeCtx({ customPropItems: [prop1, prop2] });

    const finding = makeBaseFinding('metadata.customProperties', 'deleteCustomProps');
    const results = await applyMetadata(ctx, [finding]);

    expect(results).toHaveLength(1);
    expect(results[0].result).toBe('done');
    expect(prop1.delete).toHaveBeenCalledOnce();
    expect(prop2.delete).toHaveBeenCalledOnce();
  });

  // ── 3. deleteCustomXml ────────────────────────────────────────────────────
  it('deleteCustomXml: calls delete() on all 3 XML parts and returns done', async () => {
    const parts = [{ delete: vi.fn() }, { delete: vi.fn() }, { delete: vi.fn() }];
    const { ctx } = makeFakeCtx({ xmlPartItems: parts });

    const finding = makeBaseFinding('metadata.customXml', 'deleteCustomXml');
    const results = await applyMetadata(ctx, [finding]);

    expect(results).toHaveLength(1);
    expect(results[0].result).toBe('done');
    for (const part of parts) {
      expect(part.delete).toHaveBeenCalledOnce();
    }
  });

  // ── 4. deleteComments ─────────────────────────────────────────────────────
  it('deleteComments: calls delete() on all 5 comments and returns done', async () => {
    const comments = Array.from({ length: 5 }, () => ({ delete: vi.fn() }));
    const { ctx } = makeFakeCtx({ commentItems: comments });

    const finding = makeBaseFinding('metadata.comments', 'deleteComments');
    const results = await applyMetadata(ctx, [finding]);

    expect(results).toHaveLength(1);
    expect(results[0].result).toBe('done');
    for (const c of comments) {
      expect(c.delete).toHaveBeenCalledOnce();
    }
  });

  // ── 5. GeneralException on clearBuiltinProps ──────────────────────────────
  it('marks result:error on GeneralException and continues processing other findings', async () => {
    const prop1 = { delete: vi.fn() };

    // We build the ctx manually here so we can have one sync throw on the
    // clearBuiltinProps action (syncCallCount === 1) while the subsequent
    // deleteCustomProps action (syncCallCount 2 & 3) succeeds.
    const builtinPropsObj: Record<string, string> = { author: 'Alice' };
    const customProperties = { items: [prop1], load: vi.fn() };
    const properties = new Proxy(
      { customProperties, load: vi.fn() },
      {
        get(target, prop) {
          if (prop in target) return (target as Record<string | symbol, unknown>)[prop];
          return builtinPropsObj[prop as string];
        },
        set(_target, prop, value) {
          builtinPropsObj[prop as string] = value as string;
          return true;
        },
      }
    );
    const customXmlParts = { items: [], load: vi.fn() };
    const body = { getComments: () => ({ items: [], load: vi.fn() }) };
    const document = { properties, customXmlParts, body };

    let syncCallCount = 0;
    const syncMock = vi.fn().mockImplementation(() => {
      syncCallCount++;
      if (syncCallCount === 1) {
        const e = new Error('GeneralException');
        e.name = 'GeneralException';
        return Promise.reject(e);
      }
      return Promise.resolve(undefined);
    });

    const ctx = { document, sync: syncMock } as unknown as Word.RequestContext;

    const f1 = makeBaseFinding('metadata.builtinProperties', 'clearBuiltinProps');
    const f2 = makeBaseFinding('metadata.customProperties', 'deleteCustomProps');
    const results = await applyMetadata(ctx, [f1, f2]);

    expect(results).toHaveLength(2);
    expect(results[0].result).toBe('error');
    expect(results[1].result).toBe('done');
    // The function must not propagate the error to the caller
  });

  // ── 6. Clean document (empty collections) ────────────────────────────────
  it('handles empty custom props, XML parts, and comments without error (all done)', async () => {
    // All collections empty — delete-all on empty is safe
    const { ctx } = makeFakeCtx({
      customPropItems: [],
      xmlPartItems: [],
      commentItems: [],
    });

    const findings = [
      makeBaseFinding('metadata.builtinProperties', 'clearBuiltinProps'),
      makeBaseFinding('metadata.customProperties', 'deleteCustomProps'),
      makeBaseFinding('metadata.customXml', 'deleteCustomXml'),
      makeBaseFinding('metadata.comments', 'deleteComments'),
    ];

    const results = await applyMetadata(ctx, findings);

    expect(results).toHaveLength(4);
    for (const r of results) {
      expect(r.result).toBe('done');
    }
  });

  // ── 7. Unknown action → skipped ───────────────────────────────────────────
  it('marks result:skipped for unknown action types', async () => {
    const { ctx } = makeFakeCtx({});
    const finding = makeBaseFinding('misc.unknown', 'someUnhandledAction');
    const results = await applyMetadata(ctx, [finding]);

    expect(results).toHaveLength(1);
    expect(results[0].result).toBe('skipped');
  });

  // ── 8. Input findings are not mutated ─────────────────────────────────────
  it('does not mutate the input findings array', async () => {
    const { ctx } = makeFakeCtx({});
    const finding = makeBaseFinding('metadata.builtinProperties', 'clearBuiltinProps');
    const originalResult = finding.result;

    await applyMetadata(ctx, [finding]);

    expect(finding.result).toBe(originalResult);
  });
});
