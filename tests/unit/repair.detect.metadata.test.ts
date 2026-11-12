import { describe, it, expect } from 'vitest';
import { detectMetadata } from '@features/repair/detect/metadata.js';
import type { RepairFinding } from '@features/repair/types.js';

// ── Fake Office JS context helpers ──────────────────────────────────────────

/** A minimal `ClientResult<T>` stub whose `.value` is readable after sync. */
function clientResult<T>(value: T): OfficeExtension.ClientResult<T> {
  return { value } as unknown as OfficeExtension.ClientResult<T>;
}

/**
 * Build a fake `Word.RequestContext` that the metadata detector can drive.
 *
 * The fake `ctx.sync()` is a no-op Promise.resolve — all loads are resolved
 * immediately from the values supplied here, mirroring the Office JS pattern
 * where data is available synchronously after the matching sync() completes.
 */
function makeFakeCtx(opts: {
  builtinProps?: Partial<{
    author: string;
    lastAuthor: string;
    subject: string;
    title: string;
    keywords: string;
    company: string;
    manager: string;
    comments: string;
    category: string;
    revisionNumber: number;
  }>;
  customPropItems?: unknown[];
  xmlPartXmls?: string[];
  commentItems?: unknown[];
}): Word.RequestContext {
  const {
    builtinProps = {},
    customPropItems = [],
    xmlPartXmls = [],
    commentItems = [],
  } = opts;

  const bp = {
    author: '',
    lastAuthor: '',
    subject: '',
    title: '',
    keywords: '',
    company: '',
    manager: '',
    comments: '',
    category: '',
    revisionNumber: 0,
    load: () => undefined,
    customProperties: {
      items: customPropItems,
      load: () => undefined,
    },
    ...builtinProps,
  };

  const xmlItems = xmlPartXmls.map(xml => ({
    getXml: () => clientResult(xml),
  }));

  const xmlParts = {
    items: xmlItems,
    load: () => undefined,
  };

  const commentsCollection = {
    items: commentItems,
    load: () => undefined,
  };

  const body = {
    getComments: () => commentsCollection,
  };

  const document = {
    properties: bp,
    customXmlParts: xmlParts,
    body,
  };

  return {
    document,
    sync: () => Promise.resolve(),
  } as unknown as Word.RequestContext;
}

// ── Tests ────────────────────────────────────────────────────────────────────

describe('detectMetadata', () => {
  it('returns no findings for a completely clean document', async () => {
    const ctx = makeFakeCtx({});
    const findings = await detectMetadata(ctx);
    expect(findings).toHaveLength(0);
  });

  it('reports metadata.builtinProperties when built-in props are non-empty', async () => {
    const ctx = makeFakeCtx({
      builtinProps: { author: 'Jane Smith', company: 'Big Law LLP' },
    });
    const findings = await detectMetadata(ctx);
    const f = findings.find((x: RepairFinding) => x.id === 'metadata.builtinProperties');
    expect(f).toBeDefined();
    expect(f!.category).toBe('metadata');
    expect(f!.severity).toBe('confidentiality');
    expect(f!.selectedByDefault).toBe(true);
    expect(f!.action).toBe('clearBuiltinProps');
    expect(f!.detail).toContain('author');
    expect(f!.detail).toContain('company');
    expect(f!.count).toBe(2);
  });

  it('does not report builtinProperties when all fields are empty / zero', async () => {
    const ctx = makeFakeCtx({
      builtinProps: {
        author: '',
        lastAuthor: '',
        revisionNumber: 0,
        title: '',
        subject: '',
      },
    });
    const findings = await detectMetadata(ctx);
    expect(findings.find((x: RepairFinding) => x.id === 'metadata.builtinProperties')).toBeUndefined();
  });

  it('reports metadata.customProperties when custom props are present', async () => {
    const ctx = makeFakeCtx({
      customPropItems: [{ key: 'MatterNo', value: '12345' }, { key: 'Client', value: 'Acme' }],
    });
    const findings = await detectMetadata(ctx);
    const f = findings.find((x: RepairFinding) => x.id === 'metadata.customProperties');
    expect(f).toBeDefined();
    expect(f!.category).toBe('metadata');
    expect(f!.severity).toBe('confidentiality');
    expect(f!.count).toBe(2);
    expect(f!.detail).toContain('2 custom properties');
    expect(f!.action).toBe('deleteCustomProps');
    expect(f!.selectedByDefault).toBe(true);
  });

  it('uses singular "property" for a single custom property', async () => {
    const ctx = makeFakeCtx({ customPropItems: [{ key: 'MatterNo', value: '1' }] });
    const findings = await detectMetadata(ctx);
    const f = findings.find((x: RepairFinding) => x.id === 'metadata.customProperties');
    expect(f!.detail).toContain('1 custom property');
  });

  it('reports metadata.customXml when custom XML parts are present (generic)', async () => {
    const ctx = makeFakeCtx({
      xmlPartXmls: ['<root xmlns="http://example.com/schema"><data>value</data></root>'],
    });
    const findings = await detectMetadata(ctx);
    const f = findings.find((x: RepairFinding) => x.id === 'metadata.customXml');
    expect(f).toBeDefined();
    expect(f!.category).toBe('customXml');
    expect(f!.severity).toBe('confidentiality');
    expect(f!.count).toBe(1);
    expect(f!.detail).toContain('generic custom XML');
    expect(f!.action).toBe('deleteCustomXml');
    expect(f!.selectedByDefault).toBe(true);
  });

  it('tags iManage namespace in the custom XML detail', async () => {
    const ctx = makeFakeCtx({
      xmlPartXmls: ['<root xmlns:im="http://imanage.com/worksite"><im:MatterId>42</im:MatterId></root>'],
    });
    const findings = await detectMetadata(ctx);
    const f = findings.find((x: RepairFinding) => x.id === 'metadata.customXml');
    expect(f!.detail).toContain('iManage');
    expect(f!.detail).not.toContain('generic');
  });

  it('tags NetDocuments namespace in the custom XML detail', async () => {
    const ctx = makeFakeCtx({
      xmlPartXmls: ['<props xmlns="http://netdocuments.com/doc"><clientId>99</clientId></props>'],
    });
    const findings = await detectMetadata(ctx);
    const f = findings.find((x: RepairFinding) => x.id === 'metadata.customXml');
    expect(f!.detail).toContain('NetDocuments');
  });

  it('tags OpenText namespace in the custom XML detail', async () => {
    const ctx = makeFakeCtx({
      xmlPartXmls: ['<opentext:profile xmlns:opentext="http://opentext.com/edocs"><opentext:matter>M-1</opentext:matter></opentext:profile>'],
    });
    const findings = await detectMetadata(ctx);
    const f = findings.find((x: RepairFinding) => x.id === 'metadata.customXml');
    expect(f!.detail).toContain('OpenText');
  });

  it('tags multiple DMS namespaces when a document has parts from multiple systems', async () => {
    const ctx = makeFakeCtx({
      xmlPartXmls: [
        '<root xmlns="http://imanage.com/worksite"/>',
        '<root xmlns="http://netdocuments.com/nd"/>',
      ],
    });
    const findings = await detectMetadata(ctx);
    const f = findings.find((x: RepairFinding) => x.id === 'metadata.customXml');
    expect(f!.count).toBe(2);
    expect(f!.detail).toContain('iManage');
    expect(f!.detail).toContain('NetDocuments');
  });

  it('reports metadata.comments when comments are present', async () => {
    const ctx = makeFakeCtx({
      commentItems: [{ id: 'c1', authorName: 'Alice' }, { id: 'c2', authorName: 'Bob' }],
    });
    const findings = await detectMetadata(ctx);
    const f = findings.find((x: RepairFinding) => x.id === 'metadata.comments');
    expect(f).toBeDefined();
    expect(f!.category).toBe('comments');
    expect(f!.severity).toBe('confidentiality');
    expect(f!.count).toBe(2);
    expect(f!.detail).toContain('2 comments');
    expect(f!.action).toBe('deleteComments');
    expect(f!.selectedByDefault).toBe(true);
  });

  it('uses singular "comment" when there is exactly one comment', async () => {
    const ctx = makeFakeCtx({ commentItems: [{ id: 'c1' }] });
    const findings = await detectMetadata(ctx);
    const f = findings.find((x: RepairFinding) => x.id === 'metadata.comments');
    expect(f!.detail).toContain('1 comment');
    expect(f!.detail).not.toContain('comments');
  });

  it('returns all four findings for a fully populated document', async () => {
    const ctx = makeFakeCtx({
      builtinProps: { author: 'Alice' },
      customPropItems: [{ key: 'MatterNo', value: '1' }],
      xmlPartXmls: ['<root xmlns="http://imanage.com/worksite"/>'],
      commentItems: [{ id: 'c1' }],
    });
    const findings = await detectMetadata(ctx);
    const ids = findings.map((f: RepairFinding) => f.id);
    expect(ids).toContain('metadata.builtinProperties');
    expect(ids).toContain('metadata.customProperties');
    expect(ids).toContain('metadata.customXml');
    expect(ids).toContain('metadata.comments');
    expect(findings).toHaveLength(4);
  });

  it('all findings are marked destructive:false and requiresConfirm:false', async () => {
    const ctx = makeFakeCtx({
      builtinProps: { author: 'Alice' },
      customPropItems: [{ key: 'k', value: 'v' }],
      xmlPartXmls: ['<root/>'],
      commentItems: [{ id: 'c1' }],
    });
    const findings = await detectMetadata(ctx);
    for (const f of findings) {
      expect(f.destructive).toBe(false);
      expect(f.requiresConfirm).toBe(false);
    }
  });
});
