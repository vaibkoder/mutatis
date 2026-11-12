/**
 * Feature 06 upgrade (ADR-0023) — assisted removal engine (T4, M11/M12).
 * Constraints #1 (tracked changes mandatory + restore), #2 (batch guard),
 * #5 (leading-occurrence-only).
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { removeLeadingNumbers } from '@features/numbering/upgrade/removal.js';

// ---------------------------------------------------------------------------
// Minimal Word ctx mock
// ---------------------------------------------------------------------------

interface ParaOpts {
  /** Whether the in-paragraph search locates the leading slice (false → tab/NBSP no-match). */
  findHit?: boolean;
}

function makePara(text: string, opts: ParaOpts = {}) {
  const findHit = opts.findHit ?? true;
  const inserts: { text: string; loc: string }[] = [];
  const para = {
    text: '',
    _src: text,
    _inserts: inserts,
    load(prop: string) { if (prop === 'text') this.text = this._src; },
    getRange() {
      return {
        search(needle: string, _opts: unknown) {
          const hit = {
            insertText(t: string, loc: string) { inserts.push({ text: t, loc }); },
          };
          // Leading slice always sits at offset 0 → items[0] is the leading match.
          const items = findHit && needle.length > 0 ? [hit] : [];
          return { items, load(_p: string) { /* no-op */ } };
        },
      };
    },
  };
  return para;
}

interface CtxOpts {
  priorMode?: string;
  trackingLocked?: boolean;
}

function makeCtx(paras: ReturnType<typeof makePara>[], opts: CtxOpts = {}) {
  const modeHistory: string[] = [];
  let mode = opts.priorMode ?? 'Off';
  let pendingErr = false;

  const doc: Record<string, unknown> = {
    body: {},
    getSelection: () => ({ paragraphs: { items: paras } }),
  };
  Object.defineProperty(doc, 'changeTrackingMode', {
    get() { return mode; },
    set(v: string) {
      mode = v;
      modeHistory.push(v);
      if (opts.trackingLocked && v === 'TrackAll') pendingErr = true;
    },
  });

  const ctx = {
    document: doc,
    load: vi.fn(),
    async sync() {
      if (pendingErr) { pendingErr = false; throw new Error('cannot enable tracking'); }
    },
    _modeHistory: modeHistory,
  };
  return ctx as unknown as Word.RequestContext & { _modeHistory: string[] };
}

beforeEach(() => {
  (globalThis as unknown as { Word: unknown }).Word = {
    ChangeTrackingMode: { trackAll: 'TrackAll', off: 'Off' },
  };
});
afterEach(() => {
  delete (globalThis as unknown as { Word?: unknown }).Word;
});

// ---------------------------------------------------------------------------

describe('removeLeadingNumbers — tracked-change lifecycle (#1)', () => {
  it('enables Track Changes then restores the prior mode (round-trip)', async () => {
    const paras = [makePara('1. Introduction')];
    const ctx = makeCtx(paras, { priorMode: 'Off' });

    const report = await removeLeadingNumbers('auto', ctx);

    expect(report.disabled).toBe(false);
    expect(report.deleted).toBe(1);
    // trackAll set during the op, prior mode restored at the end.
    expect(ctx._modeHistory).toContain('TrackAll');
    expect(ctx._modeHistory.at(-1)).toBe('Off');
  });

  it('restores a non-Off prior mode (e.g. already TrackAll)', async () => {
    const paras = [makePara('1. Introduction')];
    const ctx = makeCtx(paras, { priorMode: 'TrackAll' });
    await removeLeadingNumbers('auto', ctx);
    expect(ctx._modeHistory.at(-1)).toBe('TrackAll');
  });

  it('performs NO deletes and reports disabled when tracking cannot be enabled', async () => {
    const paras = [makePara('1. Introduction')];
    const ctx = makeCtx(paras, { trackingLocked: true });

    const report = await removeLeadingNumbers('auto', ctx);

    expect(report.disabled).toBe(true);
    expect(report.deleted).toBe(0);
    expect(paras[0]!._inserts).toHaveLength(0); // nothing deleted
  });
});

describe('removeLeadingNumbers — batch guard (#2) & deletion', () => {
  it('deletes the leading slice via tracked insertText(\'\', \'Replace\')', async () => {
    const paras = [makePara('1.1 The parties agree.')];
    const ctx = makeCtx(paras);
    await removeLeadingNumbers('auto', ctx);
    expect(paras[0]!._inserts).toEqual([{ text: '', loc: 'Replace' }]);
  });

  it('skips a hanging paragraph with no leading number (no delete)', async () => {
    const paras = [makePara('provided that the Buyer shall pay')];
    const ctx = makeCtx(paras);
    const report = await removeLeadingNumbers('auto', ctx);
    expect(report.skipped).toBe(1);
    expect(report.deleted).toBe(0);
    expect(paras[0]!._inserts).toHaveLength(0);
  });

  it('mixes deleted and skipped across a batch', async () => {
    const paras = [
      makePara('1. First'),
      makePara('hanging continuation'),
      makePara('2. Second'),
    ];
    const ctx = makeCtx(paras);
    const report = await removeLeadingNumbers('auto', ctx);
    expect(report.total).toBe(3);
    expect(report.deleted).toBe(2);
    expect(report.skipped).toBe(1);
  });

  it('reports notFound (safe no-op) when the host search cannot match the slice', async () => {
    const paras = [makePara('1.5\tbody', { findHit: false })];
    const ctx = makeCtx(paras);
    const report = await removeLeadingNumbers('auto', ctx);
    expect(report.notFound).toBe(1);
    expect(report.deleted).toBe(0);
    expect(paras[0]!._inserts).toHaveLength(0);
  });
});
