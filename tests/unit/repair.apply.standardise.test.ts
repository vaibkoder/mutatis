// tests/unit/repair.apply.standardise.test.ts
import { describe, it, expect, vi } from 'vitest';
import { applyStandardisation } from '@features/repair/apply/standardise.js';
import type { RepairFinding } from '@features/repair/types.js';
import { SG_COURT_FILING } from '@features/repair/profiles.js';

// ---------------------------------------------------------------------------
// Mock helpers
// ---------------------------------------------------------------------------

interface MockParaProps {
  fontName?: string;
  fontSize?: number;
}

function makeMockPara(props: MockParaProps = {}): Record<string, unknown> {
  const loadSpy = vi.fn();
  const rangeFontSpy = {
    bold: true,
    italic: false,
    underline: 'single',
    color: '#FF0000',
    highlightColor: 'Yellow',
  };
  return {
    font: {
      name: props.fontName ?? 'Arial',
      size: props.fontSize ?? 14,
      color: '#FF0000',
    },
    lineSpacing: 12,
    spaceBefore: 6,
    spaceAfter: 6,
    leftIndent: 0,
    rightIndent: 0,
    firstLineIndent: 0,
    getRange: vi.fn(() => ({ font: rangeFontSpy })),
    load: loadSpy,
    _rangeFontSpy: rangeFontSpy,  // exposed for test assertions
  };
}

interface MockStyleProps {
  nameLocal: string;
  fontName?: string;
  fontSize?: number;
}

function makeMockStyle(props: MockStyleProps): Record<string, unknown> {
  return {
    nameLocal: props.nameLocal,
    font: {
      name: props.fontName ?? 'Arial',
      size: props.fontSize ?? 14,
    },
    paragraphFormat: {
      lineSpacingRule: 'single',
      lineSpacing: 1.0,
      spaceBefore: 0,
      spaceAfter: 0,
      leftIndent: 0,
      rightIndent: 0,
    },
  };
}

interface MockFootnoteProps {
  paraFontSize?: number;
}

function makeMockNote(props: MockFootnoteProps = {}): Record<string, unknown> {
  const para = makeMockPara({ fontSize: props.paraFontSize ?? 14 });
  return {
    body: {
      paragraphs: {
        load: vi.fn(),
        items: [para],
      },
    },
  };
}

interface MockSectionProps {
  topMargin?: number;
  bottomMargin?: number;
  leftMargin?: number;
  rightMargin?: number;
}

function makeMockSection(props: MockSectionProps = {}): Record<string, unknown> {
  return {
    pageSetup: {
      topMargin: props.topMargin ?? 72,
      bottomMargin: props.bottomMargin ?? 72,
      leftMargin: props.leftMargin ?? 72,
      rightMargin: props.rightMargin ?? 72,
      load: vi.fn(),
    },
  };
}

/**
 * Build a mock Word.RequestContext.
 * ctx.sync is a vi.fn() so tests can assert call counts.
 */
function makeCtx(opts: {
  paragraphs?: Record<string, unknown>[];
  styles?: Record<string, unknown>[];
  sections?: Record<string, unknown>[];
  footnotes?: Record<string, unknown>[];
  endnotes?: Record<string, unknown>[];
}): { ctx: Word.RequestContext; syncSpy: ReturnType<typeof vi.fn> } {
  const syncSpy = vi.fn().mockResolvedValue(undefined);
  const paragraphs = opts.paragraphs ?? [];
  const styles = opts.styles ?? [];
  const sections = opts.sections ?? [];
  const footnotes = opts.footnotes ?? [];
  const endnotes = opts.endnotes ?? [];

  const ctx = {
    document: {
      body: {
        paragraphs: {
          load: vi.fn(),
          items: paragraphs,
        },
        footnotes: {
          load: vi.fn(),
          items: footnotes,
        },
        endnotes: {
          load: vi.fn(),
          items: endnotes,
        },
      },
      getStyles: () => ({
        load: vi.fn(),
        items: styles,
      }),
      sections: {
        load: vi.fn(),
        items: sections,
      },
    },
    sync: syncSpy,
  } as unknown as Word.RequestContext;

  return { ctx, syncSpy };
}

// ---------------------------------------------------------------------------
// Shared finding factories
// ---------------------------------------------------------------------------

function makeFinding(overrides: Partial<RepairFinding> = {}): RepairFinding {
  return {
    id: 'styles.directFormatting',
    category: 'styles',
    severity: 'stabilisation',
    label: 'Direct formatting',
    destructive: false,
    requiresConfirm: false,
    selectedByDefault: false,
    action: 'clearDirectFormatting',
    ...overrides,
  };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('applyStandardisation', () => {
  const profile = SG_COURT_FILING;

  // 1. clearDirectFormatting: paragraph font.name assigned from profile
  it('clearDirectFormatting assigns paragraph font.name from profile', async () => {
    const para = makeMockPara({ fontName: 'Arial', fontSize: 14 });
    const { ctx } = makeCtx({ paragraphs: [para] });

    const finding = makeFinding({ action: 'clearDirectFormatting' });
    const results = await applyStandardisation(ctx, [finding], profile, {});

    expect((para.font as Record<string, unknown>).name).toBe(profile.body.font.name);
    expect(results[0].result).toBe('done');
  });

  // 2. clearDirectFormatting: font size assigned per profile body.font.size
  it('clearDirectFormatting assigns paragraph font.size from profile', async () => {
    const para = makeMockPara({ fontSize: 9 });
    const { ctx } = makeCtx({ paragraphs: [para] });

    const finding = makeFinding({ action: 'clearDirectFormatting' });
    await applyStandardisation(ctx, [finding], profile, {});

    expect((para.font as Record<string, unknown>).size).toBe(profile.body.font.size);
  });

  // 3. conformToProfile: getStyles() called, matching style has font.name updated
  it('conformToProfile updates font.name on matching named style', async () => {
    const style = makeMockStyle({ nameLocal: 'Normal', fontName: 'Arial' });
    const { ctx } = makeCtx({ styles: [style] });

    const finding = makeFinding({ action: 'conformToProfile' });
    const results = await applyStandardisation(ctx, [finding], profile, {});

    const profileNormal = profile.styles['Normal'];
    expect((style.font as Record<string, unknown>).name).toBe(profileNormal?.font?.name);
    expect(results[0].result).toBe('done');
  });

  // 4. Margin enforcement (when applyMargins:true): sections.topMargin assigned per profile
  it('enforces page margins when applyMargins:true', async () => {
    const section = makeMockSection({ topMargin: 72 });
    const { ctx } = makeCtx({ sections: [section] });

    const finding = makeFinding({ action: 'conformToProfile' });
    await applyStandardisation(ctx, [finding], profile, { applyMargins: true });

    expect((section.pageSetup as Record<string, unknown>).topMargin).toBe(
      profile.page.margins.top,
    );
    expect((section.pageSetup as Record<string, unknown>).bottomMargin).toBe(
      profile.page.margins.bottom,
    );
    expect((section.pageSetup as Record<string, unknown>).leftMargin).toBe(
      profile.page.margins.left,
    );
    expect((section.pageSetup as Record<string, unknown>).rightMargin).toBe(
      profile.page.margins.right,
    );
  });

  // 5. Margin NOT enforced when applyMargins:false (or omitted)
  it('does NOT touch section margins when applyMargins is false or omitted', async () => {
    const originalTop = 72;
    const section = makeMockSection({ topMargin: originalTop });
    const { ctx } = makeCtx({ sections: [section] });

    const finding = makeFinding({ action: 'conformToProfile' });
    await applyStandardisation(ctx, [finding], profile, { applyMargins: false });

    expect((section.pageSetup as Record<string, unknown>).topMargin).toBe(originalTop);
  });

  // 6. Footnote propagation: footnote body paragraphs get footnoteSize (TNR 10 for SG profile)
  it('propagates font to footnote body paragraphs using Footnote Text style size', async () => {
    const note = makeMockNote({ paraFontSize: 14 });
    const { ctx } = makeCtx({ footnotes: [note] });

    const finding = makeFinding({ action: 'clearDirectFormatting' });
    await applyStandardisation(ctx, [finding], profile, {});

    const notePara = (note.body as Record<string, unknown>).paragraphs as { items: Record<string, unknown>[] };
    const para = notePara.items[0];
    const expectedSize = profile.styles['Footnote Text']?.font?.size ?? profile.body.font.size;

    // SG profile has Footnote Text font size = 10
    expect(expectedSize).toBe(10);
    expect((para.font as Record<string, unknown>).size).toBe(10);
    expect((para.font as Record<string, unknown>).name).toBe(profile.body.font.name);
  });

  // 7. stripInlineFormatting:true → run-level bold=false set via getRange()
  it('sets run bold=false when stripInlineFormatting:true', async () => {
    const para = makeMockPara();
    const { ctx } = makeCtx({ paragraphs: [para] });

    const finding = makeFinding({ action: 'clearDirectFormatting' });
    await applyStandardisation(ctx, [finding], profile, { stripInlineFormatting: true });

    // getRange() is called and its font.bold reset
    expect(para.getRange).toHaveBeenCalled();
    expect((para._rangeFontSpy as Record<string, unknown>).bold).toBe(false);
    expect((para._rangeFontSpy as Record<string, unknown>).italic).toBe(false);
  });

  // 8. stripInlineFormatting:false (default) → getRange() NOT called (emphasis preserved)
  it('does NOT reset run bold when stripInlineFormatting is false (default)', async () => {
    const para = makeMockPara();
    const { ctx } = makeCtx({ paragraphs: [para] });

    const finding = makeFinding({ action: 'clearDirectFormatting' });
    await applyStandardisation(ctx, [finding], profile, {}); // stripInlineFormatting omitted

    // getRange() should NOT have been called — we didn't touch run-level formatting
    expect(para.getRange).not.toHaveBeenCalled();
  });

  // 9. Chunking: 250 paragraphs in body → ctx.sync called multiple times (>= 2 in paragraph loop)
  it('calls ctx.sync multiple times for 250 paragraphs (chunking at 200)', async () => {
    const paragraphs = Array.from({ length: 250 }, () => makeMockPara());
    const { ctx, syncSpy } = makeCtx({ paragraphs });

    const finding = makeFinding({ action: 'clearDirectFormatting' });
    await applyStandardisation(ctx, [finding], profile, {});

    // The paragraph reset loop alone makes 2 iterations:
    //   chunk 0-199: load + sync, write + sync  → 2 syncs
    //   chunk 200-249: load + sync, write + sync → 2 syncs
    // Plus: initial paragraphs.load sync, footnotes/endnotes load sync = total >= 4
    // We just assert there are at least 4 syncs to confirm chunking occurred.
    expect(syncSpy.mock.calls.length).toBeGreaterThanOrEqual(4);
  });

  // Bonus: unknown action → skipped
  it('returns result:skipped for an unrecognised action', async () => {
    const { ctx } = makeCtx({});
    const finding = makeFinding({ action: 'unknownAction' });
    const results = await applyStandardisation(ctx, [finding], profile, {});
    expect(results[0].result).toBe('skipped');
  });

  // Bonus: GeneralException → result:error (not thrown)
  it('returns result:error and does not rethrow on GeneralException', async () => {
    const { ctx, syncSpy } = makeCtx({});
    syncSpy.mockRejectedValueOnce(Object.assign(new Error('fail'), { name: 'GeneralException' }));

    const finding = makeFinding({ action: 'clearDirectFormatting' });
    const results = await applyStandardisation(ctx, [finding], profile, {});
    expect(results[0].result).toBe('error');
  });

  // Bonus: stripInlineFormatting as standalone action — skipped when option is false
  it('skips standalone stripInlineFormatting action when option is false', async () => {
    // `makeMockPara` takes no runs, so a run built here would never reach
    // `applyStandardisation` and an assertion that it was "untouched" would be
    // about a free-standing object. This asserts the skip, which is the claim.
    const para = makeMockPara();
    const { ctx } = makeCtx({ paragraphs: [para] });

    const finding = makeFinding({ action: 'stripInlineFormatting' });
    const results = await applyStandardisation(ctx, [finding], profile, { stripInlineFormatting: false });

    expect(results[0].result).toBe('skipped');
  });

  // Bonus: stripInlineFormatting as standalone action — executes when option is true
  it('executes standalone stripInlineFormatting action when option is true', async () => {
    const para = makeMockPara();
    const { ctx } = makeCtx({ paragraphs: [para] });

    const finding = makeFinding({ action: 'stripInlineFormatting' });
    const results = await applyStandardisation(ctx, [finding], profile, { stripInlineFormatting: true });

    expect(results[0].result).toBe('done');
    expect(para.getRange).toHaveBeenCalled();
    expect((para._rangeFontSpy as Record<string, unknown>).bold).toBe(false);
  });
});
