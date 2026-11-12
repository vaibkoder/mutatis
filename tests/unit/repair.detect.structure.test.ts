// tests/unit/repair.detect.structure.test.ts
import { describe, it, expect } from 'vitest';
import { detectStructure, detectHeuristics } from '../../src/features/repair/detect/structure.js';
import type { StyleProfile } from '../../src/features/repair/profiles.js';
import { SG_COURT_FILING } from '../../src/features/repair/profiles.js';

// ---------------------------------------------------------------------------
// Mock helpers
// ---------------------------------------------------------------------------

interface MockParaProps {
  text: string;
  fontName?: string;
  fontSize?: number;
  /** Line spacing in Office JS points (12=single, 18=1.5x, 24=double). Defaults to profile value. */
  lineSpacing?: number;
  spaceBefore?: number;
  spaceAfter?: number;
}

function makePara(props: MockParaProps) {
  const profile = SG_COURT_FILING.body;
  // Profile lineSpacing is a multiplier; Office JS uses points (×12).
  const profileLineSpacingPt = profile.spacing.lineSpacing * 12;
  return {
    text: props.text,
    font: {
      name: props.fontName ?? profile.font.name,
      size: props.fontSize ?? profile.font.size,
    },
    lineSpacing:  props.lineSpacing  ?? profileLineSpacingPt,
    spaceBefore:  props.spaceBefore  ?? profile.spacing.spaceBefore,
    spaceAfter:   props.spaceAfter   ?? profile.spacing.spaceAfter,
  };
}

interface MockSectionProps {
  orientation: 'portrait' | 'landscape';
  topMargin?:    number;
  bottomMargin?: number;
  leftMargin?:   number;
  rightMargin?:  number;
}

function makeSection(props: MockSectionProps) {
  return {
    body: { paragraphs: { load: () => undefined, items: [] } },
    pageSetup: {
      orientation:  props.orientation,
      topMargin:    props.topMargin    ?? 99.21,
      bottomMargin: props.bottomMargin ?? 99.21,
      leftMargin:   props.leftMargin   ?? 99.21,
      rightMargin:  props.rightMargin  ?? 99.21,
      load: () => undefined,
    },
  };
}

/**
 * Build a minimal mock Word.RequestContext.
 * `syncs` tracks how many times ctx.sync() was called so tests can assert
 * on load/sync ordering if needed.
 */
function makeCtx(paragraphs: ReturnType<typeof makePara>[], sections: ReturnType<typeof makeSection>[]) {
  return {
    document: {
      body: {
        paragraphs: {
          load: () => undefined,
          items: paragraphs,
        },
      },
      sections: {
        load: () => undefined,
        items: sections,
      },
    },
    sync: async () => undefined,
  } as unknown as Word.RequestContext;
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('detectStructure', () => {
  const profile: StyleProfile = SG_COURT_FILING;

  // 1. Font mismatch → styles.directFormatting
  it('emits styles.directFormatting when a paragraph has a mismatched font name', async () => {
    const paras = [
      makePara({ text: 'Normal body text.' }),
      makePara({ text: 'Wrong font.', fontName: 'Arial' }),
    ];
    const ctx = makeCtx(paras, [makeSection({ orientation: 'portrait' })]);

    const findings = await detectStructure(ctx, profile);
    const f = findings.find(x => x.id === 'styles.directFormatting');

    expect(f).toBeDefined();
    expect(f!.category).toBe('styles');
    expect(f!.severity).toBe('stabilisation');
    expect(f!.count).toBe(1);
    expect(f!.selectedByDefault).toBe(false);
    expect(f!.destructive).toBe(false);
    expect(f!.action).toBe('clearDirectFormatting');
  });

  it('emits styles.directFormatting when a paragraph has a mismatched font size', async () => {
    const paras = [
      makePara({ text: 'Para 1.', fontSize: 14 }),
      makePara({ text: 'Para 2.', fontSize: 14 }),
      makePara({ text: 'Para 3.' }),
    ];
    const ctx = makeCtx(paras, [makeSection({ orientation: 'portrait' })]);

    const findings = await detectStructure(ctx, profile);
    const f = findings.find(x => x.id === 'styles.directFormatting');

    expect(f).toBeDefined();
    expect(f!.count).toBe(2);
    expect(f!.detail).toMatch(/2 paragraphs/);
  });

  // 2. Blank paragraph → blankParas.count
  it('emits blankParas.count for blank paragraphs', async () => {
    const paras = [
      makePara({ text: 'Some text.' }),
      makePara({ text: '' }),
      makePara({ text: '   ' }),
    ];
    const ctx = makeCtx(paras, [makeSection({ orientation: 'portrait' })]);

    const findings = await detectStructure(ctx, profile);
    const f = findings.find(x => x.id === 'blankParas.count');

    expect(f).toBeDefined();
    expect(f!.category).toBe('blankParas');
    expect(f!.count).toBe(2);
    expect(f!.selectedByDefault).toBe(true);
    expect(f!.destructive).toBe(false);
    expect(f!.action).toBe('deleteBlankParas');
    expect(f!.detail).toMatch(/2 blank paragraphs/);
  });

  // 3. Spacing mismatch → spacing.anomaly
  it('emits spacing.anomaly when spaceBefore or spaceAfter differs from profile', async () => {
    const paras = [
      makePara({ text: 'Good para.' }),
      makePara({ text: 'Bad spacing.', spaceBefore: 12 }),
    ];
    const ctx = makeCtx(paras, [makeSection({ orientation: 'portrait' })]);

    const findings = await detectStructure(ctx, profile);
    const f = findings.find(x => x.id === 'spacing.anomaly');

    expect(f).toBeDefined();
    expect(f!.category).toBe('spacing');
    expect(f!.count).toBe(1);
    expect(f!.selectedByDefault).toBe(false);
    expect(f!.destructive).toBe(false);
    expect(f!.action).toBe('normaliseSpacing');
    expect(f!.detail).toMatch(/1 paragraph/);
  });

  it('emits spacing.anomaly when spaceAfter differs from profile', async () => {
    const paras = [
      makePara({ text: 'Para 1.', spaceAfter: 6 }),
      makePara({ text: 'Para 2.', spaceAfter: 6 }),
    ];
    const ctx = makeCtx(paras, [makeSection({ orientation: 'portrait' })]);

    const findings = await detectStructure(ctx, profile);
    const f = findings.find(x => x.id === 'spacing.anomaly');

    // SG_COURT_FILING expects spaceAfter: 0; paragraphs have 6 → anomaly
    expect(f).toBeDefined();
    expect(f!.count).toBe(2);
  });

  // 4. Clean paragraphs under profile → no styles/blanks/spacing findings
  it('emits no paragraph findings when all paragraphs match the profile', async () => {
    const paras = [
      makePara({ text: 'Paragraph one.' }),
      makePara({ text: 'Paragraph two.' }),
    ];
    const ctx = makeCtx(paras, [makeSection({ orientation: 'portrait' })]);

    const findings = await detectStructure(ctx, profile);

    expect(findings.find(x => x.id === 'styles.directFormatting')).toBeUndefined();
    expect(findings.find(x => x.id === 'blankParas.count')).toBeUndefined();
    expect(findings.find(x => x.id === 'spacing.anomaly')).toBeUndefined();
  });

  // 5. Two sections with different orientation → sectionBreaks.0 with protected:true
  it('emits sectionBreaks.0 with protected:true when orientation changes', async () => {
    const sections = [
      makeSection({ orientation: 'portrait' }),
      makeSection({ orientation: 'landscape' }),
    ];
    const ctx = makeCtx([], sections);

    const findings = await detectStructure(ctx, profile);
    const f = findings.find(x => x.id === 'sectionBreaks.0');

    expect(f).toBeDefined();
    expect(f!.category).toBe('sectionBreaks');
    expect(f!.protected).toBe(true);
    expect(f!.destructive).toBe(true);
    expect(f!.requiresConfirm).toBe(true);
    expect(f!.selectedByDefault).toBe(false);
    expect(f!.action).toBe('removeSectionBreak');
    expect(f!.label).toMatch(/landscape/i);
  });

  // 6. Two sections same geometry → section break finding with protected:false
  it('emits sectionBreaks.0 with protected:false when orientation is the same', async () => {
    const sections = [
      makeSection({ orientation: 'portrait' }),
      makeSection({ orientation: 'portrait' }),
    ];
    const ctx = makeCtx([], sections);

    const findings = await detectStructure(ctx, profile);
    const f = findings.find(x => x.id === 'sectionBreaks.0');

    expect(f).toBeDefined();
    expect(f!.protected).toBe(false);
    expect(f!.destructive).toBe(true);
  });

  // Additional: margin-only change → protected:false
  it('emits sectionBreaks.0 with protected:false when only margins differ', async () => {
    const sections = [
      makeSection({ orientation: 'portrait', topMargin: 99.21 }),
      makeSection({ orientation: 'portrait', topMargin: 72 }),
    ];
    const ctx = makeCtx([], sections);

    const findings = await detectStructure(ctx, profile);
    const f = findings.find(x => x.id === 'sectionBreaks.0');

    expect(f).toBeDefined();
    expect(f!.protected).toBe(false);
    expect(f!.label).toMatch(/different margins/i);
  });

  // Additional: three sections → two break findings
  it('emits one finding per inter-section break for three sections', async () => {
    const sections = [
      makeSection({ orientation: 'portrait' }),
      makeSection({ orientation: 'landscape' }),
      makeSection({ orientation: 'portrait' }),
    ];
    const ctx = makeCtx([], sections);

    const findings = await detectStructure(ctx, profile);
    const breakFindings = findings.filter(x => x.category === 'sectionBreaks');

    expect(breakFindings).toHaveLength(2);
    expect(breakFindings[0].id).toBe('sectionBreaks.0');
    expect(breakFindings[1].id).toBe('sectionBreaks.1');
    expect(breakFindings[0].protected).toBe(true);   // portrait→landscape
    expect(breakFindings[1].protected).toBe(true);   // landscape→portrait
  });

  // Additional: single section → no section break findings
  it('emits no sectionBreaks findings for a document with one section', async () => {
    const ctx = makeCtx([], [makeSection({ orientation: 'portrait' })]);
    const findings = await detectStructure(ctx, profile);
    expect(findings.filter(x => x.category === 'sectionBreaks')).toHaveLength(0);
  });
});

describe('detectHeuristics', () => {
  it('reports modal font/size/spacing and smallest left/right margin', async () => {
    // SG court look: TNR 12, double-spaced (24pt = 2.0×12), 99.21pt margins.
    const paras = [
      makePara({ text: 'Body one.', fontName: 'Times New Roman', fontSize: 12, lineSpacing: 24 }),
      makePara({ text: 'Body two.', fontName: 'Times New Roman', fontSize: 12, lineSpacing: 24 }),
      makePara({ text: '   ' }),  // blank — ignored
    ];
    const ctx = makeCtx(paras, [makeSection({ orientation: 'portrait' })]);

    const h = await detectHeuristics(ctx);

    expect(h.detectedFontName).toBe('Times New Roman');
    expect(h.detectedFontSize).toBe(12);
    expect(h.detectedLineSpacing).toBeCloseTo(2.0, 5);  // 24pt ÷ 12
    expect(h.detectedMarginPt).toBeCloseTo(99.21, 1);
  });

  it('picks the most common values when paragraphs are mixed', async () => {
    const paras = [
      makePara({ text: 'A', fontName: 'Times New Roman', fontSize: 12 }),
      makePara({ text: 'B', fontName: 'Times New Roman', fontSize: 12 }),
      makePara({ text: 'C', fontName: 'Arial', fontSize: 11 }),
    ];
    const ctx = makeCtx(paras, [makeSection({ orientation: 'portrait' })]);

    const h = await detectHeuristics(ctx);

    expect(h.detectedFontName).toBe('Times New Roman');
    expect(h.detectedFontSize).toBe(12);
  });

  it('uses the smallest left/right margin across sections', async () => {
    const paras = [makePara({ text: 'Body.' })];
    const sections = [
      makeSection({ orientation: 'portrait', leftMargin: 99.21, rightMargin: 99.21 }),
      makeSection({ orientation: 'portrait', leftMargin: 72, rightMargin: 90 }),
    ];
    const ctx = makeCtx(paras, sections);

    const h = await detectHeuristics(ctx);
    expect(h.detectedMarginPt).toBe(72);
  });

  it('returns nulls for an empty document (no non-blank paragraphs)', async () => {
    const ctx = makeCtx([], []);
    const h = await detectHeuristics(ctx);

    expect(h.detectedFontName).toBeNull();
    expect(h.detectedFontSize).toBeNull();
    expect(h.detectedLineSpacing).toBeNull();
    expect(h.detectedMarginPt).toBeNull();
  });
});
