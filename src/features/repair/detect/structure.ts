// src/features/repair/detect/structure.ts
import type { RepairFinding } from '../types.js';
import type { StyleProfile } from '../profiles.js';
import type { DocumentHeuristics } from '../guards.js';

/**
 * Detect structural / stabilisation issues in the document:
 *   1. Direct-formatting overrides (font name/size, line spacing)
 *   2. Blank paragraphs
 *   3. Paragraph spacing anomalies (spaceBefore / spaceAfter)
 *   4. Section breaks (one finding per inter-section break)
 *
 * Must be called inside an existing Word.run context; the caller supplies `ctx`.
 * All Office JS loads and syncs are handled here.
 */
export async function detectStructure(
  ctx: Word.RequestContext,
  profile: StyleProfile,
): Promise<RepairFinding[]> {
  const findings: RepairFinding[] = [];

  // ── 1–3  Paragraph-level checks ────────────────────────────────────────────

  const paragraphs = ctx.document.body.paragraphs;
  paragraphs.load(
    'items/style,' +
    'items/font/name,' +
    'items/font/size,' +
    'items/lineSpacing,' +
    'items/spaceBefore,' +
    'items/spaceAfter,' +
    'items/text',
  );
  await ctx.sync();

  const bodyFont    = profile.body.font;
  const bodySpacing = profile.body.spacing;
  // Profile lineSpacing is a multiplier (1.0/1.5/2.0); Office JS uses points (×12).
  const profileLineSpacingPt = bodySpacing.lineSpacing * 12;

  let styleOverrideCount = 0;
  let blankCount         = 0;
  let spacingAnomalyCount = 0;

  for (const para of paragraphs.items) {
    // 1. Font / line-spacing overrides
    const fontMismatch =
      (para.font.name !== undefined &&
        para.font.name !== '' &&
        para.font.name !== bodyFont.name) ||
      (para.font.size !== undefined &&
        para.font.size !== 0 &&
        para.font.size !== bodyFont.size);

    const spacingMismatch =
      para.lineSpacing !== undefined &&
      para.lineSpacing !== 0 &&
      Math.abs(para.lineSpacing - profileLineSpacingPt) > 0.5;

    if (fontMismatch || spacingMismatch) {
      styleOverrideCount++;
    }

    // 2. Blank paragraphs
    if (para.text.trim() === '') {
      blankCount++;
    }

    // 3. Spacing anomalies (spaceBefore / spaceAfter)
    const spaceBeforeMismatch =
      para.spaceBefore !== undefined &&
      para.spaceBefore !== bodySpacing.spaceBefore;

    const spaceAfterMismatch =
      para.spaceAfter !== undefined &&
      para.spaceAfter !== bodySpacing.spaceAfter;

    if (spaceBeforeMismatch || spaceAfterMismatch) {
      spacingAnomalyCount++;
    }
  }

  if (styleOverrideCount > 0) {
    findings.push({
      id: 'styles.directFormatting',
      category: 'styles',
      severity: 'stabilisation',
      label: 'Direct formatting overrides present',
      detail: `${styleOverrideCount} paragraph${styleOverrideCount === 1 ? '' : 's'} have direct font or spacing overrides`,
      count: styleOverrideCount,
      destructive: false,
      requiresConfirm: false,
      selectedByDefault: false,
      action: 'clearDirectFormatting',
    });
  }

  if (blankCount > 0) {
    findings.push({
      id: 'blankParas.count',
      category: 'blankParas',
      severity: 'stabilisation',
      label: 'Blank paragraphs present',
      detail: `${blankCount} blank paragraph${blankCount === 1 ? '' : 's'}`,
      count: blankCount,
      destructive: false,
      requiresConfirm: false,
      selectedByDefault: true,
      action: 'deleteBlankParas',
    });
  }

  if (spacingAnomalyCount > 0) {
    findings.push({
      id: 'spacing.anomaly',
      category: 'spacing',
      severity: 'stabilisation',
      label: 'Paragraph spacing differs from profile',
      detail: `${spacingAnomalyCount} paragraph${spacingAnomalyCount === 1 ? '' : 's'} have non-profile spacing`,
      count: spacingAnomalyCount,
      destructive: false,
      requiresConfirm: false,
      selectedByDefault: false,
      action: 'normaliseSpacing',
    });
  }

  // ── 4  Section-break enumeration ───────────────────────────────────────────

  const sections = ctx.document.sections;
  sections.load('items');
  await ctx.sync();

  for (const section of sections.items) {
    section.body.paragraphs.load('items'); // not used, but pageSetup is on the section
    section.pageSetup.load(
      'topMargin,bottomMargin,leftMargin,rightMargin,orientation',
    );
  }
  await ctx.sync();

  // Emit one finding per section break, i.e. for sections[1..n-1] (the break
  // that precedes each section beyond the first).
  for (let i = 1; i < sections.items.length; i++) {
    const prev = sections.items[i - 1]!.pageSetup;
    const curr = sections.items[i]!.pageSetup;

    const orientationChanged = curr.orientation !== prev.orientation;
    const marginsChanged =
      curr.topMargin    !== prev.topMargin    ||
      curr.bottomMargin !== prev.bottomMargin ||
      curr.leftMargin   !== prev.leftMargin   ||
      curr.rightMargin  !== prev.rightMargin;

    let description: string;
    let isProtected: boolean;

    if (orientationChanged) {
      description = `${curr.orientation ?? 'unknown'} orientation`;
      isProtected = true;
    } else if (marginsChanged) {
      description = 'different margins';
      isProtected = false;
    } else {
      description = 'cosmetic';
      isProtected = false;
    }

    const label = orientationChanged
      ? `Section break (${description}) before section ${i + 1}`
      : marginsChanged
        ? `Section break (${description}) before section ${i + 1}`
        : `Section break before section ${i + 1}`;

    const detail = orientationChanged
      ? `Section ${i + 1} has ${curr.orientation} orientation vs ${prev.orientation} in the previous section`
      : marginsChanged
        ? `Section ${i + 1} has different page margins from section ${i}`
        : `Section break between section ${i} and section ${i + 1} appears cosmetic`;

    findings.push({
      id: `sectionBreaks.${i - 1}`,
      category: 'sectionBreaks',
      severity: 'stabilisation',
      label,
      detail,
      count: 1,
      destructive: true,
      requiresConfirm: true,
      protected: isProtected,
      selectedByDefault: false,
      action: 'removeSectionBreak',
    });
  }

  return findings;
}

/**
 * Derive document-level formatting heuristics used by the court-profile guards
 * (spec §7.2). Reads the modal (most common) font name/size and line spacing
 * among non-blank body paragraphs, and the smallest left/right page margin
 * across sections. Returns nulls when nothing can be determined (empty doc).
 *
 * Must be called inside an existing Word.run context; the caller supplies `ctx`.
 */
export async function detectHeuristics(
  ctx: Word.RequestContext,
): Promise<DocumentHeuristics> {
  const paragraphs = ctx.document.body.paragraphs;
  paragraphs.load('items/font/name,items/font/size,items/lineSpacing,items/text');

  const sections = ctx.document.sections;
  sections.load('items');
  await ctx.sync();

  for (const section of sections.items) {
    section.pageSetup.load('leftMargin,rightMargin');
  }
  await ctx.sync();

  // Tally font name / size / lineSpacing across non-blank paragraphs.
  const nameTally = new Map<string, number>();
  const sizeTally = new Map<number, number>();
  const spacingTally = new Map<number, number>();

  for (const para of paragraphs.items) {
    if (para.text.trim() === '') continue;
    const name = para.font.name;
    if (name) nameTally.set(name, (nameTally.get(name) ?? 0) + 1);
    const size = para.font.size;
    if (size) sizeTally.set(size, (sizeTally.get(size) ?? 0) + 1);
    const ls = para.lineSpacing;
    if (ls) spacingTally.set(ls, (spacingTally.get(ls) ?? 0) + 1);
  }

  const mode = <K>(tally: Map<K, number>): K | null => {
    let best: K | null = null;
    let bestN = 0;
    for (const [k, n] of tally) {
      if (n > bestN) { best = k; bestN = n; }
    }
    return best;
  };

  const detectedFontName = mode(nameTally);
  const detectedFontSize = mode(sizeTally);
  const modeSpacingPt = mode(spacingTally);
  // Office JS lineSpacing is in points; convert to a multiple (÷12).
  const detectedLineSpacing = modeSpacingPt !== null ? modeSpacingPt / 12 : null;

  // Smallest left/right margin across all sections.
  let detectedMarginPt: number | null = null;
  for (const section of sections.items) {
    const { leftMargin, rightMargin } = section.pageSetup;
    for (const m of [leftMargin, rightMargin]) {
      if (typeof m === 'number' && (detectedMarginPt === null || m < detectedMarginPt)) {
        detectedMarginPt = m;
      }
    }
  }

  return { detectedLineSpacing, detectedMarginPt, detectedFontSize, detectedFontName };
}
