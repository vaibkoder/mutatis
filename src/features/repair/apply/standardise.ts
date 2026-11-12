// src/features/repair/apply/standardise.ts
import type { RepairFinding } from '../types.js';
import type { StyleProfile } from '../profiles.js';

const CHUNK_SIZE = 200;

/** Convert profile lineSpacing multiple (1.0/1.5/2.0) to Office JS points (×12). */
function spacingToPt(multiple: number): number {
  return multiple * 12;
}

/**
 * Apply style standardisation to the document.
 *
 * Handles three actions:
 *   - `clearDirectFormatting`  — paragraph-level font/spacing reset (§5.3-B)
 *   - `stripInlineFormatting`  — run-level bold/italic/underline/color reset (§5.3-A),
 *                                only executed when options.stripInlineFormatting === true
 *   - `conformToProfile`       — update named styles via getStyles() (§5.4) and optionally
 *                                enforce per-section page margins (§5.5 margin part)
 *
 * Footnote/endnote paragraphs are propagated for clearDirectFormatting (§5.5).
 *
 * Writes are batched in windows of CHUNK_SIZE paragraphs per ctx.sync() (§8.6).
 * Runs INSIDE the caller-provided Word.RequestContext — do NOT call Word.run here.
 */
export async function applyStandardisation(
  ctx: Word.RequestContext,
  findings: RepairFinding[],
  profile: StyleProfile,
  options: { applyMargins?: boolean; stripInlineFormatting?: boolean },
): Promise<RepairFinding[]> {
  const results: RepairFinding[] = [];

  for (const finding of findings) {
    switch (finding.action) {
      case 'clearDirectFormatting': {
        try {
          await applyParagraphReset(ctx, profile);
          if (options.stripInlineFormatting === true) {
            await applyRunReset(ctx, profile);
          }
          await applyFootnoteEndnoteReset(ctx, profile);
          results.push({ ...finding, result: 'done' });
        } catch (err) {
          if (isGeneralException(err)) {
            results.push({ ...finding, result: 'error' });
          } else {
            throw err;
          }
        }
        break;
      }

      case 'stripInlineFormatting': {
        // Separate opt-in action — only executed when option flag is set.
        if (options.stripInlineFormatting !== true) {
          results.push({ ...finding, result: 'skipped' });
          break;
        }
        try {
          await applyRunReset(ctx, profile);
          results.push({ ...finding, result: 'done' });
        } catch (err) {
          if (isGeneralException(err)) {
            results.push({ ...finding, result: 'error' });
          } else {
            throw err;
          }
        }
        break;
      }

      case 'conformToProfile': {
        try {
          await applyNamedStyles(ctx, profile);
          if (options.applyMargins === true) {
            await applyMargins(ctx, profile);
          }
          results.push({ ...finding, result: 'done' });
        } catch (err) {
          if (isGeneralException(err)) {
            results.push({ ...finding, result: 'error' });
          } else {
            throw err;
          }
        }
        break;
      }

      default: {
        results.push({ ...finding, result: 'skipped' });
        break;
      }
    }
  }

  return results;
}

// ---------------------------------------------------------------------------
// §5.3-B  Paragraph-level reset
// ---------------------------------------------------------------------------

async function applyParagraphReset(
  ctx: Word.RequestContext,
  profile: StyleProfile,
): Promise<void> {
  const body = ctx.document.body;
  body.paragraphs.load('items');
  await ctx.sync();

  const allParas = body.paragraphs.items;
  for (let offset = 0; offset < allParas.length; offset += CHUNK_SIZE) {
    const chunk = allParas.slice(offset, offset + CHUNK_SIZE);
    // Load the properties we need to confirm before writing
    for (const p of chunk) {
      p.load('font/name,font/size,font/color,lineSpacing,spaceBefore,spaceAfter,leftIndent,rightIndent,firstLineIndent');
    }
    await ctx.sync();

    for (const p of chunk) {
      p.font.name = profile.body.font.name;
      p.font.size = profile.body.font.size;
      p.font.color = profile.body.font.color ?? '#000000';
      p.lineSpacing = spacingToPt(profile.body.spacing.lineSpacing);
      p.spaceBefore = profile.body.spacing.spaceBefore;
      p.spaceAfter = profile.body.spacing.spaceAfter;
      p.leftIndent = profile.body.indent.left;
      p.rightIndent = profile.body.indent.right;
      p.firstLineIndent = profile.body.indent.firstLine;
    }
    await ctx.sync();
  }
}

// ---------------------------------------------------------------------------
// §5.3-A  Run-level reset (bold/italic/underline/color/highlight)
// ---------------------------------------------------------------------------

async function applyRunReset(
  ctx: Word.RequestContext,
  profile: StyleProfile,
): Promise<void> {
  const body = ctx.document.body;
  body.paragraphs.load('items');
  await ctx.sync();

  const allParas = body.paragraphs.items;
  for (let offset = 0; offset < allParas.length; offset += CHUNK_SIZE) {
    const chunk = allParas.slice(offset, offset + CHUNK_SIZE);

    // Apply run-level reset via each paragraph's range — covers all inline runs.
    for (const p of chunk) {
      const range = p.getRange();
      range.font.bold = false;
      range.font.italic = false;
      range.font.underline = 'None' as Word.UnderlineType;
      range.font.color = profile.body.font.color ?? '#000000';
      range.font.highlightColor = 'None';
    }
    await ctx.sync();
  }
}

// ---------------------------------------------------------------------------
// §5.4  Named-style update
// ---------------------------------------------------------------------------

async function applyNamedStyles(
  ctx: Word.RequestContext,
  profile: StyleProfile,
): Promise<void> {
  const styles = ctx.document.getStyles();
  styles.load('items/nameLocal');
  await ctx.sync();

  for (const [styleName, def] of Object.entries(profile.styles)) {
    const style = styles.items.find(
      (s: Word.Style) => s.nameLocal === styleName,
    );
    if (!style) continue;

    if (def.font?.name) {
      style.font.name = def.font.name;
    }
    if (def.font?.size) {
      style.font.size = def.font.size;
    }
    if (def.paragraphFormat?.lineSpacing !== undefined) {
      style.paragraphFormat.lineSpacing = spacingToPt(def.paragraphFormat.lineSpacing);
    }
    if (def.paragraphFormat?.spaceBefore !== undefined) {
      style.paragraphFormat.spaceBefore = def.paragraphFormat.spaceBefore;
    }
    if (def.paragraphFormat?.spaceAfter !== undefined) {
      style.paragraphFormat.spaceAfter = def.paragraphFormat.spaceAfter;
    }
    if (def.paragraphFormat?.left !== undefined) {
      style.paragraphFormat.leftIndent = def.paragraphFormat.left;
    }
    if (def.paragraphFormat?.right !== undefined) {
      style.paragraphFormat.rightIndent = def.paragraphFormat.right;
    }
  }

  await ctx.sync();
}

// ---------------------------------------------------------------------------
// §5.5 (margin part)  Per-section margin enforcement
// ---------------------------------------------------------------------------

async function applyMargins(
  ctx: Word.RequestContext,
  profile: StyleProfile,
): Promise<void> {
  ctx.document.sections.load('items');
  await ctx.sync();

  for (const section of ctx.document.sections.items) {
    section.pageSetup.load('topMargin,bottomMargin,leftMargin,rightMargin');
  }
  await ctx.sync();

  for (const section of ctx.document.sections.items) {
    // DO NOT change orientation — only enforce the four margins
    section.pageSetup.topMargin = profile.page.margins.top;
    section.pageSetup.bottomMargin = profile.page.margins.bottom;
    section.pageSetup.leftMargin = profile.page.margins.left;
    section.pageSetup.rightMargin = profile.page.margins.right;
  }
  await ctx.sync();
}

// ---------------------------------------------------------------------------
// §5.5  Footnote/endnote propagation
// ---------------------------------------------------------------------------

async function applyFootnoteEndnoteReset(
  ctx: Word.RequestContext,
  profile: StyleProfile,
): Promise<void> {
  const body = ctx.document.body;
  body.footnotes.load('items');
  body.endnotes.load('items');
  await ctx.sync();

  const footnoteSize =
    profile.styles['Footnote Text']?.font?.size ?? profile.body.font.size;

  const allNotes = [
    ...body.footnotes.items,
    ...body.endnotes.items,
  ];

  for (const note of allNotes) {
    note.body.paragraphs.load('items');
    await ctx.sync();

    const noteParagraphs = note.body.paragraphs.items;
    for (let offset = 0; offset < noteParagraphs.length; offset += CHUNK_SIZE) {
      const chunk = noteParagraphs.slice(offset, offset + CHUNK_SIZE);

      for (const p of chunk) {
        p.load('font/name,font/size,lineSpacing');
      }
      await ctx.sync();

      for (const p of chunk) {
        p.font.name = profile.body.font.name;
        p.font.size = footnoteSize;
        p.lineSpacing = spacingToPt(profile.body.spacing.lineSpacing);
      }
      await ctx.sync();
    }
  }
}

// ---------------------------------------------------------------------------
// Utility
// ---------------------------------------------------------------------------

function isGeneralException(err: unknown): boolean {
  if (err instanceof Error && err.name === 'GeneralException') return true;
  if (
    typeof err === 'object' &&
    err !== null &&
    (err as Record<string, unknown>)['code'] === 'GeneralException'
  ) {
    return true;
  }
  return false;
}
