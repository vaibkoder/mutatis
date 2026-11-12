// src/features/repair/apply/spacing.ts
import type { RepairFinding } from '../types.js';
import type { StyleProfile } from '../profiles.js';

const CHUNK_SIZE = 200;

/** Convert profile lineSpacing multiple (1.0/1.5/2.0) to Office JS points (×12). */
function spacingToPt(multiple: number): number {
  return multiple * 12;
}

export async function applySpacing(
  ctx: Word.RequestContext,
  findings: RepairFinding[],
  profile: StyleProfile,
  options: { blankParaMode?: 'deleteAll' | 'collapseMultiple' }
): Promise<RepairFinding[]> {
  if (findings.length === 0) return [];

  const results: RepairFinding[] = [];

  for (const finding of findings) {
    switch (finding.action) {
      case 'normaliseSpacing': {
        await applyNormaliseSpacing(ctx, profile);
        results.push({ ...finding, result: 'done' });
        break;
      }

      case 'deleteBlankParas': {
        const mode = options.blankParaMode ?? 'collapseMultiple';
        await applyDeleteBlankParas(ctx, mode);
        results.push({ ...finding, result: 'done' });
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
// normaliseSpacing
// ---------------------------------------------------------------------------

async function applyNormaliseSpacing(
  ctx: Word.RequestContext,
  profile: StyleProfile
): Promise<void> {
  const body = ctx.document.body;
  body.paragraphs.load('items');
  await ctx.sync();

  const items = body.paragraphs.items;
  const { spaceBefore, spaceAfter, lineSpacing } = profile.body.spacing;

  for (let i = 0; i < items.length; i += CHUNK_SIZE) {
    const chunk = items.slice(i, i + CHUNK_SIZE);
    for (const p of chunk) {
      p.spaceBefore = spaceBefore;
      p.spaceAfter = spaceAfter;
      p.lineSpacing = spacingToPt(lineSpacing);
    }
    await ctx.sync();
  }
}

// ---------------------------------------------------------------------------
// deleteBlankParas
// ---------------------------------------------------------------------------

async function applyDeleteBlankParas(
  ctx: Word.RequestContext,
  mode: 'deleteAll' | 'collapseMultiple'
): Promise<void> {
  const body = ctx.document.body;

  if (mode === 'deleteAll') {
    body.paragraphs.load('items/text');
    await ctx.sync();

    const items = body.paragraphs.items;
    const toDelete: Word.Paragraph[] = [];

    for (let i = 0; i < items.length; i++) {
      const p = items[i]!;
      // Guard: never delete the very last paragraph.
      if (i === items.length - 1) continue;
      if (p.text.trim() === '') {
        toDelete.push(p);
      }
    }

    for (let i = 0; i < toDelete.length; i += CHUNK_SIZE) {
      const chunk = toDelete.slice(i, i + CHUNK_SIZE);
      for (const p of chunk) {
        p.delete();
      }
      await ctx.sync();
    }
  } else {
    // collapseMultiple mode
    body.paragraphs.load('items/text');
    await ctx.sync();

    const items = body.paragraphs.items;
    const total = items.length;

    // Identify which paragraphs to delete.
    // Walk in order, track consecutive blank runs.
    const toDelete: Word.Paragraph[] = [];

    let consecutiveBlanks = 0;

    for (let i = 0; i < total; i++) {
      const p = items[i]!;
      const isBlank = p.text.trim() === '';

      if (isBlank) {
        consecutiveBlanks++;
        if (consecutiveBlanks > 1) {
          // 2nd+ blank in a run → mark for deletion, guarding the very last paragraph.
          if (i !== total - 1) {
            toDelete.push(p);
          }
        }
        // First blank in a run is always retained (even if it precedes a Heading/Schedule/Signature —
        // the spec says "preserve a single blank before heading-like styles", which is satisfied by
        // the general rule of retaining exactly one blank per consecutive run).
      } else {
        // Non-blank: close any open run.
        consecutiveBlanks = 0;
      }
    }

    // Process deletions in chunks.
    for (let i = 0; i < toDelete.length; i += CHUNK_SIZE) {
      const chunk = toDelete.slice(i, i + CHUNK_SIZE);
      for (const p of chunk) {
        p.delete();
      }
      await ctx.sync();
    }
  }
}
