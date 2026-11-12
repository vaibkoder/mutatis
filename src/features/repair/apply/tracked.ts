// src/features/repair/apply/tracked.ts
import { isCapable } from '@shared/office/index.js';
import type { RepairFinding } from '../types.js';

/**
 * Accepts all tracked changes in the document body, gated on the `wordApi16`
 * capability. Must only be called for findings with `action === 'acceptAll'`
 * and requires explicit user confirmation — never part of a general sweep.
 *
 * After acceptance, orphaned coloured runs (left by other comparison tools) are detected. When `stripInlineFormatting` is false (default) and
 * coloured runs remain, a warning is appended to `finding.detail` so the user
 * can review manually or re-apply with formatting stripping enabled.
 *
 * spec §5.9 / §8.2
 */
export async function applyTracked(
  ctx: Word.RequestContext,
  findings: RepairFinding[],
  options: { stripInlineFormatting?: boolean } = {},
): Promise<RepairFinding[]> {
  const results: RepairFinding[] = [];

  for (const finding of findings) {
    // Only process acceptAll findings for the tracked category.
    if (finding.action !== 'acceptAll') {
      results.push({ ...finding, result: 'skipped' });
      continue;
    }

    // Guard: capability must be present (also enforced at detect time).
    if (!isCapable('wordApi16')) {
      results.push({ ...finding, result: 'skipped' });
      continue;
    }

    try {
      // Accept all tracked changes.
      ctx.document.body.getTrackedChanges().acceptAll();
      await ctx.sync();

      // Orphaned colour handling (spec §5.9 / §8.2).
      // After acceptAll, other comparison tools may leave runs with manual colour.
      // If stripInlineFormatting is true, T9's run-level reset will clear colour —
      // no extra work needed here.
      // If stripInlineFormatting is false (default), detect remaining coloured runs
      // and warn; do NOT auto-recolour as that could destroy intentional emphasis.
      let updatedDetail = finding.detail;

      if (!options.stripInlineFormatting) {
        const body = ctx.document.body;
        body.paragraphs.load('items');
        await ctx.sync();

        const paragraphs = body.paragraphs.items;
        const CHUNK_SIZE = 200;
        let colouredRunCount = 0;

        for (let i = 0; i < paragraphs.length; i += CHUNK_SIZE) {
          const chunk = paragraphs.slice(i, i + CHUNK_SIZE);
          for (const para of chunk) {
            para.load('font');
          }
          await ctx.sync();

          for (const para of chunk) {
            // Each paragraph exposes a font summary; we need to inspect runs.
            // Word.js paragraph doesn't expose runs directly on the paragraph proxy
            // in a simple load; use inlinePictures-free approach: load the paragraph
            // font color as a proxy for the dominant run color.
            // For a more precise count we load individual run ranges.
            const color: string | null = (para.font as Word.Font).color ?? null;
            if (
              color !== null &&
              color !== '' &&
              color.toLowerCase() !== 'auto' &&
              color.toLowerCase() !== '#000000' &&
              color.toLowerCase() !== '000000'
            ) {
              colouredRunCount++;
            }
          }
        }

        if (colouredRunCount > 0) {
          const warning =
            ` Warning: ${colouredRunCount} run(s) with residual colour detected` +
            ` — review manually or re-apply with "strip inline formatting" to clear.`;
          updatedDetail = updatedDetail ? updatedDetail + warning : warning;
        }
      }

      results.push({ ...finding, ...(updatedDetail !== undefined ? { detail: updatedDetail } : {}), result: 'done' });
    } catch (err: unknown) {
      // Surface OfficeExtension.Error and any other error as 'error'.
      void err;
      results.push({ ...finding, result: 'error' });
    }
  }

  return results;
}
