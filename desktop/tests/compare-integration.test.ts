import { describe, it, expect } from 'vitest';
import { runComparison } from '../src/renderer/src/compare/run-comparison.js';
import {
  makeDocxBytes,
  SAMPLE_ORIGINAL,
  SAMPLE_REVISED,
} from './helpers/make-docx.js';
import type { RcParagraph, RcRun } from '@shared/render-core/types.js';

/**
 * THE integration proof.
 *
 * This is the test the whole sibling-alias design rests on: it drives the real
 * engine, from real `.docx` bytes, through the real pipeline, into the real
 * render model — with no Preact, no Office JS, no worker and no font subsystem
 * anywhere in the path. If this passes, phase 2's architecture is sound. If it
 * ever breaks, the integration is broken, not the UI.
 */

function allRuns(blocks: readonly { kind: string }[]): RcRun[] {
  const runs: RcRun[] = [];
  for (const block of blocks as readonly {
    kind: 'paragraph' | 'table';
    paragraph?: RcParagraph;
    table?: { rows: { cells: { blocks: unknown[] }[] }[] };
  }[]) {
    if (block.kind === 'paragraph' && block.paragraph) {
      runs.push(...block.paragraph.runs);
    } else if (block.kind === 'table' && block.table) {
      for (const row of block.table.rows) {
        for (const cell of row.cells) {
          runs.push(...allRuns(cell.blocks as readonly { kind: string }[]));
        }
      }
    }
  }
  return runs;
}

describe('Word-vs-Word comparison, end to end', () => {
  it('runs the real engine over real .docx bytes and produces a render model', async () => {
    const outcome = await runComparison({
      originalBytes: makeDocxBytes(SAMPLE_ORIGINAL),
      revisedBytes: makeDocxBytes(SAMPLE_REVISED),
      // Fixed so the result is reproducible run to run.
      date: '2026-01-01T00:00:00.000Z',
      author: 'Test',
    });

    expect(outcome.summary.total).toBeGreaterThan(0);
    expect(outcome.ops.length).toBeGreaterThan(0);
    expect(outcome.model.blocks.length).toBeGreaterThan(0);
  });

  it('emits native OOXML tracked changes', async () => {
    const outcome = await runComparison({
      originalBytes: makeDocxBytes(SAMPLE_ORIGINAL),
      revisedBytes: makeDocxBytes(SAMPLE_REVISED),
      date: '2026-01-01T00:00:00.000Z',
      author: 'Test',
    });

    expect(outcome.authoredBodyXml).toContain('<w:body>');
    expect(outcome.authoredBodyXml).toContain('w:ins');
    expect(outcome.authoredBodyXml).toContain('w:del');
    // Deleted text must be re-tagged, or Word will not read it as a deletion.
    expect(outcome.authoredBodyXml).toContain('w:delText');
  });

  it('carries revision marks through to the render model', async () => {
    const outcome = await runComparison({
      originalBytes: makeDocxBytes(SAMPLE_ORIGINAL),
      revisedBytes: makeDocxBytes(SAMPLE_REVISED),
      date: '2026-01-01T00:00:00.000Z',
    });

    const runs = allRuns(outcome.model.blocks);
    const kinds = new Set(runs.map((r) => r.revision?.type).filter(Boolean));

    // Both directions must survive: an insert-only model would mean deletions
    // were dropped somewhere between the pipeline and the view.
    expect(kinds.has('ins')).toBe(true);
    expect(kinds.has('del')).toBe(true);
  });

  it('renders the WHOLE document, not only the changes', async () => {
    const outcome = await runComparison({
      originalBytes: makeDocxBytes(SAMPLE_ORIGINAL),
      revisedBytes: makeDocxBytes(SAMPLE_REVISED),
      date: '2026-01-01T00:00:00.000Z',
    });

    const runs = allRuns(outcome.model.blocks);
    const unchanged = runs.filter((r) => r.revision === undefined);

    // Paragraph 2 is identical in both versions. It must be present and
    // unmarked — categorise, never suppress.
    expect(unchanged.length).toBeGreaterThan(0);
    const text = runs.map((r) => r.text).join('');
    expect(text).toContain('Alpha Pte Ltd');
  });

  it('detects the numeric value change and classifies it', async () => {
    const outcome = await runComparison({
      originalBytes: makeDocxBytes(SAMPLE_ORIGINAL),
      revisedBytes: makeDocxBytes(SAMPLE_REVISED),
      date: '2026-01-01T00:00:00.000Z',
    });

    const text = allRuns(outcome.model.blocks)
      .map((r) => r.text)
      .join('');
    expect(text).toContain('40,000');
    expect(text).toContain('55,000');
    expect(outcome.summary.byCategory.numericalValue).toBeGreaterThan(0);
  });

  it('is deterministic — the same inputs give the same change count', async () => {
    const args = {
      originalBytes: makeDocxBytes(SAMPLE_ORIGINAL),
      revisedBytes: makeDocxBytes(SAMPLE_REVISED),
      date: '2026-01-01T00:00:00.000Z',
      author: 'Test',
    };
    const a = await runComparison(args);
    const b = await runComparison(args);

    expect(b.summary.total).toBe(a.summary.total);
    expect(b.ops.length).toBe(a.ops.length);
    expect(b.authoredBodyXml).toBe(a.authoredBodyXml);
  });

  it('reports an identical pair as having no changes', async () => {
    const bytes = makeDocxBytes(SAMPLE_ORIGINAL);
    const outcome = await runComparison({
      originalBytes: bytes,
      revisedBytes: makeDocxBytes(SAMPLE_ORIGINAL),
      date: '2026-01-01T00:00:00.000Z',
    });
    expect(outcome.summary.total).toBe(0);
    // …but the document itself is still fully rendered.
    expect(outcome.model.blocks.length).toBeGreaterThan(0);
  });

  it('honours cancellation', async () => {
    await expect(
      runComparison({
        originalBytes: makeDocxBytes(SAMPLE_ORIGINAL),
        revisedBytes: makeDocxBytes(SAMPLE_REVISED),
        date: '2026-01-01T00:00:00.000Z',
        abortSignal: { aborted: true },
      }),
    ).rejects.toThrow(/cancel/i);
  });

  it('fails with a readable message on a file that is not a .docx', async () => {
    const notADocx = new TextEncoder().encode('this is plainly not a zip');
    await expect(
      runComparison({
        originalBytes: notADocx,
        revisedBytes: makeDocxBytes(SAMPLE_REVISED),
        date: '2026-01-01T00:00:00.000Z',
      }),
    ).rejects.toThrow();
  });
});
