/**
 * Feature 08 — Compare §5.5 · the ONE place mark order can damage a document.
 * ADR-0043 (mark order) × ADR-0032 (merge/split seams).
 *
 * A `PARA_SEAM` is a PARAGRAPH MARK, not a word. `splitAtSeams`
 * (`output/author.ts`) cuts the word-diff stream at each seam and turns the
 * seam-bearing entry's op into the revision on that mark. So a run that crosses
 * a seam in the mark-order swap crosses a **paragraph boundary**: the runs
 * either side land in different `<w:p>` elements than ADR-0032 authored them
 * into. `insertBeforeDelete` therefore skips any window containing a seam.
 *
 * **A green suite proves nothing here**, which is why this file exists. The plan
 * measured the seam-adjacent shape firing 4× across the suite with zero
 * coverage, and ablating the barrier entirely changed no test outcome.
 *
 * Measured with the barrier condition deleted: the **three presentation
 * assertions below fail** — the armed-trap check, the MERGE and the SPLIT. The
 * fourth, `accept-all and reject-all still round-trip through the seam`, **passes
 * and cannot fail**: it is the never-suppress half, and the barrier does not
 * protect it — swapping across a seam moves text between the emitted paragraphs
 * but still reduces to the same two documents. That is the finding, not a gap:
 * the damage the barrier prevents is a silent rewrite of ADR-0032's output, not a
 * lost change. (This comment claimed "every assertion below" until review
 * measured it.) Both transcripts are in the plan's Evidence.
 *
 * Two facts worth not re-deriving, both measured while building this:
 *   - The seam-bearing `del` run abutting an `ins` run is EASY to produce —
 *     every ordinary merge and split makes one.
 *   - The corruption that would swap a merge for a split (the mark revisions
 *     trading places) needs seams on BOTH sides of one window, and
 *     `detectMergesAndSplits → tryPair` only ever produces 2-original→1-revised
 *     or 1→2 — never 2→2 — so an op can carry at most ONE seam and that
 *     corruption is unreachable. What IS reachable, and what these tests pin, is
 *     text moving between the emitted paragraphs.
 */
import { describe, it, expect } from 'vitest';

import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import type { CompareConfig } from '@shared/config/index.js';
import { runComparePipeline } from '@features/compare/pipeline.js';
import { extractTextRecords } from '@features/compare/io/text/extract-text-records.js';
import { insertBeforeDelete, wordDiff, PARA_SEAM } from '@features/compare/diff/index.js';

const CONFIG: CompareConfig = getCompareConfig(SG_PROFILE);
const AUTHOR = 'Mutatis Compare';
const DATE = '2026-09-17T00:00:00.000Z';

// Long enough to clear MERGE_MIN_TOKENS on both halves.
const P1 = 'The Vendor shall give notice of any Claim within 30 days.';
const P2 = 'The notice shall specify the nature of the Claim in reasonable detail.';
const JOINED =
  'The Vendor shall give notice of any Claim within 30 days, and the notice shall ' +
  'specify the nature of the Claim in reasonable detail.';

async function run(original: string, revised: string): Promise<string> {
  const result = await runComparePipeline({
    originalParas: extractTextRecords(original),
    revisedParas: extractTextRecords(revised),
    config: CONFIG,
    author: AUTHOR,
    date: DATE,
  });
  return result.authoredBodyXml;
}

/** The emitted `<w:p>` elements, in order. */
function paragraphs(xml: string): string[] {
  return [...xml.matchAll(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g)].map((m) => m[0]);
}

/** The paragraph's runs, without its `<w:pPr>` — whose revision is self-closing. */
function body(p: string): string {
  return p.replace(/<w:pPr>[\s\S]*?<\/w:pPr>/, '');
}

/** Visible text of each revision kind inside one `<w:p>`'s runs. */
function marked(p: string, kind: 'ins' | 'del'): string {
  const tag = kind === 'ins' ? 'w:ins' : 'w:del';
  return [...body(p).matchAll(new RegExp(`<${tag}\\b[^>]*>([\\s\\S]*?)</${tag}>`, 'g'))]
    .map((m) => [...m[1]!.matchAll(/<w:(?:t|delText)\b[^>]*>([\s\S]*?)<\/w:(?:t|delText)>/g)]
      .map((t) => t[1]!)
      .join(''))
    .join('');
}

/** The revision on this paragraph's own MARK, if any (`<w:pPr><w:rPr><w:ins|del/>`). */
function markRevision(p: string): 'ins' | 'del' | null {
  const pPr = /<w:pPr>([\s\S]*?)<\/w:pPr>/.exec(p);
  if (!pPr) return null;
  const rPr = /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(pPr[1]!);
  if (!rPr) return null;
  if (/<w:ins\b/.test(rPr[1]!)) return 'ins';
  if (/<w:del\b/.test(rPr[1]!)) return 'del';
  return null;
}

describe('mark order × merge/split seams (§5.5, ADR-0043 × ADR-0032)', () => {
  it('the trap is armed: a merge really does put a seam-bearing del beside an ins', () => {
    // If this ever stops holding, the two tests below are vacuous and the
    // barrier is untested — which is the exact failure mode this file exists
    // for. Assert the shape before asserting what the barrier does to it.
    const diffs = wordDiff(`${P1}${PARA_SEAM}${P2}`, JOINED, CONFIG);
    const at = diffs.findIndex((d) => d.op === 'del' && d.text.includes(PARA_SEAM));
    expect(at, 'no seam-bearing del run in this fixture').toBeGreaterThanOrEqual(0);
    // …and an ins run abuts it, so the swap window would contain the seam.
    let end = at;
    while (end < diffs.length && diffs[end]!.op === 'del') end += 1;
    expect(diffs[end]?.op, 'no ins run abuts the seam-bearing del').toBe('ins');

    // The barrier is what stops it: with the seam, nothing moves.
    expect(insertBeforeDelete(diffs.map((d) => ({ ...d })))).toEqual(diffs);
    // Without the seam, the identical shape IS swapped — so the refusal above
    // is the barrier firing, not the pass declining for some other reason.
    const noSeam = diffs.map((d) => ({ ...d, text: d.text.split(PARA_SEAM).join(' ') }));
    expect(insertBeforeDelete(noSeam.map((d) => ({ ...d })))).not.toEqual(noSeam);
  });

  it('a MERGE authors the same two paragraphs, marks and boundary as before the pass', async () => {
    const xml = await run(`${P1}\n${P2}`, JOINED);
    const ps = paragraphs(xml);
    expect(ps).toHaveLength(2);

    // The break is struck on the FIRST paragraph's mark — this is what makes the
    // merge rejectable (ADR-0032). It must not become an insertion.
    expect(markRevision(ps[0]!)).toBe('del');
    expect(markRevision(ps[1]!)).toBe(null);

    // …and the inserted joining words stay in the paragraph ADR-0032 put them
    // in. Without the barrier they cross the seam into the first paragraph.
    expect(marked(ps[0]!, 'ins')).toBe('');
    expect(marked(ps[1]!, 'ins')).toContain(', and the');
    expect(marked(ps[0]!, 'del')).toContain('.');
  });

  it('a SPLIT authors the same two paragraphs, marks and boundary as before the pass', async () => {
    const xml = await run(JOINED, `${P1}\n${P2}`);
    const ps = paragraphs(xml);
    expect(ps).toHaveLength(2);

    // The new break is INSERTED on the first paragraph's mark — rejecting it
    // restores the single paragraph.
    expect(markRevision(ps[0]!)).toBe('ins');
    expect(markRevision(ps[1]!)).toBe(null);

    // The struck joining words stay where ADR-0032 put them: in the FIRST
    // paragraph, beside the inserted full stop. Without the barrier the swap
    // carries them across the seam into the second.
    expect(marked(ps[0]!, 'del')).toContain(', and the');
    expect(marked(ps[0]!, 'ins')).toBe('.');
    expect(marked(ps[1]!, 'del')).toBe('');
    expect(marked(ps[1]!, 'ins')).toBe('The');
  });

  it('accept-all and reject-all still round-trip through the seam', async () => {
    // The never-suppress half: whatever the barrier does to presentation, the
    // two documents the redline reduces to are the two inputs (ADR-0008).
    const xml = await run(`${P1}\n${P2}`, JOINED);
    const ps = paragraphs(xml);

    // `body` strips `<w:pPr>`, whose own revision is self-closing and would
    // otherwise swallow every run up to the next `</w:del>`.
    const accept = (p: string): string =>
      [
        ...body(p)
          .replace(/<w:del\b[^>]*>[\s\S]*?<\/w:del>/g, '')
          .matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g),
      ]
        .map((m) => m[1]!)
        .join('');
    const reject = (p: string): string =>
      [
        ...body(p)
          .replace(/<w:ins\b[^>]*>[\s\S]*?<\/w:ins>/g, '')
          .matchAll(/<w:(?:t|delText)\b[^>]*>([\s\S]*?)<\/w:(?:t|delText)>/g),
      ]
        .map((m) => m[1]!)
        .join('');

    // Accept: the mark on ¶1 is struck, so the two join into the revised text.
    expect(ps.map(accept).join('')).toBe(JOINED);
    // Reject: the mark survives, so the two originals stand apart.
    expect(ps.map(reject)).toEqual([P1, P2]);
  });
});
