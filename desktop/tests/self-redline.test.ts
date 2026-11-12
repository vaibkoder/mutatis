import { describe, it, expect } from 'vitest';
import { applyEdits } from '../src/renderer/src/edits/apply-edits.js';
import { runComparison } from '../src/renderer/src/compare/run-comparison.js';
import { makeDocxBytes } from './helpers/make-docx.js';

/**
 * The rule: every AI edit lands as a tracked change the user reviews.
 *
 * This test is that rule, end to end — apply an edit, run the SAME comparison
 * engine between the document and its edited copy, and assert that what comes
 * back is native revision markup attributed to the assistant. No separate
 * "AI edit" path exists to get this wrong, and this is what proves it.
 */

const AT = '2026-01-01T00:00:00.000Z';

const DOC = [
  'This Agreement is made on 1 January 2026.',
  'The Purchase Price is GBP 1,000,000, payable on Completion.',
  'The Seller shall indemnify the Buyer against all losses.',
  'Notices must be given in writing to the address in Schedule 1.',
];

async function editAndRedline(edits: { find: string; replace: string; reason?: string }[]) {
  const target = makeDocxBytes(DOC);
  const { bytes, applied, refused } = applyEdits(target, edits);
  const outcome = await runComparison({
    originalBytes: target,
    revisedBytes: bytes,
    originalName: 'target.docx',
    revisedName: "target.docx (assistant's edits)",
    author: 'Mutatis assistant',
    date: AT,
  });
  return { outcome, applied, refused };
}

describe('an assistant edit becomes a tracked change', () => {
  it('produces native w:ins / w:del attributed to the assistant', async () => {
    const { outcome, applied } = await editAndRedline([
      { find: 'GBP 1,000,000', replace: 'GBP 1,250,000', reason: 'agreed uplift' },
    ]);

    expect(applied).toHaveLength(1);
    expect(outcome.authoredBodyXml).toContain('<w:ins');
    expect(outcome.authoredBodyXml).toContain('<w:del');
    expect(outcome.authoredBodyXml).toContain('Mutatis assistant');
  });

  it('the redline shows the change, and only the change', async () => {
    const { outcome } = await editAndRedline([
      { find: 'GBP 1,000,000', replace: 'GBP 1,250,000' },
    ]);
    expect(outcome.summary.total).toBe(1);
  });

  it('gives the assistant\'s turn a change set it can talk about', async () => {
    const { outcome } = await editAndRedline([
      { find: 'in writing', replace: 'in writing or by email' },
    ]);
    expect(outcome.context.changes).toHaveLength(1);
    const change = outcome.context.changes[0];
    expect(change?.original).toContain('in writing');
    expect(change?.revised).toContain('or by email');
    expect(change?.id).toBe(outcome.ops.find((op) => op.type !== 'equal')?.id);
  });

  it('applies several edits in one pass and redlines them together', async () => {
    const { outcome, applied } = await editAndRedline([
      { find: 'GBP 1,000,000', replace: 'GBP 2,000,000' },
      { find: 'all losses', replace: 'all direct losses' },
    ]);
    expect(applied).toHaveLength(2);
    expect(outcome.summary.total).toBe(2);
  });

  it('a refused edit produces no change at all — not a partial one', async () => {
    const { outcome, applied, refused } = await editAndRedline([
      { find: 'The Purchase Price is GBP 1,000,000 payable on Completion', replace: 'x' },
    ]);
    expect(applied).toEqual([]);
    expect(refused).toHaveLength(1);
    expect(outcome.summary.total).toBe(0);
  });

  it('is deterministic — the same edit yields the same redline', async () => {
    const edits = [{ find: 'all losses', replace: 'all direct losses' }];
    const a = await editAndRedline(edits);
    const b = await editAndRedline(edits);
    expect(a.outcome.authoredBodyXml).toBe(b.outcome.authoredBodyXml);
  });

  it('leaves the untouched paragraphs out of the change set entirely', async () => {
    const { outcome } = await editAndRedline([
      { find: 'all losses', replace: 'all direct losses' },
    ]);
    const touched = outcome.context.changes.map((change) => change.original ?? '').join(' ');
    expect(touched).not.toContain('This Agreement is made');
    expect(touched).not.toContain('Schedule 1');
  });
});
