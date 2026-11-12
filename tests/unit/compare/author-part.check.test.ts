/**
 * A part whose authored blocks do not reconstruct is not authored
 * (ADR-0047; ADR-0008).
 *
 * No input found trips the body's author inside a part — the yardstick is a
 * backstop, as it is for the body — so the guard is driven by making the
 * yardstick report a loss: the part must then come back unauthored
 * (`blocksXml: null`, the caller keeps the original part) with its ops still
 * listed.
 */
import { describe, it, expect, vi, afterEach } from 'vitest';

let reportLoss = false;
vi.mock('@features/compare/output/reconstruction.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@features/compare/output/reconstruction.js')>();
  return {
    ...actual,
    reconstruct: (...args: Parameters<typeof actual.reconstruct>) => {
      const report = actual.reconstruct(...args);
      return reportLoss ? { ...report, reject: { ...report.reject, textDiffers: [0] } } : report;
    },
  };
});

const { analyseAndAuthorPart } = await import('@features/compare/output/author-part.js');
const { makeAuthorContext } = await import('@features/compare/output/index.js');
const { getCompareConfig, getActiveProfile } = await import('@shared/config/index.js');

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const footer = (text: string): string =>
  `<w:ftr ${W}><w:p><w:r><w:t>${text}</w:t></w:r></w:p><w:p><w:r><w:t>Page</w:t></w:r></w:p></w:ftr>`;

afterEach(() => { reportLoss = false; });

const author = () => analyseAndAuthorPart({
  partId: 'footer:s0:default',
  originalXml: footer('Draft for discussion'),
  revisedXml: footer('Execution version'),
  config: getCompareConfig(getActiveProfile()),
  author: 'Compare', date: '2026-10-26T00:00:00Z',
  ctx: makeAuthorContext('Compare', '2026-10-26T00:00:00Z', 1),
});

describe('the per-part self-check', () => {
  it('authors a part that reconstructs', async () => {
    const result = await author();
    expect(result.changed).toBe(true);
    expect(result.blocksXml).toMatch(/<w:ins /);
  });

  it('refuses a part the yardstick says loses text — not authored, still listed', async () => {
    reportLoss = true;
    const result = await author();
    expect(result.blocksXml).toBeNull();
    expect(result.ops.filter((op) => op.type !== 'equal').length).toBeGreaterThan(0);
  });
});
