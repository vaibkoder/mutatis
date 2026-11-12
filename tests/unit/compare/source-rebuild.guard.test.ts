/**
 * the rebuild of an edited paragraph from its source runs is
 * kept only when its own dry run reconstructs both documents exactly
 * (`reconstructsExactly`, which runs the yardstick's `isClean`). This file makes
 * that check fail every time and asserts the author falls back to the text
 * rebuild — the guard's ablation: without the check, the source rebuild would
 * be written regardless and these tests go red.
 */
import { describe, it, expect, vi } from 'vitest';

import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import { runPipeline } from '@features/compare/pipeline.js';

import { inlineParts, run, para } from '../../fixtures/inline-objects-docx.js';

vi.mock('@features/compare/output/reconstruction.js', async (importOriginal) => ({
  ...(await importOriginal<typeof import('@features/compare/output/reconstruction.js')>()),
  isClean: () => false,
}));

const OPEN = para(run('This Agreement is made on the date set out above.'));
const bold = (t: string): string => `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${t}</w:t></w:r>`;

async function authored(o: string, r: string): Promise<string> {
  const result = await runPipeline({
    originalParts: inlineParts(OPEN, para(o)), revisedParts: inlineParts(OPEN, para(r)),
    config: getCompareConfig(SG_PROFILE), author: 'Mutatis', date: '2026-10-26T00:00:00Z',
  });
  return (result.authoredBodyXml.match(/<w:p[\s>][\s\S]*?<\/w:p>/g) ?? [])[1] ?? '';
}

describe('a source rebuild that does not reconstruct is not written', () => {
  it('an edited clause with a bold word falls back to the text rebuild', async () => {
    const xml = await authored(
      run('The ') + bold('Seller') + run(' shall deliver the Shares.'),
      run('The ') + bold('Seller') + run(' shall transfer the Shares.'),
    );
    expect(xml).not.toContain('<w:b/>');
    expect(xml).toContain('transfer');
  });

  it('the text rebuild marks a changed manual number too: the new inserted, the old struck (task 4)', async () => {
    const xml = await authored(run('1.1 The Vendor grants an exclusive licence.'), run('1.2 The Vendor grants a non-exclusive licence.'));
    expect(xml).toMatch(/<w:ins [^>]*><w:r>(?:<w:rPr>[\s\S]*?<\/w:rPr>)?<w:t[^>]*>1\.2 <\/w:t><\/w:r><\/w:ins>/);
    expect(xml).toMatch(/<w:delText[^>]*>1\.1 <\/w:delText>/);
  });

  it('a merge of two clauses with a bold word falls back to the text rebuild (ADR-0032)', async () => {
    const result = await runPipeline({
      originalParts: inlineParts(OPEN,
        para(run('The ') + bold('Purchaser') + run(' shall notify the Vendor of any Third Party Claim.')),
        para(run('The notice shall specify the nature of the claim in reasonable detail.'))),
      revisedParts: inlineParts(OPEN,
        para(run('The ') + bold('Purchaser') + run(' shall notify the Vendor of any Third Party Claim, and the notice shall specify the nature of the claim in reasonable detail.'))),
      config: getCompareConfig(SG_PROFILE), author: 'Mutatis', date: '2026-10-26T00:00:00Z',
    });
    expect(result.authoredBodyXml).toContain('Purchaser');
    expect(result.authoredBodyXml).not.toContain('<w:b/>');
  });
});
