/**
 * the side map trusts `textLeaves` only when it names one
 * source for every character `collectText` reads (`text-leaves.test.ts` pins
 * that they agree). Here `textLeaves` is made to disagree by one trailing entry
 * that changes no character, so a rebuild from the misaligned leaves would
 * still reconstruct and the dry-run check could not catch it: only the guard
 * can refuse the source rebuild, and the author falls back to the text rebuild.
 */
import { describe, it, expect, vi } from 'vitest';

import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import { runPipeline } from '@features/compare/pipeline.js';

import { inlineParts, run, para } from '../../fixtures/inline-objects-docx.js';

vi.mock('@features/compare/preprocess/inline-units.js', async (importOriginal) => {
  const real = await importOriginal<typeof import('@features/compare/preprocess/inline-units.js')>();
  return { ...real, textLeaves: (...args: Parameters<typeof real.textLeaves>) => [...real.textLeaves(...args), null] };
});

const OPEN = para(run('This Agreement is made on the date set out above.'));
const bold = (t: string): string => `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${t}</w:t></w:r>`;

describe('leaves that disagree with the text are not trusted', () => {
  it('an edited clause with a bold word is rebuilt from text', async () => {
    const result = await runPipeline({
      originalParts: inlineParts(OPEN, para(run('The ') + bold('Seller') + run(' shall deliver the Shares.'))),
      revisedParts: inlineParts(OPEN, para(run('The ') + bold('Seller') + run(' shall transfer the Shares.'))),
      config: getCompareConfig(SG_PROFILE), author: 'Mutatis', date: '2026-10-26T00:00:00Z',
    });
    const xml = (result.authoredBodyXml.match(/<w:p[\s>][\s\S]*?<\/w:p>/g) ?? [])[1] ?? '';
    expect(xml).toContain('transfer');
    expect(xml).not.toContain('<w:b/>');
  });
});
