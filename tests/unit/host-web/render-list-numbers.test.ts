/**
 * host-web — renderer seam tests (authored against the 12c stub; updated for the
 * 12d implementation).
 *
 * The rendering behaviour matrix lives in `numbering-render.todo.test.ts`; this
 * file pins the seam's degenerate inputs: empty scheme / empty paragraph list
 * never throw and yield an empty map (the loader's guard makes these unreachable
 * in practice, but the function must stay total for direct callers).
 */
import { describe, it, expect } from 'vitest';
import { renderListNumbers } from '@host-web/render-list-numbers.js';
import { parseNumberingScheme } from '@host-web/numbering-scheme.js';
import { PizZipBundle } from '@shared/ooxml/index.js';

describe('renderListNumbers (seam)', () => {
  it('returns an empty map for an empty scheme (every para unresolvable)', () => {
    const scheme = parseNumberingScheme(new PizZipBundle());
    const out = renderListNumbers(scheme, [{ paraId: 'body-para-0', numId: '1', ilvl: 0 }]);
    expect(out.size).toBe(0);
  });

  it('returns an empty map for no list paragraphs', () => {
    const scheme = parseNumberingScheme(new PizZipBundle());
    expect(renderListNumbers(scheme, []).size).toBe(0);
  });
});
