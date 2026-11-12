/**
 * Component tests for the Interactive Lookup panel (feature 04 + 04b).
 * Focus: the definition-list integration and the preview-cap boundary
 * (single/Online preview must stay capped at previewMaxChars; the expanded
 * active card shows the fuller view).
 */

import { describe, it, expect, beforeEach } from 'vitest';
import { render } from 'preact';
import { act } from 'preact/test-utils';
import { LookupPanel } from '@features/interactive-lookup/LookupPanel.js';
import {
  definitionList,
  activeDefinitionKey,
  isDesktopHost,
} from '@features/interactive-lookup/driver.js';
import type { DefinedTerm, LookupResult } from '@features/interactive-lookup/types.js';

const TAIL = 'TAILMARKER';
/** previewText longer than previewMaxChars (600), with a marker past the cap. */
const LONG_PREVIEW = 'x'.repeat(610) + TAIL;

function term(): DefinedTerm {
  return {
    termKey: 'long term',
    display: 'Long Term',
    previewText: LONG_PREVIEW,
    defKind: 'inline',
    sourceParagraphIndex: 0,
    anchor: { paragraphId: 'body-para-0', startOffset: 0, endOffset: 1 },
  };
}

function mount(result: LookupResult) {
  const div = document.createElement('div');
  act(() => { render(<LookupPanel result={result} />, div); });
  return div;
}

describe('LookupPanel — preview cap boundary (04b)', () => {
  beforeEach(() => {
    definitionList.value = [];
    activeDefinitionKey.value = null;
    isDesktopHost.value = false;
  });

  it('single definition preview is capped at previewMaxChars (Online / no list)', () => {
    const div = mount({ type: 'definition', term: term() });
    // Marker sits past char 600 → must be trimmed away in the single preview.
    expect(div.textContent).not.toContain(TAIL);
    expect(div.textContent).toContain('…');
  });

  it('expanded active card shows the fuller view (beyond previewMaxChars)', () => {
    const t = term();
    isDesktopHost.value = true;
    definitionList.value = [{ termKey: t.termKey, display: t.display, result: { type: 'definition', term: t } }];
    activeDefinitionKey.value = t.termKey;

    const div = mount({ type: 'definition', term: t });
    // The active card bypasses the cap → the tail marker is present.
    expect(div.textContent).toContain(TAIL);
  });

  it('Online host (isDesktop false) never renders the accumulating list', () => {
    const t = term();
    isDesktopHost.value = false;
    definitionList.value = [{ termKey: t.termKey, display: t.display, result: { type: 'definition', term: t } }];
    activeDefinitionKey.value = t.termKey;

    const div = mount({ type: 'definition', term: t });
    expect(div.textContent).not.toContain('Definitions (');
    // and the single preview stays capped
    expect(div.textContent).not.toContain(TAIL);
  });

  it('desktop list renders a Definitions heading and collapses non-active entries', () => {
    const a = term();
    const b: DefinedTerm = { ...term(), termKey: 'second term', display: 'Second Term', previewText: 'Short def.' };
    isDesktopHost.value = true;
    definitionList.value = [
      { termKey: b.termKey, display: b.display, result: { type: 'definition', term: b } },
      { termKey: a.termKey, display: a.display, result: { type: 'definition', term: a } },
    ];
    activeDefinitionKey.value = b.termKey; // b active, a collapsed

    const div = mount({ type: 'definition', term: b });
    expect(div.textContent).toContain('Definitions (2)');
    // collapsed entry 'a' shows its display label
    expect(div.textContent).toContain('Long Term');
  });
});
