/**
 * Tests for B7 — Library UI (LibraryPanel + SaveClauseForm).
 *
 * Covers:
 *   - LibraryPanel renders with correct structure (search box, facet chips, results area).
 *   - SaveClauseForm renders required fields and respects disabled state.
 *   - Large-library (>2000) warning rendered when indexSize > threshold.
 *   - Cross-jurisdiction modal renders when insertModalOpen is true.
 *   - DD Month YYYY date display in result cards.
 *   - Save form blocks until required fields filled.
 *
 * Pattern: mount into a fresh div, set signals before rendering, read DOM.
 * No `cleanup` needed — we mount into isolated divs each test.
 * Office JS not needed for rendering tests.
 */

import { describe, it, expect, beforeEach, vi } from 'vitest';
import { render } from 'preact';
import { act } from 'preact/test-utils';
import { LibraryPanel } from '@features/content/library/LibraryPanel.js';
import {
  libraryIndexSize,
  libraryResults,
  insertModalOpen,
  insertModalWarning,
  saveFormOpen,
  libraryBusy,
  saveFormTitle,
  saveFormMetadata,
} from '@features/content/library/driver.js';

// Mock initLibrarySearch so it does not set libraryBusy=true asynchronously
// during component mount. The search init behaviour is tested separately via library.store.test.ts.
vi.mock('@features/content/library/driver.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@features/content/library/driver.js')>();
  return {
    ...actual,
    initLibrarySearch: vi.fn().mockResolvedValue(undefined),
  };
});
import { LARGE_LIBRARY_THRESHOLD } from '@features/content/library/types.js';
import type { SearchResult } from '@features/content/library/search.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function mount(): HTMLDivElement {
  const div = document.createElement('div');
  act(() => { render(<LibraryPanel />, div); });
  return div;
}

function makeResult(id = 'clz-ui-1', jurisdiction = 'SG'): SearchResult {
  return {
    record: {
      id,
      title: 'Test Clause ' + id,
      metadata: {
        jurisdiction: jurisdiction as 'SG' | 'EW' | 'OTHER',
        practiceArea: 'Commercial',
        lastReviewed: '2026-02-14',
        approvedBy: 'J. Tan',
        visibility: 'personal',
      },
      plainText: 'Test clause content.',
      version: 1,
    },
  };
}

// ---------------------------------------------------------------------------
// Reset signals before each test
// ---------------------------------------------------------------------------

beforeEach(() => {
  libraryIndexSize.value = 0;
  libraryResults.value = [];
  insertModalOpen.value = false;
  insertModalWarning.value = '';
  saveFormOpen.value = false;
  libraryBusy.value = false;
  saveFormTitle.value = '';
  saveFormMetadata.value = {};
});

// ---------------------------------------------------------------------------
// Basic render
// ---------------------------------------------------------------------------

describe('LibraryPanel — basic render (B7)', () => {
  it('renders the Clause Library heading', () => {
    const div = mount();
    expect(div.textContent).toContain('Clause Library');
  });

  it('renders the search input', () => {
    const div = mount();
    expect(div.querySelector('#dp-lib-search')).not.toBeNull();
  });

  it('renders facet chips for SG, E&W, Other, Personal, Shared, Stale only', () => {
    const div = mount();
    const text = div.textContent ?? '';
    expect(text).toContain('Singapore');
    expect(text).toContain('England & Wales');
    expect(text).toContain('Other');
    expect(text).toContain('Personal');
    expect(text).toContain('Shared');
    expect(text).toContain('Stale only');
  });

  it('renders "Check clauses in document" button', () => {
    const div = mount();
    expect(div.textContent).toContain('Check clauses in document');
  });

  it('renders "Save selection as clause" button', () => {
    const div = mount();
    expect(div.textContent).toContain('Save selection as clause');
  });
});

// ---------------------------------------------------------------------------
// Empty state
// ---------------------------------------------------------------------------

describe('LibraryPanel — empty state', () => {
  it('shows empty-library message when no entries and no results', () => {
    libraryIndexSize.value = 0;
    libraryResults.value = [];
    const div = mount();
    expect(div.textContent).toContain('No clauses in the library yet');
  });

  it('shows no-match message when index non-empty but results empty', () => {
    libraryIndexSize.value = 5;
    libraryResults.value = [];
    const div = mount();
    expect(div.textContent).toContain('No clauses match');
  });
});

// ---------------------------------------------------------------------------
// Result card date display (DD Month YYYY per spec §B.6)
// ---------------------------------------------------------------------------

describe('LibraryPanel — result card date display (B7)', () => {
  it('displays lastReviewed as DD Month YYYY not ISO format', () => {
    libraryResults.value = [makeResult('clz-date-test', 'SG')];
    // Override lastReviewed in the record
    libraryResults.value = [
      {
        record: {
          id: 'clz-date',
          title: 'Clause with Date',
          metadata: {
            jurisdiction: 'SG',
            practiceArea: 'Commercial',
            lastReviewed: '2026-02-14',
            approvedBy: 'J. Tan',
            visibility: 'personal',
          },
          plainText: 'Content.',
          version: 1,
        },
      },
    ];

    const div = mount();
    const text = div.textContent ?? '';
    // Should show "14 February 2026", not "2026-02-14"
    expect(text).toContain('14 February 2026');
    expect(text).not.toContain('2026-02-14');
  });
});

// ---------------------------------------------------------------------------
// Jurisdiction chip colour coding
// ---------------------------------------------------------------------------

describe('LibraryPanel — jurisdiction chip (B7)', () => {
  it('renders SG chip for SG clause', () => {
    libraryResults.value = [makeResult('clz-sg', 'SG')];
    const div = mount();
    const chip = div.querySelector('.dp-library__juris-chip--sg');
    expect(chip).not.toBeNull();
    expect(chip?.textContent).toBe('SG');
  });

  it('renders E&W chip for EW clause', () => {
    libraryResults.value = [makeResult('clz-ew', 'EW')];
    const div = mount();
    const chip = div.querySelector('.dp-library__juris-chip--ew');
    expect(chip).not.toBeNull();
  });

  it('renders Other chip for OTHER clause', () => {
    libraryResults.value = [
      {
        record: {
          id: 'clz-other',
          title: 'Other Clause',
          metadata: {
            jurisdiction: 'OTHER',
            jurisdictionOther: 'New York',
            practiceArea: 'Commercial',
            lastReviewed: '2026-01-01',
            approvedBy: 'A',
            visibility: 'personal',
          },
          plainText: 'Content.',
          version: 1,
        },
      },
    ];
    const div = mount();
    const chip = div.querySelector('.dp-library__juris-chip--other');
    expect(chip).not.toBeNull();
  });
});

// ---------------------------------------------------------------------------
// Large-library warning (B-T7)
// ---------------------------------------------------------------------------

describe('LibraryPanel — large-library warning (B-T7)', () => {
  it('shows large-library warning when index size > 2000', () => {
    libraryIndexSize.value = LARGE_LIBRARY_THRESHOLD + 1;
    const div = mount();
    expect(div.textContent).toContain('more than');
    expect(div.textContent?.toLowerCase()).toContain('sharepoint');
  });

  it('does NOT show large-library warning when index size <= 2000', () => {
    libraryIndexSize.value = LARGE_LIBRARY_THRESHOLD;
    const div = mount();
    expect(div.textContent).not.toContain('more than 2000');
  });
});

// ---------------------------------------------------------------------------
// Cross-jurisdiction modal (B-T3: non-suppressible warning)
// ---------------------------------------------------------------------------

describe('LibraryPanel — cross-jurisdiction modal (B-T3)', () => {
  it('renders the blocking modal when insertModalOpen is true', () => {
    insertModalOpen.value = true;
    insertModalWarning.value =
      'This clause is governed by Singapore. The current document is governed by England & Wales. Insert anyway?';
    const div = mount();

    const modal = div.querySelector('[role="dialog"]');
    expect(modal).not.toBeNull();
    expect(modal?.textContent).toContain('Singapore');
    expect(modal?.textContent).toContain('England & Wales');
  });

  it('modal has both "Insert anyway" and "Cancel" buttons', () => {
    insertModalOpen.value = true;
    insertModalWarning.value = 'Warning text';
    const div = mount();

    const buttons = Array.from(div.querySelectorAll('[role="dialog"] button'));
    const buttonText = buttons.map(b => b.textContent?.trim());
    expect(buttonText).toContain('Insert anyway');
    expect(buttonText).toContain('Cancel');
  });

  it('modal is not rendered when insertModalOpen is false', () => {
    insertModalOpen.value = false;
    const div = mount();
    expect(div.querySelector('[role="dialog"]')).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// Save form overlay
// ---------------------------------------------------------------------------

describe('LibraryPanel — save form overlay (B7)', () => {
  it('renders SaveClauseForm when saveFormOpen is true', () => {
    saveFormOpen.value = true;
    const div = mount();
    expect(div.querySelector('[aria-label="Save clause to library"]')).not.toBeNull();
  });

  it('does NOT render the search panel when save form is open', () => {
    saveFormOpen.value = true;
    const div = mount();
    // No search input in save form mode
    expect(div.querySelector('#dp-lib-search')).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// SaveClauseForm — required-field blocking
// ---------------------------------------------------------------------------

describe('SaveClauseForm — required field enforcement (B7, B-T1)', () => {
  it('renders required-field inputs', () => {
    saveFormOpen.value = true;
    const div = mount();

    expect(div.querySelector('#dp-lib-title')).not.toBeNull();
    expect(div.querySelector('#dp-lib-jurisdiction')).not.toBeNull();
    expect(div.querySelector('#dp-lib-practice-area')).not.toBeNull();
    expect(div.querySelector('#dp-lib-last-reviewed')).not.toBeNull();
    expect(div.querySelector('#dp-lib-approved-by')).not.toBeNull();
  });

  it('Save Clause button is disabled by default (nothing filled)', () => {
    saveFormOpen.value = true;
    saveFormTitle.value = '';
    saveFormMetadata.value = {};
    const div = mount();
    const submitBtn = Array.from(div.querySelectorAll('button[type="submit"]'))
      .find(b => b.textContent?.includes('Save Clause'));
    expect(submitBtn?.hasAttribute('disabled')).toBe(true);
  });

  it('shared visibility radio is disabled when no backend configured', () => {
    saveFormOpen.value = true;
    const div = mount();
    const sharedRadio = Array.from(div.querySelectorAll('input[type="radio"]'))
      .find(r => (r as HTMLInputElement).value === 'shared') as HTMLInputElement | undefined;
    expect(sharedRadio?.disabled).toBe(true);
  });
});
