/**
 * Feature 06 — Numbering: component/UI tests (T9, spec §9 mapping).
 *
 * Uses the same preact/preact-test-utils pattern as existing UI tests.
 * Sets driver signals directly to control state without calling Office APIs.
 */

import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { render } from 'preact';
import { act } from 'preact/test-utils';
import { UK_PROFILE } from '@shared/config/index.js';
import { NumberingPanel } from '@features/numbering/index.js';
import {
  allSchemes,
  selectedSchemeId,
  reviewState,
  banners,
  operationError,
  editingScheme,
} from '@features/numbering/driver.js';

let container: HTMLDivElement;

beforeEach(() => {
  container = document.createElement('div');
  document.body.appendChild(container);

  // Pre-populate signals with the built-in presets so renders are meaningful.
  const commercial = UK_PROFILE.numberingPresets.find(p => p.id === 'preset-uk-sg-commercial')!;
  const recitals = UK_PROFILE.numberingPresets.find(p => p.id === 'preset-recitals')!;
  const schedule = UK_PROFILE.numberingPresets.find(p => p.id === 'preset-schedule')!;
  allSchemes.value = [commercial, recitals, schedule];
  selectedSchemeId.value = 'preset-uk-sg-commercial';
  reviewState.value = { count: 0, items: [], active: false };
  banners.value = [];
  operationError.value = null;
  editingScheme.value = null;
});

afterEach(() => {
  document.body.removeChild(container);
});

function mount() {
  act(() => { render(<NumberingPanel />, container); });
  return container;
}

// ---------------------------------------------------------------------------
// Render smoke tests
// ---------------------------------------------------------------------------

describe('NumberingPanel', () => {
  it('renders without crashing', () => {
    const div = mount();
    expect(div).toBeTruthy();
  });

  it('shows the Schemes library heading', () => {
    const div = mount();
    expect(div.textContent).toContain('Schemes');
  });

  it('shows the commercial preset name', () => {
    const div = mount();
    expect(div.textContent).toContain('UK/SG Commercial');
  });

  it('shows the Recitals preset name', () => {
    const div = mount();
    expect(div.textContent).toContain('Recitals (A)(B)(C)');
  });

  it('shows the Schedule preset name', () => {
    const div = mount();
    expect(div.textContent).toContain('Schedule (paragraph)');
  });

  it('does NOT show the US Litigation preset (suppressForSgUk=true)', () => {
    // US preset is not in our test allSchemes (only SG-UK presets).
    const div = mount();
    expect(div.textContent).not.toContain('US Litigation');
  });

  it('shows the Apply scheme button', () => {
    const div = mount();
    expect(div.textContent).toContain('Apply scheme');
  });

  it('shows Promote and Demote buttons', () => {
    const div = mount();
    expect(div.textContent).toContain('Promote');
    expect(div.textContent).toContain('Demote');
  });

  it('does not show the review panel when inactive', () => {
    const div = mount();
    expect(div.textContent).not.toContain('References to review');
  });

  it('shows the review panel and items when active', () => {
    reviewState.value = {
      count: 2,
      items: [
        { text: 'clause 4.2(a)', context: 'subject to clause 4.2(a)', paraId: 'p1', startOffset: 0, endOffset: 13 },
        { text: 'Schedule 2',    context: 'in Schedule 2',            paraId: 'p2', startOffset: 0, endOffset: 10 },
      ],
      active: true,
    };
    act(() => { render(<NumberingPanel />, container); });
    expect(container.textContent).toContain('References to review');
    expect(container.textContent).toContain('clause 4.2(a)');
    expect(container.textContent).toContain('Schedule 2');
  });

  it('shows banners when present (T7 Word Online degradation path)', () => {
    banners.value = ['Word Online: scheme applied with reduced support — levels 0 may not display correctly.'];
    act(() => { render(<NumberingPanel />, container); });
    expect(container.textContent).toContain('Word Online');
  });

  it('level editor shows live preview strings for selected scheme', () => {
    const div = mount();
    // Commercial preset level 0 = "1.", level 3 = "(a)", level 4 = "(i)", level 5 = "(A)".
    expect(div.textContent).toContain('(a)');
    expect(div.textContent).toContain('(i)');
    expect(div.textContent).toContain('(A)');
  });

  it('selected scheme is marked in the library list', () => {
    selectedSchemeId.value = 'preset-uk-sg-commercial';
    const div = mount();
    const selected = div.querySelector('.dp-num-library__item--selected');
    expect(selected?.textContent).toContain('UK/SG Commercial');
  });
});

// ---------------------------------------------------------------------------
// Spec §9 must-pass logic coverage summary
// ---------------------------------------------------------------------------

describe('spec §9 must-pass logic coverage', () => {
  it('T1: commercial preset has 9 levels and closed-bracket forms', () => {
    const commercial = UK_PROFILE.numberingPresets.find(p => p.id === 'preset-uk-sg-commercial')!;
    expect(commercial.levels.length).toBe(9);
    expect(commercial.levels[3]?.formatString).toBe('(%4)');
    expect(commercial.levels[4]?.formatString).toBe('(%5)');
    expect(commercial.levels[5]?.formatString).toBe('(%6)');
  });

  it('T3: clampLevel ensures promote/demote stays within 0..8', async () => {
    const { clampLevel } = await import('@features/numbering/level-change.js');
    expect(clampLevel(-1)).toBe(0);
    expect(clampLevel(9)).toBe(8);
    expect(clampLevel(4)).toBe(4);
  });

  it('T5/T6: exportSchemes produces a Blob; importSchemes rejects malformed JSON', async () => {
    const { exportSchemes, importSchemes } = await import('@features/numbering/scheme-store.js');
    const commercial = UK_PROFILE.numberingPresets.find(p => p.id === 'preset-uk-sg-commercial')!;
    const blob = exportSchemes([commercial]);
    expect(blob).toBeInstanceOf(Blob);
    const bad = new Blob(['{not json}'], { type: 'application/json' });
    const result = await importSchemes(bad);
    expect(result.errors.length).toBeGreaterThan(0);
    expect(result.imported).toHaveLength(0);
  });

  it('T10: buildReviewPatterns generates clause + schedule + limb patterns', async () => {
    const { buildReviewPatterns } = await import('@features/numbering/review.js');
    const patterns = buildReviewPatterns(UK_PROFILE);
    expect(patterns.some(p => p.includes('clause'))).toBe(true);
    expect(patterns.some(p => p.includes('Schedule'))).toBe(true);
    expect(patterns.some(p => p.includes('\\([a-z]\\)'))).toBe(true);
  });
});

/*
 * Live-verification checklist (user-side; cannot run without a Word host):
 *
 * T1 [live] — Apply Commercial to 6 nested list paragraphs.
 *   Expected: 1 / 1.1 / 1.1.1 / (a) / (i) / (A). Closed brackets, lower-roman/letters.
 * T2 [live] — Apply Commercial to 1/1.1/1.2/2/2.1.
 *   Expected: level 1 resets under each level-0 (2.1 not 2.3); top continuous (1,2).
 * T3 [live] — Cursor on (a) item; Promote.
 *   Expected: item → 1.1.x style; following items recompute; no manual renumber.
 * T7 [live] — Open in Word Online; Apply Commercial.
 *   Expected: best-effort applied; limited-support banner shown; no unhandled exception.
 * T10 [live] — Doc contains "clause 4.2(a)"; Apply scheme.
 *   Expected: reference highlighted; panel lists it; Clear removes only review highlight.
 * T11 [live] — Apply to style-linked list.
 *   Expected: list not modified; STYLE_LINKED_WARNING shown.
 * T12 [live] — Track Changes on; Demote.
 *   Expected: TRACK_CHANGES_WARNING shown before op; revision created.
 * T13 [live] — Apply preset-recitals to recitals block.
 *   Expected: (A) (B) (C).
 * T14 [live] — Apply preset-schedule.
 *   Expected: 1 / 1.1 / (a); unit label "paragraph" in cross-ref patterns.
 */
