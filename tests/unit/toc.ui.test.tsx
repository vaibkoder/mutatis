/**
 * tests/unit/toc.ui.test.tsx
 *
 * Component tests for A6: TocPanel.tsx.
 *
 * Acceptance checks for the TOC feature:
 *  - Profile selector rendered and wired to driver.
 *  - Insert TOC buttons wired to driver actions.
 *  - Scoped-TOC button present and wired.
 *  - Generate Index of Schedules button present.
 *  - Refresh this/all buttons present.
 *  - F9 banner shown when showF9Banner signal is true (A-T5).
 *  - Registered-TOC list with go-to rendered when registry is non-empty.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { render } from 'preact';
import { act } from 'preact/test-utils';

// ---------------------------------------------------------------------------
// Mock: the driver module.
// vi.mock is hoisted so the factory is self-contained; we import the module
// after the mock is installed and mutate signal values via the imported refs.
// ---------------------------------------------------------------------------

vi.mock('@features/content/toc/driver.js', () => ({
  activeProfile: { value: { id: 'house-default', label: 'Firm Standard TOC', switches: {}, styleLevelMap: [], tabLeader: 'dots' } },
  availableProfiles: { value: [{ id: 'house-default', label: 'Firm Standard TOC' }] },
  registry: { value: [] as Array<{ fieldId: string; scopeLabel: string; bookmark?: string }> },
  busy: { value: false },
  error: { value: null as string | null },
  showF9Banner: { value: false },
  banners: { value: [] as string[] },
  activateToc: vi.fn(async () => {}),
  selectProfile: vi.fn(async () => {}),
  runInsertTocAtCursor: vi.fn(async () => {}),
  runInsertTocAtTop: vi.fn(async () => {}),
  runInsertScopedToc: vi.fn(async () => {}),
  runGenerateScheduleIndex: vi.fn(async () => {}),
  runRefreshOne: vi.fn(async () => {}),
  runRefreshAll: vi.fn(async () => {}),
  runNavigateToToc: vi.fn(async () => {}),
}));

// ---------------------------------------------------------------------------
// Import under test + the mocked driver (to mutate signals in tests).
// ---------------------------------------------------------------------------

import { TocPanel } from '@features/content/toc/TocPanel.js';
import * as driver from '@features/content/toc/driver.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function renderPanel(): HTMLDivElement {
  const container = document.createElement('div');
  document.body.appendChild(container);
  act(() => { render(<TocPanel />, container); });
  return container;
}

type AnyProfile = { id: string; label: string; switches: Record<string,unknown>; styleLevelMap: unknown[]; tabLeader: string };
type RegistryEntry = { fieldId: string; scopeLabel: string; bookmark?: string };

function resetDriver() {
  (driver.activeProfile as unknown as { value: AnyProfile }).value = {
    id: 'house-default', label: 'Firm Standard TOC', switches: {}, styleLevelMap: [], tabLeader: 'dots',
  };
  (driver.availableProfiles as unknown as { value: AnyProfile[] }).value = [
    { id: 'house-default', label: 'Firm Standard TOC', switches: {}, styleLevelMap: [], tabLeader: 'dots' },
  ];
  (driver.registry as { value: RegistryEntry[] }).value = [];
  (driver.busy as { value: boolean }).value = false;
  (driver.error as { value: string | null }).value = null;
  (driver.showF9Banner as { value: boolean }).value = false;
  (driver.banners as { value: string[] }).value = [];
  vi.clearAllMocks();
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('TocPanel — structure', () => {
  beforeEach(resetDriver);

  it('renders without crashing', () => {
    const container = renderPanel();
    expect(container.querySelector('.dp-toc')).not.toBeNull();
  });

  it('calls activateToc on mount', () => {
    renderPanel();
    expect(driver.activateToc).toHaveBeenCalledTimes(1);
  });

  it('renders the profile selector', () => {
    const container = renderPanel();
    const select = container.querySelector('#dp-toc-profile') as HTMLSelectElement;
    expect(select).not.toBeNull();
    expect(select.options.length).toBeGreaterThan(0);
    expect(select.options[0]!.textContent).toBe('Firm Standard TOC');
  });
});

// ---------------------------------------------------------------------------
// F9 banner (A-T5)
// ---------------------------------------------------------------------------

describe('TocPanel — F9 banner (A-T5)', () => {
  beforeEach(resetDriver);

  it('does NOT show F9 banner when showF9Banner is false', () => {
    (driver.showF9Banner as { value: boolean }).value = false;
    const container = renderPanel();
    expect(container.querySelector('.dp-toc__f9-banner')).toBeNull();
  });

  it('shows F9 banner when showF9Banner is true', () => {
    (driver.showF9Banner as { value: boolean }).value = true;
    const container = renderPanel();
    const banner = container.querySelector('.dp-toc__f9-banner');
    expect(banner).not.toBeNull();
    expect(banner?.textContent?.toLowerCase()).toContain('f9');
  });
});

// ---------------------------------------------------------------------------
// Error display
// ---------------------------------------------------------------------------

describe('TocPanel — error display', () => {
  beforeEach(resetDriver);

  it('does not show error area when error is null', () => {
    (driver.error as { value: string | null }).value = null;
    const container = renderPanel();
    expect(container.querySelector('.dp-toc__error')).toBeNull();
  });

  it('shows error message when error is set', () => {
    (driver.error as { value: string | null }).value = 'Something went wrong.';
    const container = renderPanel();
    expect(container.querySelector('.dp-toc__error')?.textContent).toContain('Something went wrong.');
  });
});

// ---------------------------------------------------------------------------
// Insert TOC buttons (A-T1)
// ---------------------------------------------------------------------------

describe('TocPanel — Insert TOC buttons', () => {
  beforeEach(resetDriver);

  it('has an "Insert at cursor" button', () => {
    const container = renderPanel();
    const btn = Array.from(container.querySelectorAll('button'))
      .find(b => b.textContent?.toLowerCase().includes('insert at cursor'));
    expect(btn).not.toBeUndefined();
  });

  it('has an "Insert at top of document" button', () => {
    const container = renderPanel();
    const btn = Array.from(container.querySelectorAll('button'))
      .find(b => b.textContent?.toLowerCase().includes('insert at top'));
    expect(btn).not.toBeUndefined();
  });

  it('clicking "Insert at cursor" calls runInsertTocAtCursor', () => {
    const container = renderPanel();
    const btn = Array.from(container.querySelectorAll('button'))
      .find(b => b.textContent?.toLowerCase().includes('insert at cursor')) as HTMLButtonElement;
    act(() => { btn.click(); });
    expect(driver.runInsertTocAtCursor).toHaveBeenCalledTimes(1);
  });

  it('clicking "Insert at top of document" calls runInsertTocAtTop', () => {
    const container = renderPanel();
    const btn = Array.from(container.querySelectorAll('button'))
      .find(b => b.textContent?.toLowerCase().includes('insert at top')) as HTMLButtonElement;
    act(() => { btn.click(); });
    expect(driver.runInsertTocAtTop).toHaveBeenCalledTimes(1);
  });

  it('buttons are disabled when busy', () => {
    (driver.busy as { value: boolean }).value = true;
    const container = renderPanel();
    const btn = Array.from(container.querySelectorAll('button'))
      .find(b => b.textContent?.toLowerCase().includes('insert at cursor')) as HTMLButtonElement;
    expect(btn.disabled).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// Scoped TOC (A-T3)
// ---------------------------------------------------------------------------

describe('TocPanel — scoped TOC button', () => {
  beforeEach(resetDriver);

  it('renders the scoped TOC section', () => {
    const container = renderPanel();
    const text = container.textContent?.toLowerCase() ?? '';
    expect(text).toContain('scoped toc');
  });

  it('clicking "Make scoped TOC from selection" calls runInsertScopedToc', () => {
    const container = renderPanel();
    const btn = Array.from(container.querySelectorAll('button'))
      .find(b => b.textContent?.toLowerCase().includes('make scoped toc')) as HTMLButtonElement;
    act(() => { btn.click(); });
    expect(driver.runInsertScopedToc).toHaveBeenCalledTimes(1);
  });
});

// ---------------------------------------------------------------------------
// Index of Schedules (A-T4)
// ---------------------------------------------------------------------------

describe('TocPanel — Index of Schedules', () => {
  beforeEach(resetDriver);

  it('renders a "Generate / Regenerate Index of Schedules" button', () => {
    const container = renderPanel();
    const btn = Array.from(container.querySelectorAll('button'))
      .find(b => b.textContent?.toLowerCase().includes('index of schedules'));
    expect(btn).not.toBeUndefined();
  });

  it('clicking the button calls runGenerateScheduleIndex', () => {
    const container = renderPanel();
    const btn = Array.from(container.querySelectorAll('button'))
      .find(b => b.textContent?.toLowerCase().includes('index of schedules')) as HTMLButtonElement;
    act(() => { btn.click(); });
    expect(driver.runGenerateScheduleIndex).toHaveBeenCalledTimes(1);
  });
});

// ---------------------------------------------------------------------------
// Refresh (A-T2)
// ---------------------------------------------------------------------------

describe('TocPanel — Refresh all', () => {
  beforeEach(resetDriver);

  it('renders a "Refresh all TOCs" button', () => {
    const container = renderPanel();
    const btn = Array.from(container.querySelectorAll('button'))
      .find(b => b.textContent?.toLowerCase().includes('refresh all'));
    expect(btn).not.toBeUndefined();
  });

  it('clicking "Refresh all TOCs" calls runRefreshAll', () => {
    const container = renderPanel();
    const btn = Array.from(container.querySelectorAll('button'))
      .find(b => b.textContent?.toLowerCase().includes('refresh all')) as HTMLButtonElement;
    act(() => { btn.click(); });
    expect(driver.runRefreshAll).toHaveBeenCalledTimes(1);
  });
});

// ---------------------------------------------------------------------------
// Registered TOC list with go-to
// ---------------------------------------------------------------------------

describe('TocPanel — registered TOC list', () => {
  beforeEach(resetDriver);

  it('does not render the TOC list when registry is empty', () => {
    (driver.registry as { value: RegistryEntry[] }).value = [];
    const container = renderPanel();
    expect(container.querySelector('.dp-toc__list')).toBeNull();
  });

  it('renders each registry entry', () => {
    (driver.registry as { value: RegistryEntry[] }).value = [
      { fieldId: 'f1', scopeLabel: 'Main TOC' },
      { fieldId: 'f2', scopeLabel: 'Part 1', bookmark: 'TOC_SCOPE_part1_1' },
    ];
    const container = renderPanel();
    const list = container.querySelector('.dp-toc__list');
    expect(list).not.toBeNull();
    expect(list?.querySelectorAll('li').length).toBe(2);
  });

  it('shows "(scoped)" badge for entries with a bookmark', () => {
    (driver.registry as { value: RegistryEntry[] }).value = [
      { fieldId: 'f2', scopeLabel: 'Part 1', bookmark: 'TOC_SCOPE_part1_1' },
    ];
    const container = renderPanel();
    expect(container.querySelector('.dp-toc__list-badge')?.textContent).toContain('scoped');
  });

  it('clicking "Go to" calls runNavigateToToc with the fieldId', () => {
    (driver.registry as { value: RegistryEntry[] }).value = [
      { fieldId: 'f1', scopeLabel: 'Main TOC' },
    ];
    const container = renderPanel();
    const goToBtn = Array.from(container.querySelectorAll('button'))
      .find(b => b.textContent?.toLowerCase() === 'go to') as HTMLButtonElement;
    act(() => { goToBtn.click(); });
    expect(driver.runNavigateToToc).toHaveBeenCalledWith('f1');
  });

  it('clicking "Refresh" on a list item calls runRefreshOne with the fieldId', () => {
    (driver.registry as { value: RegistryEntry[] }).value = [
      { fieldId: 'f1', scopeLabel: 'Main TOC' },
    ];
    const container = renderPanel();
    const refreshBtns = Array.from(container.querySelectorAll('button'))
      .filter(b => b.textContent?.toLowerCase() === 'refresh') as HTMLButtonElement[];
    act(() => { refreshBtns[0]!.click(); });
    expect(driver.runRefreshOne).toHaveBeenCalledWith('f1');
  });

  it('"Refresh" button in list is disabled when fieldId is empty (F9-path entry)', () => {
    (driver.registry as { value: RegistryEntry[] }).value = [
      { fieldId: '', scopeLabel: 'Fallback TOC' },
    ];
    const container = renderPanel();
    const refreshBtns = Array.from(container.querySelectorAll('button'))
      .filter(b => b.textContent?.toLowerCase() === 'refresh') as HTMLButtonElement[];
    expect(refreshBtns[0]!.disabled).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// Busy state
// ---------------------------------------------------------------------------

describe('TocPanel — busy state', () => {
  beforeEach(resetDriver);

  it('shows "Working…" overlay when busy', () => {
    (driver.busy as { value: boolean }).value = true;
    const container = renderPanel();
    expect(container.querySelector('.dp-toc__busy')).not.toBeNull();
  });

  it('does not show busy overlay when not busy', () => {
    (driver.busy as { value: boolean }).value = false;
    const container = renderPanel();
    expect(container.querySelector('.dp-toc__busy')).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// No US-isms in labels
// ---------------------------------------------------------------------------

describe('TocPanel — SG/UK labels', () => {
  beforeEach(resetDriver);

  it('labels use en-GB terminology (no "Exhibit", no "Table of Authorities")', () => {
    const container = renderPanel();
    const text = container.textContent ?? '';
    expect(text).not.toContain('Exhibit');
    expect(text).not.toContain('Table of Authorities');
    // "Index of Schedules" is the correct SG/UK label.
    expect(text).toContain('Index of Schedules');
  });
});
