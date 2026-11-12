/**
 * Feature 05 — Document Repair: component/UI tests (T14).
 *
 * Controls all state by writing directly to the driver signals —
 * no Office API, no mocks of the async orchestrators.
 */

import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { render } from 'preact';
import { act } from 'preact/test-utils';
import { RepairPanel } from '@features/repair/index.js';
import {
  repairProfile,
  repairScanReport,
  repairApplyReport,
  repairBusy,
  repairError,
  repairChecked,
  repairConfirmedSectionBreaks,
} from '@features/repair/driver.js';
import { SG_COURT_FILING, UK_COURT_FILING, FIRM_COMMERCIAL } from '@features/repair/profiles.js';
import type { RepairReport, RepairFinding } from '@features/repair/types.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

let container: HTMLDivElement;

beforeEach(() => {
  container = document.createElement('div');
  document.body.appendChild(container);

  // Reset all driver signals to a clean pre-scan state.
  repairProfile.value               = SG_COURT_FILING;
  repairScanReport.value            = null;
  repairApplyReport.value           = null;
  repairBusy.value                  = false;
  repairError.value                 = null;
  repairChecked.value               = new Set();
  repairConfirmedSectionBreaks.value = new Set();
});

afterEach(() => {
  document.body.removeChild(container);
});

function mount() {
  act(() => { render(<RepairPanel />, container); });
  return container;
}

/** Build a minimal RepairReport for test use. */
function makeReport(
  phase: 'scan' | 'apply',
  findings: RepairFinding[],
): RepairReport {
  return {
    documentName: 'Test.docx',
    generatedAt: '2026-01-01T00:00:00.000Z',
    profile: { id: SG_COURT_FILING.id, name: SG_COURT_FILING.name },
    phase,
    findings,
    metrics: {
      fileSizeBeforeBytes: 1024,
      fileSizeAfterBytes:  null,
      rsidCountBefore:     10,
      rsidCountAfter:      null,
    },
    errors: [],
  };
}

/** Minimal non-destructive finding. */
function makeFinding(
  id: string,
  severity: 'confidentiality' | 'stabilisation' | 'info',
  override: Partial<RepairFinding> = {},
): RepairFinding {
  return {
    id,
    category: 'metadata',
    severity,
    label: `Finding ${id}`,
    destructive: false,
    requiresConfirm: false,
    selectedByDefault: true,
    action: 'clearBuiltinProps',
    result: null,
    ...override,
  };
}

// ---------------------------------------------------------------------------
// 1. Trust banner
// ---------------------------------------------------------------------------

describe('Trust banner', () => {
  it('is always visible and contains "on your device"', () => {
    const div = mount();
    const banner = div.querySelector('.dp-repair__trust-banner');
    expect(banner).toBeTruthy();
    expect(banner!.textContent).toContain('on your device');
  });

  it('is still visible after a scan report is set', () => {
    repairScanReport.value = makeReport('scan', [makeFinding('f1', 'info')]);
    const div = mount();
    expect(div.querySelector('.dp-repair__trust-banner')).toBeTruthy();
  });
});

// ---------------------------------------------------------------------------
// 2. Profile selector
// ---------------------------------------------------------------------------

describe('Profile selector', () => {
  it('renders an option for every profile id', () => {
    const div = mount();
    const select = div.querySelector('select.dp-repair__profile-select') as HTMLSelectElement;
    expect(select).toBeTruthy();
    const optionValues = Array.from(select.options).map(o => o.value);
    expect(optionValues).toContain(SG_COURT_FILING.id);
    expect(optionValues).toContain(UK_COURT_FILING.id);
    expect(optionValues).toContain(FIRM_COMMERCIAL.id);
  });

  it('shows all three shipped profile names', () => {
    const div = mount();
    const text = div.querySelector('select.dp-repair__profile-select')!.textContent;
    expect(text).toContain('SG Court Filing');
    expect(text).toContain('UK Court Filing');
    expect(text).toContain('Firm Commercial');
  });

  it('pre-selects the current profile signal value', () => {
    repairProfile.value = UK_COURT_FILING;
    const div = mount();
    const select = div.querySelector('select.dp-repair__profile-select') as HTMLSelectElement;
    expect(select.value).toBe(UK_COURT_FILING.id);
  });

  it('is disabled while busy', () => {
    repairBusy.value = true;
    const div = mount();
    const select = div.querySelector('select.dp-repair__profile-select') as HTMLSelectElement;
    expect(select.disabled).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// 3. Scan button (pre-scan state)
// ---------------------------------------------------------------------------

describe('Scan button', () => {
  it('renders a scan button when no report is set', () => {
    const div = mount();
    expect(div.querySelector('.dp-repair__scan-btn')).toBeTruthy();
  });

  it('shows "Scan document" when not busy', () => {
    const div = mount();
    expect(div.querySelector('.dp-repair__scan-btn')!.textContent).toContain('Scan document');
  });

  it('shows "Scanning…" when busy with no report', () => {
    repairBusy.value = true;
    const div = mount();
    expect(div.querySelector('.dp-repair__scan-btn')!.textContent).toContain('Scanning');
  });

  it('disappears once a scan report is present', () => {
    repairScanReport.value = makeReport('scan', []);
    const div = mount();
    expect(div.querySelector('.dp-repair__scan-btn')).toBeNull();
  });

  it('disappears once an apply report is present', () => {
    repairApplyReport.value = makeReport('apply', []);
    const div = mount();
    expect(div.querySelector('.dp-repair__scan-btn')).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// 4. Findings checklist renders with correct groups
// ---------------------------------------------------------------------------

describe('Findings checklist', () => {
  beforeEach(() => {
    const findings: RepairFinding[] = [
      makeFinding('c1', 'confidentiality', { label: 'Built-in properties', count: 3 }),
      makeFinding('c2', 'confidentiality', { label: 'Comments', count: 1 }),
      makeFinding('s1', 'stabilisation',   { label: 'RSID cruft' }),
      makeFinding('i1', 'info',            { label: 'Style anomaly' }),
    ];
    repairScanReport.value = makeReport('scan', findings);
    repairChecked.value    = new Set(['c1', 'c2', 's1', 'i1']);
  });

  it('renders a Confidentiality group heading', () => {
    const div = mount();
    expect(div.textContent).toContain('Confidentiality');
  });

  it('renders a Stabilisation group heading', () => {
    const div = mount();
    expect(div.textContent).toContain('Stabilisation');
  });

  it('renders an Info group heading', () => {
    const div = mount();
    expect(div.textContent).toContain('Info');
  });

  it('renders each finding label', () => {
    const div = mount();
    expect(div.textContent).toContain('Built-in properties');
    expect(div.textContent).toContain('Comments');
    expect(div.textContent).toContain('RSID cruft');
    expect(div.textContent).toContain('Style anomaly');
  });

  it('renders the count in parentheses when present', () => {
    const div = mount();
    expect(div.textContent).toContain('(3)');
    expect(div.textContent).toContain('(1)');
  });

  it('does not render confidentiality group when there are none', () => {
    repairScanReport.value = makeReport('scan', [
      makeFinding('s1', 'stabilisation', { label: 'Only stabilisation' }),
    ]);
    const div = mount();
    expect(div.textContent).not.toContain('Confidentiality');
    expect(div.textContent).toContain('Stabilisation');
  });
});

// ---------------------------------------------------------------------------
// 5. Default selection and destructive findings
// ---------------------------------------------------------------------------

describe('Default selection and destructive findings', () => {
  it('non-destructive selectedByDefault finding has a checked checkbox', () => {
    const finding = makeFinding('f1', 'confidentiality', { selectedByDefault: true, destructive: false });
    repairScanReport.value = makeReport('scan', [finding]);
    repairChecked.value    = new Set(['f1']);
    const div = mount();
    const checkbox = div.querySelector('input[type="checkbox"]') as HTMLInputElement;
    expect(checkbox.checked).toBe(true);
  });

  it('non-destructive unchecked finding has an unchecked checkbox', () => {
    const finding = makeFinding('f1', 'confidentiality', { selectedByDefault: false, destructive: false });
    repairScanReport.value = makeReport('scan', [finding]);
    repairChecked.value    = new Set(); // not in checked set
    const div = mount();
    const checkbox = div.querySelector('input[type="checkbox"]') as HTMLInputElement;
    expect(checkbox.checked).toBe(false);
  });

  it('destructive finding has a disabled, unchecked checkbox (confirmed=false)', () => {
    const finding = makeFinding('d1', 'stabilisation', {
      destructive: true,
      requiresConfirm: true,
      label: 'Remove section breaks',
    });
    repairScanReport.value = makeReport('scan', [finding]);
    repairChecked.value    = new Set(); // destructive never auto-checked
    repairConfirmedSectionBreaks.value = new Set();
    const div = mount();
    const checkbox = div.querySelector('input[type="checkbox"]') as HTMLInputElement;
    expect(checkbox.disabled).toBe(true);
    expect(checkbox.checked).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// 6. Destructive finding — 🔒 badge and confirm button
// ---------------------------------------------------------------------------

describe('Destructive finding UI', () => {
  beforeEach(() => {
    const finding: RepairFinding = makeFinding('d1', 'stabilisation', {
      label: 'Remove section breaks',
      destructive: true,
      requiresConfirm: true,
    });
    repairScanReport.value = makeReport('scan', [finding]);
    repairChecked.value    = new Set();
    repairConfirmedSectionBreaks.value = new Set();
  });

  it('shows the 🔒 lock indicator', () => {
    const div = mount();
    const lock = div.querySelector('.dp-repair__lock');
    expect(lock).toBeTruthy();
    expect(lock!.getAttribute('aria-label')).toContain('confirmation');
  });

  it('shows the confirm button when not yet confirmed', () => {
    const div = mount();
    expect(div.querySelector('.dp-repair__confirm-btn')).toBeTruthy();
  });

  it('confirm button has accessible text', () => {
    const div = mount();
    const btn = div.querySelector('.dp-repair__confirm-btn');
    expect(btn!.textContent).toContain('Confirm');
  });

  it('confirm button does not appear when already confirmed', () => {
    repairConfirmedSectionBreaks.value = new Set(['d1']);
    const div = mount();
    expect(div.querySelector('.dp-repair__confirm-btn')).toBeNull();
  });

  it('confirm button says "Confirm (changes orientation)" for protected findings', () => {
    const prot: RepairFinding = makeFinding('p1', 'stabilisation', {
      label: 'Landscape section break',
      destructive: true,
      requiresConfirm: true,
      protected: true,
    });
    repairScanReport.value = makeReport('scan', [prot]);
    const div = mount();
    const btn = div.querySelector('.dp-repair__confirm-btn');
    expect(btn!.textContent).toContain('changes orientation');
  });

  it('finding row has the --destructive modifier class', () => {
    const div = mount();
    expect(div.querySelector('.dp-repair__finding--destructive')).toBeTruthy();
  });
});

// ---------------------------------------------------------------------------
// 7. Apply-phase result summary
// ---------------------------------------------------------------------------

describe('Apply-phase result summary', () => {
  it('shows "Repair complete." and count of done findings', () => {
    const findings: RepairFinding[] = [
      makeFinding('a1', 'confidentiality', { result: 'done' }),
      makeFinding('a2', 'confidentiality', { result: 'done' }),
      makeFinding('a3', 'stabilisation',   { result: 'skipped' }),
    ];
    repairApplyReport.value = makeReport('apply', findings);
    const div = mount();
    const summary = div.querySelector('.dp-repair__apply-summary');
    expect(summary).toBeTruthy();
    expect(summary!.textContent).toContain('Repair complete');
    expect(summary!.textContent).toContain('2');
  });

  it('shows "Repair complete." even when 0 actions applied', () => {
    repairApplyReport.value = makeReport('apply', [
      makeFinding('a1', 'info', { result: 'skipped' }),
    ]);
    const div = mount();
    expect(div.querySelector('.dp-repair__apply-summary')!.textContent).toContain('0');
  });

  it('shows result badge in apply phase', () => {
    repairApplyReport.value = makeReport('apply', [
      makeFinding('a1', 'confidentiality', { result: 'done', label: 'Props' }),
    ]);
    repairChecked.value = new Set(['a1']);
    const div = mount();
    expect(div.textContent).toContain('[done]');
  });

  it('shows a Scan again button after apply', () => {
    repairApplyReport.value = makeReport('apply', []);
    const div = mount();
    expect(div.querySelector('.dp-repair__rescan-btn')).toBeTruthy();
    expect(div.querySelector('.dp-repair__rescan-btn')!.textContent).toContain('Scan again');
  });

  it('does not show Apply selected button after apply', () => {
    repairApplyReport.value = makeReport('apply', []);
    const div = mount();
    expect(div.querySelector('.dp-repair__apply-btn')).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// 8. Error state
// ---------------------------------------------------------------------------

describe('Error state', () => {
  it('shows the error message with role="alert"', () => {
    repairError.value = 'Office API unavailable';
    const div = mount();
    const el = div.querySelector('[role="alert"]');
    expect(el).toBeTruthy();
    expect(el!.textContent).toContain('Office API unavailable');
  });

  it('applies the --error class to the error element', () => {
    repairError.value = 'Something went wrong';
    const div = mount();
    expect(div.querySelector('.dp-repair__error')).toBeTruthy();
  });

  it('does not show an error element when repairError is null', () => {
    repairError.value = null;
    const div = mount();
    expect(div.querySelector('.dp-repair__error')).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// 9. Apply button (scan phase only)
// ---------------------------------------------------------------------------

describe('Apply button', () => {
  it('shows Apply selected when a scan report is present and not busy', () => {
    repairScanReport.value = makeReport('scan', [makeFinding('f1', 'info')]);
    repairBusy.value       = false;
    const div = mount();
    expect(div.querySelector('.dp-repair__apply-btn')).toBeTruthy();
  });

  it('hides Apply selected while busy', () => {
    repairScanReport.value = makeReport('scan', [makeFinding('f1', 'info')]);
    repairBusy.value       = true;
    const div = mount();
    expect(div.querySelector('.dp-repair__apply-btn')).toBeNull();
  });

  it('hides Apply selected when only an apply report is present', () => {
    repairApplyReport.value = makeReport('apply', []);
    repairBusy.value        = false;
    const div = mount();
    expect(div.querySelector('.dp-repair__apply-btn')).toBeNull();
  });
});
