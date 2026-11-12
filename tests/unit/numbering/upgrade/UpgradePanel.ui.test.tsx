/**
 * Feature 06 upgrade (ADR-0023) — UpgradePanel component/UI test (T7).
 * Sets driver signals directly to control state without Office APIs.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { render } from 'preact';
import { act } from 'preact/test-utils';
import { UpgradePanel } from '@features/numbering/upgrade/UpgradePanel.js';
import {
  trackingUnavailable,
  whitespaceOverride,
  sequenceAnomalies,
  lastRemovalReport,
  removalQueue,
} from '@features/numbering/upgrade/upgrade-driver.js';

let container: HTMLDivElement;

beforeEach(() => {
  localStorage.clear();
  container = document.createElement('div');
  document.body.appendChild(container);
});

afterEach(() => {
  document.body.removeChild(container);
  trackingUnavailable.value = false;
  sequenceAnomalies.value = [];
  lastRemovalReport.value = null;
  removalQueue.value = [];
});

/** Mount and flush the async activateUpgrade() effect so later signal sets win. */
async function mount() {
  await act(async () => {
    render(<UpgradePanel />, container);
    await Promise.resolve();
  });
}

describe('UpgradePanel', () => {
  it('renders both step buttons', async () => {
    await mount();
    const text = container.textContent ?? '';
    expect(text).toContain('Remove numbers (tracked)');
    expect(text).toContain('Number selected paragraphs');
  });

  it('disables Step 1 and shows a banner when tracking is unavailable', async () => {
    await mount();
    act(() => { trackingUnavailable.value = true; });

    const banner = container.querySelector('.dp-num-upgrade__banner');
    expect(banner).not.toBeNull();

    const step1 = container.querySelector('.dp-num-upgrade__step .dp-num-btn--primary') as HTMLButtonElement;
    expect(step1.disabled).toBe(true);
  });

  it('reflects the whitespace override signal in the select control', async () => {
    await mount();
    act(() => { whitespaceOverride.value = 'keep-one-tab'; });

    const select = container.querySelector('#dp-num-ws-override') as HTMLSelectElement;
    expect(select.value).toBe('keep-one-tab');
  });

  it('renders sequence anomalies with both readings', async () => {
    await mount();
    act(() => {
      sequenceAnomalies.value = [{
        kind: 'duplicate', index: 1, token: '1.5', previousToken: '1.5',
        readings: { typo: 'should be 1.6', reserved: 'reserved deliberately' },
      }];
    });

    const text = container.textContent ?? '';
    expect(text).toContain('Sequence anomalies');
    expect(text).toContain('should be 1.6');
    expect(text).toContain('reserved deliberately');
  });
});
