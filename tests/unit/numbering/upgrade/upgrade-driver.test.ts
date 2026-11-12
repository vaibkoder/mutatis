/**
 * Feature 06 upgrade (ADR-0023) — driver signal wiring (T6).
 * Queue populates from feature-02 findings; the tracking-disabled state raises
 * the banner signal; sequence anomalies surface into a signal (constraint 4).
 */
import { describe, it, expect, vi, beforeEach } from 'vitest';
import type { RemovalReport } from '@features/numbering/upgrade/removal.js';

// Drive the engines through mocks so the driver runs without a live Word host.
const removalResult = vi.fn<() => RemovalReport>();
vi.mock('@features/numbering/upgrade/removal.js', () => ({
  removeLeadingNumbers: () => Promise.resolve(removalResult()),
}));
vi.mock('@shared/office/index.js', () => ({
  wordRun: (fn: (ctx: unknown) => unknown) => Promise.resolve(fn({})),
}));

import {
  activateUpgrade,
  runRemoval,
  removalQueue,
  trackingUnavailable,
  sequenceAnomalies,
  lastRemovalReport,
} from '@features/numbering/upgrade/upgrade-driver.js';
import { addFindings, clearFindings } from '@shared/findings/index.js';
import type { Finding } from '@shared/findings/types.js';

function manuallyTyped(id: string): Finding {
  return {
    id,
    category: 'static-numbering/manually-typed',
    severity: 'warning',
    title: `Static number ${id}`,
    location: { label: `Clause ${id}` },
  };
}

function report(over: Partial<RemovalReport> = {}): RemovalReport {
  return { disabled: false, total: 0, deleted: 0, skipped: 0, notFound: 0, tokens: [], ...over };
}

beforeEach(() => {
  clearFindings();
  localStorage.clear();
  removalResult.mockReset();
});

describe('upgrade-driver — queue', () => {
  it('populates the queue from manually-typed findings on activate', async () => {
    addFindings([manuallyTyped('a'), manuallyTyped('b')]);
    await activateUpgrade();
    expect(removalQueue.value).toHaveLength(2);
  });

  it('ignores findings of other categories', async () => {
    addFindings([{ ...manuallyTyped('a'), category: 'static-numbering/other' }]);
    await activateUpgrade();
    expect(removalQueue.value).toHaveLength(0);
  });
});

describe('upgrade-driver — removal wiring', () => {
  it('raises the tracking-disabled banner when the host refuses tracking', async () => {
    removalResult.mockReturnValue(report({ disabled: true }));
    await runRemoval();
    expect(trackingUnavailable.value).toBe(true);
    expect(lastRemovalReport.value?.disabled).toBe(true);
  });

  it('surfaces sequence anomalies among the cleared tokens (both readings)', async () => {
    removalResult.mockReturnValue(report({ deleted: 2, total: 2, tokens: ['1.5', '1.5'] }));
    await runRemoval();
    expect(sequenceAnomalies.value).toHaveLength(1);
    expect(sequenceAnomalies.value[0]!.kind).toBe('duplicate');
    expect(sequenceAnomalies.value[0]!.readings.typo).toBeTruthy();
    expect(sequenceAnomalies.value[0]!.readings.reserved).toBeTruthy();
    expect(trackingUnavailable.value).toBe(false);
  });

  it('clears anomalies when the cleared tokens form a clean sequence', async () => {
    removalResult.mockReturnValue(report({ deleted: 2, total: 2, tokens: ['1.1', '1.2'] }));
    await runRemoval();
    expect(sequenceAnomalies.value).toEqual([]);
  });
});
