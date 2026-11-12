import { describe, it, expect, beforeEach } from 'vitest';
import { render } from 'preact';
import { Shell } from '@taskpane/shell/index.js';
import { addFinding, clearFindings } from '@shared/findings/index.js';
import type { Finding } from '@shared/findings/index.js';

function makeFinding(id: string, category: string): Finding {
  return {
    id,
    category,
    severity: 'error',
    title: `Issue ${id}`,
    location: { label: `Clause ${id}` },
  };
}

describe('Shell', () => {
  let container: HTMLDivElement;

  beforeEach(() => {
    clearFindings();
    container = document.createElement('div');
    document.body.appendChild(container);
  });

  it('renders the tab bar without error', () => {
    render(<Shell />, container);
    const tabs = container.querySelectorAll('[role="tab"]');
    expect(tabs.length).toBe(7);
    expect(container.textContent).toContain('Proof');
    expect(container.textContent).toContain('Settings');
  });

  it('switches to the Settings panel when its tab is clicked', () => {
    render(<Shell />, container);
    const settingsTab = Array.from(container.querySelectorAll('[role="tab"]'))
      .find(el => el.textContent?.includes('Settings')) as HTMLButtonElement;
    settingsTab.click();
    render(<Shell />, container);
    expect(container.textContent).toContain('Active jurisdiction');
  });

  it('reflects findings count in a tab badge', () => {
    addFinding(makeFinding('1', 'repair/metadata'));
    addFinding(makeFinding('2', 'repair/comments'));
    render(<Shell />, container);
    const repairTab = Array.from(container.querySelectorAll('[role="tab"]'))
      .find(el => el.textContent?.includes('Repair'));
    expect(repairTab?.textContent).toContain('2');
  });
});
