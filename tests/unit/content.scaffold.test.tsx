import { describe, it, expect, beforeEach } from 'vitest';
import { render } from 'preact';
import { ContentPanel } from '@features/content/index.js';
import { Shell } from '@taskpane/shell/index.js';
import { clearFindings } from '@shared/findings/index.js';

// Asserts the S1 scaffold contract structurally (active sub-tab + which
// tabpanel is mounted), so it stays valid after sub-features A/B/C replace
// their placeholder stubs with real panels.

describe('ContentPanel scaffold', () => {
  let container: HTMLDivElement;

  beforeEach(() => {
    container = document.createElement('div');
    document.body.appendChild(container);
  });

  it('renders three sub-tabs with Contents active by default', () => {
    render(<ContentPanel />, container);
    expect(container.querySelectorAll('[role="tab"]').length).toBe(3);
    expect(
      container.querySelector('#dp-content-subtab-toc')?.getAttribute('aria-selected'),
    ).toBe('true');
    expect(container.querySelector('[role="tabpanel"]')?.id).toBe('dp-content-panel-toc');
  });

  it('switches the rendered sub-panel when another sub-tab is clicked', () => {
    render(<ContentPanel />, container);

    (container.querySelector('#dp-content-subtab-library') as HTMLButtonElement).click();
    render(<ContentPanel />, container);
    expect(
      container.querySelector('#dp-content-subtab-library')?.getAttribute('aria-selected'),
    ).toBe('true');
    expect(container.querySelector('[role="tabpanel"]')?.id).toBe('dp-content-panel-library');

    (container.querySelector('#dp-content-subtab-templates') as HTMLButtonElement).click();
    render(<ContentPanel />, container);
    expect(container.querySelector('[role="tabpanel"]')?.id).toBe('dp-content-panel-templates');

    // Reset the module-level sub-tab signal so sibling tests start on Contents.
    (container.querySelector('#dp-content-subtab-toc') as HTMLButtonElement).click();
    render(<ContentPanel />, container);
  });
});

describe('Shell — Content tab', () => {
  let container: HTMLDivElement;

  beforeEach(() => {
    clearFindings();
    container = document.createElement('div');
    document.body.appendChild(container);
  });

  it('renders the ContentPanel when the Content tab is activated', () => {
    render(<Shell />, container);
    const contentTab = Array.from(container.querySelectorAll('[role="tab"]'))
      .find(el => el.textContent?.includes('Content')) as HTMLButtonElement;
    contentTab.click();
    render(<Shell />, container);
    expect(container.querySelector('.dp-content')).not.toBeNull();
    expect(container.querySelector('#dp-content-subtab-library')).not.toBeNull();
  });
});
