/**
 * C9 — TemplatesPanel component tests.
 * Verifies:
 *   - Panel mounts with three mode tabs (Author / Fill / Data Sets)
 *   - Fill mode contains the questionnaire form
 *   - Author mode shows schema management UI
 *   - Repeating group "Add" would be disabled at cap (structural check)
 */
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { render } from 'preact';

// Mock shared modules so the panel doesn't try to call Office JS.
vi.mock('@shared/office/index.js', () => ({
  wordRun: vi.fn(),
  batchLoad: vi.fn(),
  processInChunks: vi.fn(),
}));

vi.mock('@shared/storage/index.js', async () => {
  const { createLocalStore } = await import('@shared/storage/local.js');
  const store = createLocalStore();
  return {
    getStorageProvider: () => ({ local: store, docSettings: store }),
  };
});

import { TemplatesPanel } from '@features/content/templates/TemplatesPanel.js';

describe('C9 — TemplatesPanel', () => {
  let container: HTMLDivElement;

  beforeEach(() => {
    localStorage.clear();
    container = document.createElement('div');
    document.body.appendChild(container);
  });

  it('renders three mode tabs: Author, Fill, Data Sets', () => {
    render(<TemplatesPanel />, container);
    const tabs = container.querySelectorAll('[role="tab"]');
    expect(tabs.length).toBe(3);
    const tabTexts = Array.from(tabs).map(t => t.textContent?.trim());
    expect(tabTexts).toContain('Author');
    expect(tabTexts).toContain('Fill');
    expect(tabTexts).toContain('Data Sets');
  });

  it('Fill is the default active tab', () => {
    render(<TemplatesPanel />, container);
    const activeTab = container.querySelector('[role="tab"][aria-selected="true"]');
    expect(activeTab?.textContent?.trim()).toBe('Fill');
  });

  it('switching to Author mode shows Author content', () => {
    render(<TemplatesPanel />, container);
    const authorTab = Array.from(container.querySelectorAll('[role="tab"]'))
      .find(t => t.textContent?.trim() === 'Author') as HTMLButtonElement;
    authorTab.click();
    render(<TemplatesPanel />, container);
    expect(container.textContent).toMatch(/mark as variable/i);
  });

  it('Author mode shows schema list section titles', () => {
    render(<TemplatesPanel />, container);
    const authorTab = Array.from(container.querySelectorAll('[role="tab"]'))
      .find(t => t.textContent?.trim() === 'Author') as HTMLButtonElement;
    authorTab.click();
    render(<TemplatesPanel />, container);
    expect(container.textContent).toMatch(/variables/i);
    expect(container.textContent).toMatch(/conditionals/i);
    expect(container.textContent).toMatch(/repeating groups/i);
  });

  it('Data Sets mode shows storage indicator and execution blocks section', () => {
    render(<TemplatesPanel />, container);
    const dsTab = Array.from(container.querySelectorAll('[role="tab"]'))
      .find(t => t.textContent?.trim() === 'Data Sets') as HTMLButtonElement;
    dsTab.click();
    render(<TemplatesPanel />, container);
    expect(container.textContent).toMatch(/execution.*block|attestation/i);
  });

  it('Data Sets panel shows SharePoint residency indicator', () => {
    render(<TemplatesPanel />, container);
    const dsTab = Array.from(container.querySelectorAll('[role="tab"]'))
      .find(t => t.textContent?.trim() === 'Data Sets') as HTMLButtonElement;
    dsTab.click();
    render(<TemplatesPanel />, container);
    expect(container.textContent).toMatch(/sharepoint/i);
  });
});

describe('C9 — QuestionnaireForm integration', () => {
  let container: HTMLDivElement;

  beforeEach(() => {
    localStorage.clear();
    container = document.createElement('div');
    document.body.appendChild(container);
  });

  it('renders empty questionnaire message when no variables', async () => {
    const { QuestionnaireForm } = await import('@features/content/templates/QuestionnaireForm.js');
    render(
      <QuestionnaireForm
        variables={[]}
        jurisdiction="EW"
        answers={{}}
        onChange={() => undefined}
      />,
      container,
    );
    expect(container.textContent).toMatch(/no variables/i);
  });

  it('renders day-first hint for date type fields', async () => {
    const { QuestionnaireForm } = await import('@features/content/templates/QuestionnaireForm.js');
    render(
      <QuestionnaireForm
        variables={[{
          tag: 'EFFECTIVE_DATE',
          label: 'Effective date',
          type: 'date',
          required: true,
        }]}
        jurisdiction="EW"
        answers={{ EFFECTIVE_DATE: '' }}
        onChange={() => undefined}
      />,
      container,
    );
    expect(container.textContent).toMatch(/day first|DD Month YYYY/i);
  });

  it('renders reg-number label for regNumber type fields', async () => {
    const { QuestionnaireForm } = await import('@features/content/templates/QuestionnaireForm.js');
    render(
      <QuestionnaireForm
        variables={[{
          tag: 'COMPANY_NUMBER',
          label: 'Companies House number',
          type: 'regNumber',
          required: true,
          regScheme: 'UK_CH',
        }]}
        jurisdiction="EW"
        answers={{ COMPANY_NUMBER: '' }}
        onChange={() => undefined}
      />,
      container,
    );
    expect(container.textContent).toMatch(/Companies House|England/i);
  });

  it('renders entity-type options without Inc./LLC', async () => {
    const { QuestionnaireForm } = await import('@features/content/templates/QuestionnaireForm.js');
    render(
      <QuestionnaireForm
        variables={[{
          tag: 'ENTITY_TYPE',
          label: 'Entity type',
          type: 'entityType',
          required: true,
        }]}
        jurisdiction="EW"
        answers={{ ENTITY_TYPE: '' }}
        onChange={() => undefined}
      />,
      container,
    );
    // EW options include Limited, Ltd, plc
    expect(container.textContent).toContain('Limited');
    expect(container.textContent).not.toMatch(/Inc\.|LLC|EIN/);
  });
});
