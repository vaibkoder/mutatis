import { describe, it, expect } from 'vitest';
import { toMarkdown } from '@features/ai-export/formats/markdown.js';
import type { AiExportEnvelope } from '@features/ai-export/types.js';

const env: AiExportEnvelope = {
  _mutatis: '1.0',
  generated: '2026-06-09T14:32:00Z',
  document: 'agreement.docx',
  jurisdiction: 'SG',
  checks_run: ['defined-terms', 'factual-consistency'],
  blind_spots: [
    'Compare was not run; change set absent.',
    'Headers, footers, and text boxes are not scanned.',
  ],
  defined_terms: { glossary: [{ term: 'Buyer', definition: 'means ACME.', defined_at: 'Clause 1.1', source: 'prose' }] },
  deal_facts: {
    dates: [{ raw: '6 June 2026', iso: '2026-06-06', day_of_week: 'Saturday', role: 'execution', flagged_weekend: true, clause: 'Clause 2' }],
    amounts: [{ raw: '£200', value: 200, currency: 'GBP', words_raw: 'Two Hundred Pounds', mismatch: 'value', clause: 'Clause 3' }],
    placeholders: [{ raw: '[TBC]', category: 'keyword', clause: 'Clause 4' }],
  },
};

describe('toMarkdown', () => {
  const md = toMarkdown(env);

  it('renders a heading per contributed section', () => {
    expect(md).toContain('## Defined terms');
    expect(md).toContain('## Deal facts');
  });

  it('puts blind spots in a > ⚠ blockquote (F6)', () => {
    expect(md).toContain('> ⚠');
    expect(md).toContain('> - Compare was not run; change set absent.');
  });

  it('renders amounts and dates as Markdown tables', () => {
    expect(md).toContain('| Raw | Value | Currency | Words | Mismatch | Clause |');
    expect(md).toContain('| Raw | ISO | Day | Role | Weekend | Clause |');
    expect(md).toContain('| £200 | 200 | GBP |');
  });

  it('carries provenance and never leaks an internal paraId field', () => {
    expect(md).toContain('**Document:** agreement.docx');
    expect(md).not.toContain('paragraphId');
    expect(md).not.toContain('paraId');
  });
});
