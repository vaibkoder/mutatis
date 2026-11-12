import { describe, it, expect } from 'vitest';
import { toJson } from '@features/ai-export/formats/json.js';
import type { AiExportEnvelope } from '@features/ai-export/types.js';

const env: AiExportEnvelope = {
  _mutatis: '1.0',
  generated: '2026-06-09T14:32:00Z',
  document: 'agreement.docx',
  jurisdiction: 'SG',
  checks_run: ['defined-terms'],
  blind_spots: ['Compare was not run; change set absent.'],
  defined_terms: { glossary: [{ term: 'Buyer', definition: 'means ACME.', defined_at: 'Clause 1.1', source: 'prose' }] },
};

describe('toJson', () => {
  it('produces valid parseable JSON', () => {
    const parsed = JSON.parse(toJson(env));
    expect(parsed.document).toBe('agreement.docx');
    expect(parsed.defined_terms.glossary[0].term).toBe('Buyer');
  });

  it('carries the _mutatis sentinel and a non-empty blind_spots', () => {
    const parsed = JSON.parse(toJson(env));
    expect(parsed._mutatis).toBe('1.0');
    expect(Array.isArray(parsed.blind_spots)).toBe(true);
    expect(parsed.blind_spots.length).toBeGreaterThan(0);
  });

  it('indents with two spaces', () => {
    expect(toJson(env)).toContain('\n  "generated"');
  });
});
