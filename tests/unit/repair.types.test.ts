import { describe, it, expect } from 'vitest';
import { mapRepairSeverity } from '@features/repair/severity.js';
import type { RepairFinding, RepairReport } from '@features/repair/types.js';

describe('repair types', () => {
  it('RepairFinding satisfies structural shape', () => {
    const f: RepairFinding = {
      id: 'metadata.author',
      category: 'metadata',
      severity: 'confidentiality',
      label: 'Document author present',
      destructive: false,
      requiresConfirm: false,
      selectedByDefault: true,
      action: 'clear',
    };
    expect(f.id).toBe('metadata.author');
    expect(f.result).toBeUndefined();
  });

  it('RepairReport satisfies structural shape', () => {
    const r: RepairReport = {
      documentName: 'test.docx',
      generatedAt: '2026-06-08T09:00:00Z',
      profile: { id: 'sg-court-filing', name: 'SG Court Filing (PD 2021)' },
      phase: 'scan',
      findings: [],
      metrics: { fileSizeBeforeBytes: null, fileSizeAfterBytes: null, rsidCountBefore: null, rsidCountAfter: null },
      errors: [],
    };
    expect(r.phase).toBe('scan');
  });

  it('mapRepairSeverity maps confidentiality → warning', () => {
    expect(mapRepairSeverity('confidentiality')).toBe('warning');
  });

  it('mapRepairSeverity maps stabilisation → info', () => {
    expect(mapRepairSeverity('stabilisation')).toBe('info');
  });

  it('mapRepairSeverity maps info → info', () => {
    expect(mapRepairSeverity('info')).toBe('info');
  });
});
