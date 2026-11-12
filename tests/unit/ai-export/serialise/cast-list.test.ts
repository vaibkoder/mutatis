import { describe, it, expect } from 'vitest';
import { serialiseCastList } from '@features/ai-export/serialise/cast-list.js';
import type { RoleRecord, PartyAnchor } from '@features/party-role/index.js';

const anchor = (paragraphId: string) => ({ paragraphId, startOffset: 0, endOffset: 1 });
const cite = (id: string) => (id === 'pIntro' ? 'Clause 1.2' : 'Clause 1');

function role(p: Partial<RoleRecord> & { key: string; label: string }): RoleRecord {
  return { introduced: false, introSites: [], uses: [], ...p };
}
function pa(entityKey: string, entityText: string, roleKey: string, roleLabel: string): PartyAnchor {
  return { entityKey, entityText, roleKey, roleLabel, anchor: anchor('pIntro') };
}

describe('serialiseCastList', () => {
  it('emits entities, body_uses, and introduced_at for an introduced role', () => {
    const roles = [role({
      key: 'purchaser', label: 'Purchaser', introduced: true,
      introSites: [anchor('pIntro')], uses: [anchor('a'), anchor('b'), anchor('c')],
    })];
    const anchors = [pa('acme', 'ACME Pte Ltd', 'purchaser', 'Purchaser')];
    const out = serialiseCastList({ roles, anchors }, { cite });
    const p = out.parties[0]!;
    expect(p.role).toBe('Purchaser');
    expect(p.entities).toEqual(['ACME Pte Ltd']);
    expect(p.body_uses).toBe(3);
    expect(p.introduced_at).toBe('Clause 1.2');
    expect(p.finding).toBeUndefined();
  });

  it('omits introduced_at and flags used-not-introduced', () => {
    const roles = [role({ key: 'vendor', label: 'Vendor', introduced: false, uses: [anchor('a'), anchor('b')] })];
    const out = serialiseCastList({ roles, anchors: [] }, { cite });
    const p = out.parties[0]!;
    expect('introduced_at' in p).toBe(false);
    expect(p.entities).toEqual([]);
    expect(p.finding).toBe('used-not-introduced');
  });

  it('flags inconsistent-label when two roles share an introducing entity', () => {
    const roles = [
      role({ key: 'buyer', label: 'Buyer', introduced: true, introSites: [anchor('pIntro')], uses: [anchor('a')] }),
      role({ key: 'purchaser', label: 'Purchaser', introduced: true, introSites: [anchor('pIntro')], uses: [anchor('a')] }),
    ];
    const anchors = [
      pa('acme', 'ACME Pte Ltd', 'buyer', 'Buyer'),
      pa('acme', 'ACME Pte Ltd', 'purchaser', 'Purchaser'),
    ];
    const out = serialiseCastList({ roles, anchors }, { cite });
    expect(out.parties.every(p => p.finding === 'inconsistent-label')).toBe(true);
  });
});
