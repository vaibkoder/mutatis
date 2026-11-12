import { describe, it, expect } from 'vitest';
import { runPartyRole, analysePartyRole } from '@features/party-role/index.js';
import { SG_PROFILE } from '@shared/config/index.js';
import type { ParaNode } from '@shared/doc-model/index.js';

/** Build body paragraphs; an optional `{ list }` marker makes a clause node. */
function paras(...rows: Array<string | { text: string; list: string }>): ParaNode[] {
  return rows.map((row, i) => {
    const text = typeof row === 'string' ? row : row.text;
    const base: ParaNode = {
      id: `body-para-${i}`,
      index: i,
      text,
      styleName: 'Normal',
      container: 'body' as const,
      getRange: () => { throw new Error('not used'); },
    };
    if (typeof row !== 'string') return { ...base, listString: row.list, listLevel: 0 };
    return base;
  });
}

function run(...rows: Array<string | { text: string; list: string }>) {
  return runPartyRole(paras(...rows), SG_PROFILE);
}

const undef = (fs: ReturnType<typeof run>) => fs.filter(f => f.category === 'party-role/undefined-role');
const incon = (fs: ReturnType<typeof run>) => fs.filter(f => f.category === 'party-role/inconsistent-label');

describe('Party-Role Consistency — PR2 inconsistent label (same party, two names)', () => {
  it('T2 — Form 1: two labels introduced against the same entity', () => {
    const fs = run(
      'ACME Pte Ltd (the "Buyer") is a party to this agreement.',
      'ACME Pte Ltd (the "Purchaser") shall complete the transaction.',
    );
    const c = incon(fs);
    expect(c).toHaveLength(1);
    expect(c[0]?.title).toMatch(/Buyer/);
    expect(c[0]?.title).toMatch(/Purchaser/);
    expect(c[0]?.featureId).toBe('party-role');
    expect(c[0]?.detail).toContain('ACME');
  });

  it('Form 1 still groups when a capitalised lead-in precedes the entity', () => {
    // A sentence lead-in ("Separately,") must not be absorbed into the entity
    // key — otherwise the two introductions would split across keys and miss.
    const fs = run(
      'ACME Pte Ltd (the "Buyer") is a party.',
      'Separately, ACME Pte Ltd (the "Purchaser") shall complete.',
    );
    expect(incon(fs)).toHaveLength(1);
  });

  it('T3 — Form 2: one label defined as the other', () => {
    const fs = run(
      'ACME Pte Ltd (the "Buyer") is the buyer.',
      '"Purchaser" means the Buyer.',
    );
    const c = incon(fs);
    expect(c).toHaveLength(1);
    expect(c[0]?.title).toMatch(/Purchaser/);
    expect(c[0]?.title).toMatch(/Buyer/);
  });

  it('T4 — dual-role deal (different entities) is NOT flagged', () => {
    const fs = run(
      'ACME Pte Ltd (the "Buyer") is a party.',
      'BIDCO Pte Ltd (the "Purchaser") is a party.',
      'The Buyer shall pay the Purchaser.',
    );
    expect(incon(fs)).toHaveLength(0);
  });

  it('T10 — co-presence without co-reference evidence is NOT flagged', () => {
    const fs = run(
      'ACME Pte Ltd (the "Buyer") is a party.',
      'BIDCO Pte Ltd (the "Seller") is a party.',
      'The Buyer shall pay the Seller on completion.',
    );
    expect(incon(fs)).toHaveLength(0);
  });
});

describe('Party-Role Consistency — PR1 used but never introduced', () => {
  it('T5 — a role used >= threshold times but never introduced is flagged', () => {
    const fs = run(
      'The Agent shall notify the other parties.',
      'The Agent may delegate its functions.',
      "The Agent's fees are payable on demand.",
    );
    const c = undef(fs);
    expect(c).toHaveLength(1);
    expect(c[0]?.title).toMatch(/Agent/);
    expect(c[0]?.occurrences?.length).toBe(3);
    expect(c[0]?.severity).toBe('warning');
  });

  it('T6 — a role used below the threshold is NOT flagged', () => {
    const fs = run('The Agent shall notify the parties.');
    expect(undef(fs)).toHaveLength(0);
  });

  it('T7 — a role introduced in the parties preamble is NOT flagged', () => {
    const fs = run(
      'BETWEEN: (1) ACME Pte Ltd and (2) BANKCO (the Agent).',
      { text: 'Payments', list: '1' },
      'The Agent shall notify the parties.',
      'The Agent may delegate its functions.',
      "The Agent's fees are payable on demand.",
    );
    expect(undef(fs)).toHaveLength(0);
  });

  it('T8 — the collective role "Parties" is never flagged', () => {
    const fs = run(
      'The Parties agree to the terms.',
      'The Parties shall act in good faith.',
      'The Parties may amend this agreement.',
      'The Parties acknowledge the risks.',
    );
    expect(undef(fs)).toHaveLength(0);
  });
});

describe('Party-Role Consistency — clean / plural cases', () => {
  it('T1 — a single consistently-labelled party produces no findings', () => {
    const fs = run(
      'ACME Pte Ltd (the "Buyer") shall pay the price.',
      'The Buyer shall complete on the completion date.',
    );
    expect(fs).toHaveLength(0);
  });

  it('T9 — plural / possessive of an introduced role produces no findings', () => {
    const fs = run(
      'ACME Pte Ltd (the "Shareholder") holds the shares.',
      "The Shareholders shall meet and each Shareholder's vote counts.",
    );
    expect(fs).toHaveLength(0);
  });
});

describe('Party-Role Consistency — analysis shape', () => {
  it('reports counts and routes all findings to featureId party-role', () => {
    const a = analysePartyRole(
      paras(
        'ACME Pte Ltd (the "Buyer") is a party.',
        'ACME Pte Ltd (the "Purchaser") shall complete.',
        'The Agent shall act. The Agent may delegate. The Agent decides.',
      ),
      SG_PROFILE,
    );
    expect(a.inconsistentLabelCount).toBe(1);
    expect(a.undefinedRoleCount).toBe(1);
    expect(a.findings.every(f => f.featureId === 'party-role')).toBe(true);
  });
});
