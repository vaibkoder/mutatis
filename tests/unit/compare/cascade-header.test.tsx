/**
 * Feature 08 — the collapsed cascade header in the NAVIGATION list (§6.1, §8.3).
 *
 * ADR-0042 split one sentence into two, because they are two facts: a
 * cross-reference cascade is the document POINTING AT a clause whose number
 * moved; a numbering cascade is the clause numbers themselves moving. Saying
 * "references renumbered" of the second claims text the document does not have.
 *
 * That split was applied to `SummaryPanel` and not to `ChangeList` — and
 * `ChangeList` is the surface a reader actually works from. Its header read
 * "Clause numbering (level 1) renumbered (+1) — 20 references" for twenty
 * paragraphs containing no reference to anything. Neither half had a test.
 */
import { describe, it, expect, afterEach } from 'vitest';
import { render } from 'preact';

import { ChangeList, cascadeLabelMap } from '@features/compare/components/ChangeList.js';
import type { CascadeAnnotation, DiffOp } from '@features/compare/model.js';

let host: HTMLDivElement | null = null;

afterEach(() => {
  if (host) render(null, host);
  host?.remove();
  host = null;
});

const cascade = (over: Partial<CascadeAnnotation> = {}): CascadeAnnotation => ({
  id: 'c1',
  tokenKind: 'Clause',
  shift: 1,
  affectedCount: 3,
  exampleRefs: ['Clause 4.2 → 4.3'],
  source: 'crossReference',
  ...over,
});

const member = (id: string): DiffOp => ({
  id,
  type: 'property',
  origRange: { start: 0, end: 0 },
  revRange: { start: 0, end: 0 },
  partId: 'document',
  classification: 'crossRefRenumber',
  classes: ['crossRefRenumber', 'formatting'],
  author: 'Mutatis',
  date: '2026-09-12T00:00:00Z',
  collapsedInto: 'c1',
});

function mount(c: CascadeAnnotation, members: DiffOp[]): string {
  host = document.createElement('div');
  document.body.appendChild(host);
  render(
    <ChangeList
      ops={members}
      filters={{ hideFormattingOnly: false, collapseRenumbering: true }}
      cascadeLabels={cascadeLabelMap([c])}
      onGoToChange={() => {}}
    />,
    host,
  );
  return host.textContent ?? '';
}

describe('cascadeLabelMap', () => {
  it('calls a cross-reference cascade what it is: references renumbered', () => {
    const label = cascadeLabelMap([cascade()]).get('c1')!;
    expect(label.text).toBe('Clause references renumbered (+1)');
    expect(label.memberNoun).toBe('reference');
  });

  it('calls a numbering cascade what it is: the numbers themselves shifted', () => {
    const label = cascadeLabelMap([
      cascade({ source: 'numbering', tokenKind: 'Clause numbering (level 1)' }),
    ]).get('c1')!;
    expect(label.text).toBe('Clause numbering (level 1) shifted (+1)');
    // These paragraphs contain no reference to anything. Their own numbers moved.
    expect(label.text).not.toContain('references');
    expect(label.memberNoun).toBe('paragraph');
  });

  it('treats a cascade with no source as a cross-reference one (pre-ADR-0042)', () => {
    const { source: _omitted, ...noSource } = cascade();
    expect(cascadeLabelMap([noSource as CascadeAnnotation]).get('c1')!.memberNoun).toBe('reference');
  });

  it('renders a negative shift with its sign', () => {
    expect(cascadeLabelMap([cascade({ shift: -2 })]).get('c1')!.text).toContain('(-2)');
  });
});

describe('the collapsed header as the reader sees it', () => {
  const members = [member('m1'), member('m2'), member('m3')];

  it('counts references for a cross-reference cascade', () => {
    const text = mount(cascade(), members);
    expect(text).toContain('Clause references renumbered (+1) — 3 references');
  });

  it('counts PARAGRAPHS for a numbering cascade, and never says "references"', () => {
    const text = mount(
      cascade({ source: 'numbering', tokenKind: 'Clause numbering (level 1)' }),
      members,
    );
    expect(text).toContain('Clause numbering (level 1) shifted (+1) — 3 paragraphs');
    expect(text).not.toContain('reference');
  });

  it('singularises a one-member group', () => {
    const text = mount(cascade({ source: 'numbering', affectedCount: 1 }), [member('m1')]);
    expect(text).toContain('— 1 paragraph');
    expect(text).not.toContain('1 paragraphs');
  });
});
