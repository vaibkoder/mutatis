import { describe, it, expect } from 'vitest';

import {
  flattenTrackedChanges,
  quarantineRevisions,
} from '@features/compare/preprocess/index.js';

const P = (inner: string): string => `<w:p>${inner}</w:p>`;
const run = (t: string): string => `<w:r><w:t>${t}</w:t></w:r>`;

describe('flattenTrackedChanges (§5.0)', () => {
  it('accept-all keeps <w:ins> content and drops <w:del> content, unwrapping both', () => {
    const xml = P(
      `${run('Keep ')}<w:ins w:id="1" w:author="A"><w:r><w:t>added</w:t></w:r></w:ins>` +
        `<w:del w:id="2"><w:r><w:delText>removed</w:delText></w:r></w:del>`,
    );
    const out = flattenTrackedChanges(xml, 'acceptAll');
    expect(out).toContain('Keep ');
    expect(out).toContain('added');
    expect(out).not.toContain('removed');
    // Wrapper elements are gone.
    expect(out).not.toContain('<w:ins');
    expect(out).not.toContain('<w:del');
  });

  it('reject-all keeps <w:del> content (delText→t) and drops <w:ins> content', () => {
    const xml = P(
      `${run('Keep ')}<w:ins w:id="1"><w:r><w:t>added</w:t></w:r></w:ins>` +
        `<w:del w:id="2"><w:r><w:delText>removed</w:delText></w:r></w:del>`,
    );
    const out = flattenTrackedChanges(xml, 'rejectAll');
    expect(out).toContain('Keep ');
    expect(out).toContain('removed');
    expect(out).not.toContain('added');
    // delText was converted back to a plain <w:t>.
    expect(out).toContain('<w:t>removed</w:t>');
    expect(out).not.toContain('delText');
  });

  it('accept-all drops <w:moveFrom> and keeps <w:moveTo>; strips range markers', () => {
    const xml = P(
      `<w:moveFromRangeStart w:id="3" w:name="m"/>` +
        `<w:moveFrom w:id="4"><w:r><w:t>oldspot</w:t></w:r></w:moveFrom>` +
        `<w:moveFromRangeEnd w:id="3"/>` +
        `<w:moveToRangeStart w:id="5" w:name="m"/>` +
        `<w:moveTo w:id="6"><w:r><w:t>newspot</w:t></w:r></w:moveTo>` +
        `<w:moveToRangeEnd w:id="5"/>`,
    );
    const out = flattenTrackedChanges(xml, 'acceptAll');
    expect(out).toContain('newspot');
    expect(out).not.toContain('oldspot');
    expect(out).not.toContain('moveFromRangeStart');
    expect(out).not.toContain('moveToRangeStart');
    expect(out).not.toContain('<w:moveTo');
  });

  it('preserves the w:namespace prefix on surviving elements', () => {
    const xml = P(`${run('Hello')}`);
    const out = flattenTrackedChanges(xml, 'acceptAll');
    expect(out).toContain('<w:r>');
    expect(out).toContain('<w:t>Hello</w:t>');
  });
});

describe('quarantineRevisions (§5.0)', () => {
  it('strips residual comment anchors and counts them', () => {
    const xml = P(
      `<w:commentRangeStart w:id="0"/>${run('text')}<w:commentRangeEnd w:id="0"/>` +
        `<w:r><w:commentReference w:id="0"/></w:r>`,
    );
    const { cleanXml, quarantinedCount } = quarantineRevisions(xml);
    expect(cleanXml).toContain('text');
    expect(cleanXml).not.toContain('commentRangeStart');
    expect(cleanXml).not.toContain('commentReference');
    expect(quarantinedCount).toBe(3);
  });

  it('reports zero on a clean part', () => {
    const { quarantinedCount } = quarantineRevisions(P(run('clean')));
    expect(quarantinedCount).toBe(0);
  });
});
