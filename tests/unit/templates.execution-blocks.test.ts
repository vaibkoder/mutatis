/**
 * C8 — Execution blocks tests.
 * C-T8: UK director+witness variant renders "EXECUTED as a DEED" header +
 *        correct witness attestation; SG block available under SG profile;
 *        both s.44 variants present.
 */
import { describe, it, expect } from 'vitest';
import {
  getExecutionBlocks,
  getExecutionBlockById,
  renderExecutionBlock,
} from '@features/content/templates/execution-blocks.js';

describe('C8 — getExecutionBlocks (C-T8)', () => {
  it('returns UK deed blocks under EW profile', () => {
    const blocks = getExecutionBlocks('EW');
    expect(blocks.length).toBeGreaterThanOrEqual(2);
    const ids = blocks.map(b => b.id);
    expect(ids).toContain('uk-deed-two-signatory');
    expect(ids).toContain('uk-deed-director-witness');
  });

  it('returns SG blocks under SG profile', () => {
    const blocks = getExecutionBlocks('SG');
    expect(blocks.length).toBeGreaterThanOrEqual(1);
    const ids = blocks.map(b => b.id);
    expect(ids).toContain('sg-common-seal');
  });

  it('EW blocks do not include SG blocks and vice versa', () => {
    const ewBlocks = getExecutionBlocks('EW');
    const sgBlocks = getExecutionBlocks('SG');
    expect(ewBlocks.every(b => b.jurisdiction === 'EW')).toBe(true);
    expect(sgBlocks.every(b => b.jurisdiction === 'SG')).toBe(true);
  });
});

describe('C8 — UK deed blocks (C-T8)', () => {
  it('UK deed director+witness block has "EXECUTED as a DEED" header', () => {
    const block = getExecutionBlockById('uk-deed-director-witness', 'EW');
    expect(block).toBeTruthy();
    expect(block?.header).toBe('EXECUTED as a DEED');
  });

  it('UK deed director+witness block template contains witness attestation wording', () => {
    const block = getExecutionBlockById('uk-deed-director-witness', 'EW');
    expect(block?.templateText).toMatch(/in the presence of/i);
    expect(block?.templateText).toMatch(/witness/i);
  });

  it('UK deed two-signatory block has "EXECUTED as a DEED" header', () => {
    const block = getExecutionBlockById('uk-deed-two-signatory', 'EW');
    expect(block?.header).toBe('EXECUTED as a DEED');
  });

  it('UK deed two-signatory has two signatory slots', () => {
    const block = getExecutionBlockById('uk-deed-two-signatory', 'EW');
    expect(block?.signatories).toHaveLength(2);
  });

  it('UK deed director+witness has one signatory with requiresWitness=true', () => {
    const block = getExecutionBlockById('uk-deed-director-witness', 'EW');
    expect(block?.signatories[0]?.requiresWitness).toBe(true);
    expect(block?.signatories[0]?.witnessNamePlaceholder).toBeTruthy();
  });

  it('basis cites Companies Act 2006 s.44 for both UK variants', () => {
    const twoSig = getExecutionBlockById('uk-deed-two-signatory', 'EW');
    const dirWit = getExecutionBlockById('uk-deed-director-witness', 'EW');
    expect(twoSig?.basis).toMatch(/s\.44/);
    expect(dirWit?.basis).toMatch(/s\.44/);
    expect(twoSig?.basis).toMatch(/Companies Act 2006/);
    expect(dirWit?.basis).toMatch(/Companies Act 2006/);
  });
});

describe('C8 — SG execution blocks (C-T8)', () => {
  it('SG common seal block is available under SG profile', () => {
    const block = getExecutionBlockById('sg-common-seal', 'SG');
    expect(block).toBeTruthy();
    expect(block?.jurisdiction).toBe('SG');
  });

  it('SG block references Singapore Companies Act', () => {
    const block = getExecutionBlockById('sg-common-seal', 'SG');
    expect(block?.basis).toMatch(/Singapore/);
    expect(block?.basis).toMatch(/Companies Act/);
  });

  it('SG block does not reference CA 2006 (wrong jurisdiction)', () => {
    const block = getExecutionBlockById('sg-common-seal', 'SG');
    expect(block?.basis).not.toContain('2006');
  });
});

describe('C8 — renderExecutionBlock', () => {
  it('substitutes company name', () => {
    const block = getExecutionBlockById('uk-deed-two-signatory', 'EW')!;
    const rendered = renderExecutionBlock(block, 'Acme Ltd');
    expect(rendered).toContain('Acme Ltd');
    expect(rendered).not.toContain('[COMPANY NAME]');
  });

  it('uses placeholder when company name is not provided', () => {
    const block = getExecutionBlockById('uk-deed-two-signatory', 'EW')!;
    const rendered = renderExecutionBlock(block);
    expect(rendered).toContain('[COMPANY NAME]');
  });
});
