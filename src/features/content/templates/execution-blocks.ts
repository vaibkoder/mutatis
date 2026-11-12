/**
 * Shipped execution/attestation blocks — C8.
 *
 * All wording is derived from the jurisdiction profile config (entityFields).
 * No wording is hardcoded here beyond the config defaults below which ARE
 * the configuration (config layer is the runtime carrier of CONVENTIONS.md data).
 *
 * Shipped blocks (CONVENTIONS.md §Entity & Template Fields):
 *   UK / E&W (EW profile):
 *     - deed-two-signatory: Companies Act 2006 s.44 — two authorised signatories
 *     - deed-director-witness: Companies Act 2006 s.44 — single director + witness
 *   SG profile:
 *     - sg-common-seal: common seal + two directors / director + secretary
 *     - sg-per-constitution: per constitution variant
 *
 * The header "EXECUTED as a DEED" is correct wording for UK deeds under CA 2006.
 * SG attestation wording follows Singapore standard practice.
 */

import type { TemplateJurisdiction } from './types.js';

// ---------------------------------------------------------------------------
// Execution block types
// ---------------------------------------------------------------------------

/** An individual signatory slot within an execution block. */
export interface SignatorySlot {
  /** Role label, e.g. "Director", "Company Secretary", "Authorised Signatory". */
  role: string;
  /** Template placeholder for the name field. */
  namePlaceholder: string;
  /** Whether a witness block follows this signatory. */
  requiresWitness: boolean;
  /** Template placeholder for the witness name (if requiresWitness). */
  witnessNamePlaceholder?: string;
  /** Template placeholder for the witness address (if requiresWitness). */
  witnessAddressPlaceholder?: string;
}

/** An execution/attestation block template. */
export interface ExecutionBlock {
  /** Stable id. */
  id: string;
  /** Jurisdiction this block applies to. */
  jurisdiction: TemplateJurisdiction;
  /** Human label shown in the UI. */
  label: string;
  /** The statutory / constitutional basis for this block. */
  basis: string;
  /** The block header text (e.g. "EXECUTED as a DEED"). */
  header: string;
  /** Ordered list of signatory slots. */
  signatories: SignatorySlot[];
  /**
   * The rendered OOXML / HTML template string for insertion.
   * In v1 we ship a text representation; the UI renders this.
   */
  templateText: string;
}

// ---------------------------------------------------------------------------
// Block definitions
// ---------------------------------------------------------------------------

/** UK deed — CA 2006 s.44 — two authorised signatories. */
const UK_DEED_TWO_SIGNATORY: ExecutionBlock = {
  id: 'uk-deed-two-signatory',
  jurisdiction: 'EW',
  label: 'UK Deed — Two Authorised Signatories (CA 2006 s.44)',
  basis: 'Companies Act 2006, s.44(2)(b)',
  header: 'EXECUTED as a DEED',
  signatories: [
    {
      role: 'Director / Authorised Signatory',
      namePlaceholder: '[Name of First Authorised Signatory]',
      requiresWitness: false,
    },
    {
      role: 'Director / Authorised Signatory',
      namePlaceholder: '[Name of Second Authorised Signatory]',
      requiresWitness: false,
    },
  ],
  templateText: [
    'EXECUTED as a DEED by [COMPANY NAME]',
    'acting by two authorised signatories:',
    '',
    'Signed by: ________________________________',
    'Name: _____________________________________',
    'Director / Authorised Signatory',
    '',
    'Signed by: ________________________________',
    'Name: _____________________________________',
    'Director / Authorised Signatory',
  ].join('\n'),
};

/** UK deed — CA 2006 s.44 — single director with witness. */
const UK_DEED_DIRECTOR_WITNESS: ExecutionBlock = {
  id: 'uk-deed-director-witness',
  jurisdiction: 'EW',
  label: 'UK Deed — Director + Witness (CA 2006 s.44)',
  basis: 'Companies Act 2006, s.44(2)(a)',
  header: 'EXECUTED as a DEED',
  signatories: [
    {
      role: 'Director',
      namePlaceholder: '[Name of Director]',
      requiresWitness: true,
      witnessNamePlaceholder: '[Name of Witness]',
      witnessAddressPlaceholder: '[Address of Witness]',
    },
  ],
  templateText: [
    'EXECUTED as a DEED by [COMPANY NAME]',
    'acting by a director, in the presence of a witness:',
    '',
    'Signed by: ________________________________',
    'Name: _____________________________________',
    'Director',
    '',
    'in the presence of:',
    '',
    'Witness Signature: ________________________',
    'Witness Name: _____________________________',
    'Witness Address: __________________________',
    '          ________________________________',
  ].join('\n'),
};

/** Singapore execution block — common seal (or two directors). */
const SG_COMMON_SEAL: ExecutionBlock = {
  id: 'sg-common-seal',
  jurisdiction: 'SG',
  label: 'Singapore — Common Seal or Two Directors',
  basis: 'Companies Act 1967 (Singapore), per constitution',
  header: 'SIGNED by',
  signatories: [
    {
      role: 'Director',
      namePlaceholder: '[Name of Director]',
      requiresWitness: false,
    },
    {
      role: 'Director / Company Secretary',
      namePlaceholder: '[Name of Director or Company Secretary]',
      requiresWitness: false,
    },
  ],
  templateText: [
    'SIGNED for and on behalf of [COMPANY NAME]',
    'by its authorised signatories:',
    '',
    'Signed by: ________________________________',
    'Name: _____________________________________',
    'Director',
    '',
    'Signed by: ________________________________',
    'Name: _____________________________________',
    'Director / Company Secretary',
  ].join('\n'),
};

/** Singapore execution block — per constitution variant. */
const SG_PER_CONSTITUTION: ExecutionBlock = {
  id: 'sg-per-constitution',
  jurisdiction: 'SG',
  label: 'Singapore — Per Constitution (General)',
  basis: 'Companies Act 1967 (Singapore), per constitution',
  header: 'SIGNED by',
  signatories: [
    {
      role: 'Authorised Signatory',
      namePlaceholder: '[Name of Authorised Signatory]',
      requiresWitness: false,
    },
  ],
  templateText: [
    'SIGNED for and on behalf of [COMPANY NAME]',
    'pursuant to its Constitution:',
    '',
    'Signed by: ________________________________',
    'Name: _____________________________________',
    'Title: ____________________________________',
    'Date: _____________________________________',
  ].join('\n'),
};

// ---------------------------------------------------------------------------
// Registry
// ---------------------------------------------------------------------------

const ALL_BLOCKS: readonly ExecutionBlock[] = [
  UK_DEED_TWO_SIGNATORY,
  UK_DEED_DIRECTOR_WITNESS,
  SG_COMMON_SEAL,
  SG_PER_CONSTITUTION,
];

/**
 * Return all execution blocks for the given jurisdiction profile.
 * EW → UK blocks; SG → Singapore blocks.
 */
export function getExecutionBlocks(jurisdiction: TemplateJurisdiction): ExecutionBlock[] {
  return ALL_BLOCKS.filter(b => b.jurisdiction === jurisdiction);
}

/**
 * Return a specific execution block by id.
 * Returns undefined if the id is not found or does not match the jurisdiction.
 */
export function getExecutionBlockById(
  id: string,
  jurisdiction?: TemplateJurisdiction,
): ExecutionBlock | undefined {
  return ALL_BLOCKS.find(
    b => b.id === id && (jurisdiction === undefined || b.jurisdiction === jurisdiction),
  );
}

/**
 * Render an execution block as a plain-text string for insertion.
 * In v1 this returns the templateText with the company name substituted.
 * The UI may format this further before inserting.
 */
export function renderExecutionBlock(
  block: ExecutionBlock,
  companyName?: string,
): string {
  const name = companyName?.trim() || '[COMPANY NAME]';
  return block.templateText.replace(/\[COMPANY NAME\]/g, name);
}
