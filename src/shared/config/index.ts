export type {
  JurisdictionProfile, JurisdictionId,
  DefinedTermConfig, DefinedTermExclusions, PartyRoleConsistencyConfig,
  LookupConfig, LookupDefinitionConfig,
  CrossRefConfig, StaticNumberingConfig, StaticSeverityLabel,
  XrefUnitsConfig, XrefNumberingConfig, XrefRangesConfig, XrefSectionConfig, XrefStylesConfig,
  NumberingPreset, NumberingLevel, NumberingConfig, NumberingWhitespaceOverride,
  ListNumberingStyle, LevelLabelStyle,
  CompareConfig, FlattenPolicy, FootnoteCompareMode,
  CompareBatchConfig, CompareBatchPairingConfig,
} from './types.js';
export { SG_PROFILE } from './sg.js';
export { UK_PROFILE } from './uk.js';

import type { JurisdictionId, JurisdictionProfile, LookupConfig, NumberingConfig, CompareConfig } from './types.js';
import { SG_PROFILE } from './sg.js';
import { UK_PROFILE } from './uk.js';

const PROFILES: Readonly<Record<JurisdictionId, JurisdictionProfile>> = {
  SG: SG_PROFILE,
  UK: UK_PROFILE,
};

/** Retrieve a jurisdiction profile by ID. */
export function getProfile(id: JurisdictionId): JurisdictionProfile {
  return PROFILES[id];
}

const STORAGE_KEY = 'mutatis.jurisdiction';

/** Read the user's persisted jurisdiction preference (falls back to UK). */
export function getStoredJurisdiction(): JurisdictionId {
  const stored = localStorage.getItem(STORAGE_KEY);
  if (stored === 'SG' || stored === 'UK') return stored;
  return 'UK';
}

/** Persist the user's jurisdiction preference. */
export function setStoredJurisdiction(id: JurisdictionId): void {
  localStorage.setItem(STORAGE_KEY, id);
}

/** Get the active profile from localStorage. */
export function getActiveProfile(): JurisdictionProfile {
  return getProfile(getStoredJurisdiction());
}

const DEFAULT_LOOKUP_CONFIG: LookupConfig = {
  debounceMs: 400,
  dictRebuildDebounceMs: 1000,
  sentenceWindowChars: 240,
  previewMaxChars: 600,
  expandedPreviewMaxChars: 4000,
  definition: {
    quoteStyle: 'smart',
    introVerbs: ['means', 'shall mean', 'has the meaning'],
  },
};

const DEFAULT_NUMBERING_CONFIG: NumberingConfig = {
  defaultSchemeId: 'preset-uk-sg-commercial',
  reviewHighlightColour: 'Yellow',
  whitespaceOverride: 'auto',
};

/** Return the numbering config for a profile, falling back to documented defaults. */
export function getNumberingConfig(profile: JurisdictionProfile): NumberingConfig {
  return profile.numbering ?? DEFAULT_NUMBERING_CONFIG;
}

/** Return the lookup config for a profile, falling back to documented defaults. */
export function getLookupConfig(profile: JurisdictionProfile): LookupConfig {
  const override = profile.lookup;
  if (!override) return DEFAULT_LOOKUP_CONFIG;
  return {
    ...DEFAULT_LOOKUP_CONFIG,
    ...override,
    definition: { ...DEFAULT_LOOKUP_CONFIG.definition, ...override.definition },
  };
}

/**
 * Bidirectional en-US → en-GB variant map shared by both SG and UK profiles.
 * Keys are lowercase US spellings; values are lowercase en-GB equivalents.
 * Detection is case-insensitive (spec §7.1, CONVENTIONS.md §en-GB Spelling).
 */
export const COMMON_SPELLING_VARIANTS: Readonly<Record<string, string>> = {
  // -ize/-ise
  organize: 'organise', organizes: 'organises', organized: 'organised', organizing: 'organising',
  recognize: 'recognise', recognizes: 'recognises', recognized: 'recognised', recognizing: 'recognising',
  analyze: 'analyse', analyzes: 'analyses', analyzed: 'analysed', analyzing: 'analysing',
  realize: 'realise', realizes: 'realises', realized: 'realised', realizing: 'realising',
  specialize: 'specialise', specializes: 'specialises', specialized: 'specialised', specializing: 'specialising',
  authorize: 'authorise', authorizes: 'authorises', authorized: 'authorised', authorizing: 'authorising',
  finalize: 'finalise', finalizes: 'finalises', finalized: 'finalised', finalizing: 'finalising',
  prioritize: 'prioritise', prioritizes: 'prioritises', prioritized: 'prioritised', prioritizing: 'prioritising',
  minimize: 'minimise', minimizes: 'minimises', minimized: 'minimised', minimizing: 'minimising',
  maximize: 'maximise', maximizes: 'maximises', maximized: 'maximised', maximizing: 'maximising',
  memorize: 'memorise', memorizes: 'memorises', memorized: 'memorised', memorizing: 'memorising',
  capitalize: 'capitalise', capitalizes: 'capitalises', capitalized: 'capitalised', capitalizing: 'capitalising',
  emphasize: 'emphasise', emphasizes: 'emphasises', emphasized: 'emphasised', emphasizing: 'emphasising',
  summarize: 'summarise', summarizes: 'summarises', summarized: 'summarised', summarizing: 'summarising',
  utilize: 'utilise', utilizes: 'utilises', utilized: 'utilised', utilizing: 'utilising',
  standardize: 'standardise', standardizes: 'standardises', standardized: 'standardised', standardizing: 'standardising',
  harmonize: 'harmonise', harmonizes: 'harmonises', harmonized: 'harmonised', harmonizing: 'harmonising',
  // -or/-our
  color: 'colour', colors: 'colours', colored: 'coloured', coloring: 'colouring',
  honor: 'honour', honors: 'honours', honored: 'honoured', honoring: 'honouring',
  labor: 'labour', labors: 'labours', labored: 'laboured', laboring: 'labouring',
  behavior: 'behaviour', behaviors: 'behaviours',
  favor: 'favour', favors: 'favours', favored: 'favoured', favoring: 'favouring',
  endeavor: 'endeavour', endeavors: 'endeavours',
  flavor: 'flavour', flavors: 'flavours',
  humor: 'humour', humors: 'humours',
  // -ense/-ence
  defense: 'defence', offense: 'offence', license: 'licence',
  // -er/-re
  center: 'centre', centers: 'centres', centered: 'centred',
  theater: 'theatre', theaters: 'theatres',
  fiber: 'fibre', fibers: 'fibres',
  // -og/-ogue
  catalog: 'catalogue', catalogs: 'catalogues',
  dialog: 'dialogue', dialogs: 'dialogues',
  // program/programme
  program: 'programme', programs: 'programmes',
  // fulfill/fulfil
  fulfill: 'fulfil', fulfills: 'fulfils',
  // practice (verb, en-US) → practise (verb, en-GB); noun form is identical in both
  practice: 'practise',
};

const DEFAULT_COMPARE_CONFIG: CompareConfig = {
  author: 'Mutatis Compare',
  referenceTokens: ['Clause', 'Schedule', 'paragraph'],
  primaryUnit: 'Clause',
  renumberThreshold: 3,
  replaceSimilarityThreshold: 0.5,
  replaceContiguityThreshold: 0.8,
  spellingVariants: COMMON_SPELLING_VARIANTS,
  // ADR-0047: compared by default; with headers and footers enabled an
  // unchanged pair produces zero part ops (no save-to-save noise).
  excludeHeaders: false,
  excludeFooters: false,
  footnotes: 'compare',
  flattenOriginal: 'acceptAll',
  flattenRevised: 'acceptAll',
};

/** Return the compare config for a profile, falling back to documented defaults. */
export function getCompareConfig(profile: JurisdictionProfile): CompareConfig {
  // The clause noun has ONE home in the Config Layer — `xref.units.primaryBody`
  // (spec/CONVENTIONS.md § Terminology) — and Compare reads it from there rather
  // than keeping a second copy nothing would keep in sync. A profile that wants
  // Compare to name the unit differently still says so explicitly, below.
  const primaryUnit = profile.xref.units.primaryBody;
  const override = profile.compare;
  if (!override) return { ...DEFAULT_COMPARE_CONFIG, primaryUnit };
  return {
    ...DEFAULT_COMPARE_CONFIG,
    ...override,
    primaryUnit: override.primaryUnit ?? primaryUnit,
    spellingVariants: {
      ...DEFAULT_COMPARE_CONFIG.spellingVariants,
      ...override.spellingVariants,
    },
    referenceTokens: override.referenceTokens ?? DEFAULT_COMPARE_CONFIG.referenceTokens,
  };
}
