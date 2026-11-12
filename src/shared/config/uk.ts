import type { JurisdictionProfile } from './types.js';
import {
  DEFINITION_VERBS, DETERMINERS, PARTY_ROLES, CONVENTIONAL_TERMS,
  STATUTORY_WORDS, STATUTORY_CITATION, UK_ACRONYMS, UK_PROPER_NOUNS,
} from './defined-term-lexicons.js';

export const UK_PROFILE: JurisdictionProfile = {
  id: 'UK',
  spellingLocale: 'en-GB',

  terminology: {
    topLevelUnit: 'Clause',
    bodyAliases: ['Clause', 'Cl', 'Section', 'Sec', '§'],
    scheduleWord: 'Schedule',
    annexureAliases: ['Schedule', 'Sch', 'Exhibit', 'Exh'],
    scheduleSubUnit: 'paragraph',
    otherDivisions: ['Annex', 'Annexure', 'Appendix', 'Part'],
    recitalNouns: ['Recital', 'Background'],
    recitalStyle: 'letteredUpper',
    sectionIsTopLevelDivision: false,
    selfReferenceTerms: [
      'this Agreement', 'this Deed', 'this Clause', 'this Schedule',
      'this Annex', 'this Appendix', 'this Part', 'this paragraph',
      'hereof', 'herein', 'hereunder', 'the foregoing',
    ],
  },

  currency: {
    table: {
      '£': { code: 'GBP' },
      'GBP': { code: 'GBP' },
      'Pounds': { code: 'GBP' },
      'Pounds Sterling': { code: 'GBP' },
      'S$': { code: 'SGD' },
      'SGD': { code: 'SGD' },
      'US$': { code: 'USD' },
      'USD': { code: 'USD' },
      '€': { code: 'EUR' },
      'EUR': { code: 'EUR' },
      'Euros': { code: 'EUR' },
      'RM': { code: 'MYR' },
      'MYR': { code: 'MYR' },
      'Ringgit': { code: 'MYR' },
      'HK$': { code: 'HKD' },
      'HKD': { code: 'HKD' },
      '$': { code: 'AMBIGUOUS' },
      'Dollars': { code: 'DOLLAR_FAMILY' },
    },
    ambiguousDollarPolicy: 'flag-soft',
    centTolerance: 0.005,
  },

  dateLocale: {
    order: 'little-endian',
    defaultDisplayFormat: 'DD Month YYYY',
    weekendDays: [0, 6],
    executionAnchorKeywords: [
      'executed', 'execution date', 'dated', 'signed',
      'completion', 'completion date', 'shall complete',
    ],
    effectiveDateKeywords: [
      'with effect from', 'as of', 'effective date', 'effective as of',
    ],
    flagAllWeekendDates: false,
    roleHintWindow: 120,
  },

  definedTerms: {
    quoteStyle: 'double',
    verbs: DEFINITION_VERBS,
    styleNudges: { shallMean: false },
    capMode: 'smart',
    determiners: DETERMINERS,
    // Allowlist composed from the shared lexicon + UK-specific entries (ADR-0015).
    exclusions: {
      partyRoles: PARTY_ROLES,
      conventional: CONVENTIONAL_TERMS,
      acronyms: UK_ACRONYMS,
      statutoryWords: STATUTORY_WORDS,
      properNouns: UK_PROPER_NOUNS,
      statutoryCitation: STATUTORY_CITATION,
    },
    flagLocalRedefinitions: false,
    recogniseUnquoted: true,
    recogniseColon: true,
  },

  placeholderVocab: {
    bulletChars: ['●', '•', '∙', '·', '*', ''],
    keywords: ['TBC', 'TBD', 'DATE', 'NAME', 'AMOUNT', 'PARTY', 'ADDRESS', 'INSERT', 'NTD', 'drafting note'],
    whitelist: ['[INTENTIONALLY LEFT BLANK]', '[Reserved]'],
    citationReporters: [
      'EWHC', 'EWCA', 'UKSC', 'UKHL',
      'SGHC', 'SGHC(A)', 'SGCA', 'SGCA(I)',
      'SLR', 'WLR', 'AC', 'All ER', 'QB', 'Ch',
    ],
    signatureAnchors: ['Signature', 'By:', 'Name:', 'Title:', 'Witnessed by', 'In the presence of'],
    flagBareBrackets: false,
  },

  numberingPresets: [
    {
      id: 'preset-uk-sg-commercial',
      name: 'UK/SG Commercial',
      jurisdiction: 'SG-UK',
      description: 'Standard UK & Singapore commercial hierarchy: 1 / 1.1 / 1.1.1 / (a) / (i) / (A). Closed-bracket form. Default scheme.',
      builtIn: true,
      unitLabel: 'clause',
      schemaVersion: 1,
      suppressForSgUk: false,
      levels: [
        { levelIndex: 0, listNumbering: 'arabic',      formatString: '%1.',       prefix: '', suffix: '', indent: { textIndentPt: 0,   numberIndentPt: 0   }, labelStyle: 'decimal-compound',   restartAfterLevel: -1 },
        { levelIndex: 1, listNumbering: 'arabic',      formatString: '%1.%2.',    prefix: '', suffix: '', indent: { textIndentPt: 36,  numberIndentPt: 18  }, labelStyle: 'decimal-compound',   restartAfterLevel: 0  },
        { levelIndex: 2, listNumbering: 'arabic',      formatString: '%1.%2.%3.', prefix: '', suffix: '', indent: { textIndentPt: 72,  numberIndentPt: 36  }, labelStyle: 'decimal-compound',   restartAfterLevel: 1  },
        { levelIndex: 3, listNumbering: 'lowerLetter', formatString: '(%4)',      prefix: '', suffix: '', indent: { textIndentPt: 108, numberIndentPt: 72  }, labelStyle: 'paren-lower-alpha',  restartAfterLevel: 2  },
        { levelIndex: 4, listNumbering: 'lowerRoman',  formatString: '(%5)',      prefix: '', suffix: '', indent: { textIndentPt: 144, numberIndentPt: 108 }, labelStyle: 'paren-lower-roman',  restartAfterLevel: 3  },
        { levelIndex: 5, listNumbering: 'upperLetter', formatString: '(%6)',      prefix: '', suffix: '', indent: { textIndentPt: 180, numberIndentPt: 144 }, labelStyle: 'paren-upper-alpha',  restartAfterLevel: 4  },
        { levelIndex: 6, listNumbering: 'arabic',      formatString: '(%7)',      prefix: '', suffix: '', indent: { textIndentPt: 216, numberIndentPt: 180 }, labelStyle: 'plain',              restartAfterLevel: 5  },
        { levelIndex: 7, listNumbering: 'lowerLetter', formatString: '(%8)',      prefix: '', suffix: '', indent: { textIndentPt: 252, numberIndentPt: 216 }, labelStyle: 'paren-lower-alpha',  restartAfterLevel: 6  },
        { levelIndex: 8, listNumbering: 'lowerRoman',  formatString: '(%9)',      prefix: '', suffix: '', indent: { textIndentPt: 288, numberIndentPt: 252 }, labelStyle: 'paren-lower-roman',  restartAfterLevel: 7  },
      ],
    },
    {
      id: 'preset-schedule',
      name: 'Schedule (paragraph)',
      jurisdiction: 'SG-UK',
      description: "Schedule numbering: 1 / 1.1 / (a). Unit = 'paragraph'. Top level often grouped under 'Part 1 / Part 2' headings.",
      builtIn: true,
      unitLabel: 'paragraph',
      schemaVersion: 1,
      suppressForSgUk: false,
      levels: [
        { levelIndex: 0, listNumbering: 'arabic',      formatString: '%1.',    prefix: '', suffix: '', indent: { textIndentPt: 0,   numberIndentPt: 0  }, labelStyle: 'decimal-compound',  restartAfterLevel: -1 },
        { levelIndex: 1, listNumbering: 'arabic',      formatString: '%1.%2.', prefix: '', suffix: '', indent: { textIndentPt: 36,  numberIndentPt: 18 }, labelStyle: 'decimal-compound',  restartAfterLevel: 0  },
        { levelIndex: 2, listNumbering: 'lowerLetter', formatString: '(%3)',   prefix: '', suffix: '', indent: { textIndentPt: 72,  numberIndentPt: 36 }, labelStyle: 'paren-lower-alpha', restartAfterLevel: 1  },
        { levelIndex: 3, listNumbering: 'lowerRoman',  formatString: '(%4)',   prefix: '', suffix: '', indent: { textIndentPt: 108, numberIndentPt: 72 }, labelStyle: 'paren-lower-roman', restartAfterLevel: 2  },
      ],
    },
    {
      id: 'preset-schedule-part',
      name: 'Schedule (Part 1 / Part 2 grouping)',
      jurisdiction: 'SG-UK',
      description: "Variant of Schedule preset where level 0 is 'Part 1 / Part 2' as a numbered list item rather than a heading style.",
      builtIn: true,
      unitLabel: 'paragraph',
      schemaVersion: 1,
      suppressForSgUk: false,
      levels: [
        { levelIndex: 0, listNumbering: 'arabic',      formatString: 'Part %1', prefix: '', suffix: '', indent: { textIndentPt: 0,   numberIndentPt: 0  }, labelStyle: 'plain',             restartAfterLevel: -1 },
        { levelIndex: 1, listNumbering: 'arabic',      formatString: '%2.',     prefix: '', suffix: '', indent: { textIndentPt: 36,  numberIndentPt: 18 }, labelStyle: 'decimal-compound',  restartAfterLevel: 0  },
        { levelIndex: 2, listNumbering: 'arabic',      formatString: '%2.%3.',  prefix: '', suffix: '', indent: { textIndentPt: 72,  numberIndentPt: 36 }, labelStyle: 'decimal-compound',  restartAfterLevel: 1  },
        { levelIndex: 3, listNumbering: 'lowerLetter', formatString: '(%4)',    prefix: '', suffix: '', indent: { textIndentPt: 108, numberIndentPt: 72 }, labelStyle: 'paren-lower-alpha', restartAfterLevel: 2  },
      ],
    },
    {
      id: 'preset-recitals',
      name: 'Recitals (A)(B)(C)',
      jurisdiction: 'SG-UK',
      description: 'Recital lettering: (A) (B) (C). Single level, parenthesised upper-letter.',
      builtIn: true,
      unitLabel: 'recital',
      schemaVersion: 1,
      suppressForSgUk: false,
      levels: [
        { levelIndex: 0, listNumbering: 'upperLetter', formatString: '(%1)', prefix: '', suffix: '', indent: { textIndentPt: 36, numberIndentPt: 0 }, labelStyle: 'paren-upper-alpha', restartAfterLevel: -1 },
      ],
    },
    {
      id: 'preset-plain-outline',
      name: 'Plain Outline (memo)',
      jurisdiction: 'SG-UK',
      description: 'Plain decimal outline for memos and advisory notes: 1 / 1.1 / 1.1.1 / 1.1.1.1. No parenthesised limbs.',
      builtIn: true,
      unitLabel: 'paragraph',
      schemaVersion: 1,
      suppressForSgUk: false,
      levels: [
        { levelIndex: 0, listNumbering: 'arabic', formatString: '%1.',         prefix: '', suffix: '', indent: { textIndentPt: 0,   numberIndentPt: 0  }, labelStyle: 'decimal-compound', restartAfterLevel: -1 },
        { levelIndex: 1, listNumbering: 'arabic', formatString: '%1.%2.',      prefix: '', suffix: '', indent: { textIndentPt: 36,  numberIndentPt: 18 }, labelStyle: 'decimal-compound', restartAfterLevel: 0  },
        { levelIndex: 2, listNumbering: 'arabic', formatString: '%1.%2.%3.',   prefix: '', suffix: '', indent: { textIndentPt: 72,  numberIndentPt: 36 }, labelStyle: 'decimal-compound', restartAfterLevel: 1  },
        { levelIndex: 3, listNumbering: 'arabic', formatString: '%1.%2.%3.%4.',prefix: '', suffix: '', indent: { textIndentPt: 108, numberIndentPt: 72 }, labelStyle: 'decimal-compound', restartAfterLevel: 2  },
      ],
    },
    {
      id: 'preset-us-litigation',
      name: 'US Litigation style',
      jurisdiction: 'US',
      description: 'I. / A. / 1. / a. outline — US litigation style. NOT for SG/UK drafting. Hidden by default in SG/UK config.',
      builtIn: true,
      unitLabel: 'section',
      schemaVersion: 1,
      suppressForSgUk: true,
      levels: [
        { levelIndex: 0, listNumbering: 'upperRoman',  formatString: '%1.', prefix: '', suffix: '', indent: { textIndentPt: 0,   numberIndentPt: 0  }, labelStyle: 'plain', restartAfterLevel: -1 },
        { levelIndex: 1, listNumbering: 'upperLetter', formatString: '%2.', prefix: '', suffix: '', indent: { textIndentPt: 36,  numberIndentPt: 18 }, labelStyle: 'plain', restartAfterLevel: 0  },
        { levelIndex: 2, listNumbering: 'arabic',      formatString: '%3.', prefix: '', suffix: '', indent: { textIndentPt: 72,  numberIndentPt: 36 }, labelStyle: 'plain', restartAfterLevel: 1  },
        { levelIndex: 3, listNumbering: 'lowerLetter', formatString: '%4.', prefix: '', suffix: '', indent: { textIndentPt: 108, numberIndentPt: 72 }, labelStyle: 'plain', restartAfterLevel: 2  },
      ],
    },
  ],

  entityFields: {
    registrationNumbers: {
      UK_CH: /^[A-Z0-9]{8}$/,
    },
    suffixes: ['Limited', 'Ltd', 'plc', 'LLP'],
  },

  xref: {
    units: {
      primaryBody: 'Clause',
      bodyAliases: ['Clause', 'Cl', 'Section', 'Sec', '§'],
      annexurePrimary: 'Schedule',
      annexureAliases: ['Schedule', 'Sch', 'Exhibit', 'Exh'],
      scheduleSubUnit: 'paragraph',
      otherDivisions: ['Annex', 'Annexure', 'Appendix', 'Part'],
      recitalNouns: ['Recital', 'Background'],
      recitalStyle: 'letteredUpper',
    },
    numbering: {
      bodyStyle: 'decimal',
      subClauseStyle: ['lowerAlphaParen', 'lowerRomanParen'],
    },
    ranges: {
      connectors: ['to', '–', '—', '-'],
      listConnectors: [',', 'and', '&'],
      inclusiveMarker: '(inclusive)',
    },
    selfReferenceTerms: [
      'this Agreement', 'this Deed', 'this Clause', 'this Schedule',
      'this Annex', 'this Appendix', 'this Part', 'this paragraph',
      'hereof', 'herein', 'hereunder', 'the foregoing',
    ],
    section: {
      isTopLevelDivision: false,
    },
    styles: {
      headingStyles: ['Heading 1', 'Heading 2', 'Heading 3', 'Heading 4'],
      scheduleHeading: ['Schedule Heading', 'Annex Heading', 'Appendix Heading'],
      tocStyles: ['TOC 1', 'TOC 2', 'TOC 3', 'TOC 4', 'TOC 5', 'TOC Heading'],
    },
  },

  staticNumbering: {
    headingSeverity: 'medium',
    bodySeverity: 'low',
    maxOoxmlProbes: 200,
  },

  partyRoleConsistency: {
    minUsesForUndefined: 3,
    pr1Severity: 'warning',
    pr2Severity: 'warning',
    excludeCollectiveRoles: ['Party', 'Parties'],
  },
};
