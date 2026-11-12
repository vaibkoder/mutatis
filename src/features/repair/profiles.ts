// Style Profile schema, shipped fixtures, and loader.
// All margin/spacing values are in points (Office JS unit).

export interface ProfileFont {
  name: string;
  size: number;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: string;
}

export type LineSpacingRule = 'single' | 'onePtFive' | 'double' | 'multiple' | 'exactly' | 'atLeast';

export interface ProfileSpacing {
  lineSpacingRule: LineSpacingRule;
  lineSpacing: number;
  spaceBefore: number;
  spaceAfter: number;
}

export interface ProfileIndent {
  firstLine: number;
  left: number;
  right: number;
}

export interface ProfilePageMargins {
  top: number;
  bottom: number;
  left: number;
  right: number;
}

export interface ProfilePage {
  margins: ProfilePageMargins;
  size?: string;
  orientation?: 'portrait' | 'landscape';
  pageNumber?: { position: string; enforced: boolean };
}

export interface ProfileBody {
  font: ProfileFont;
  spacing: ProfileSpacing;
  indent: ProfileIndent;
}

export interface ProfileStyleDef {
  font?: Partial<ProfileFont>;
  paragraphFormat?: Partial<{
    lineSpacingRule: LineSpacingRule;
    lineSpacing: number;
    spaceBefore: number;
    spaceAfter: number;
    left: number;
    right: number;
  }>;
}

export interface ProfileRules {
  citationBracketStyle?: 'square' | 'round';
  quotedAuthoritySize?: number;
  normalisationOptIn: boolean;
  protectOrientationSectionBreaks: boolean;
}

export interface StyleProfile {
  id: string;
  name: string;
  jurisdiction: 'SG' | 'UK' | 'GENERIC';
  kind: 'court' | 'commercial';
  description?: string;
  version?: string;
  locked: boolean;
  page: ProfilePage;
  body: ProfileBody;
  styles: Record<string, ProfileStyleDef>;
  rules: ProfileRules;
}

const MM_TO_PT = 2.83464567;

/** SG Court Filing (Practice Directions 2021): TNR 12, double-spaced, 35mm margins. */
export const SG_COURT_FILING: StyleProfile = {
  id: 'sg-court-filing',
  name: 'SG Court Filing (PD 2021)',
  jurisdiction: 'SG',
  kind: 'court',
  locked: true,
  page: {
    margins: { top: 99.21, bottom: 99.21, left: 99.21, right: 99.21 },
    size: 'A4',
    orientation: 'portrait',
    pageNumber: { position: 'top-right', enforced: false },
  },
  body: {
    font: { name: 'Times New Roman', size: 12, color: '#000000', bold: false, italic: false },
    spacing: { lineSpacingRule: 'multiple', lineSpacing: 2.0, spaceBefore: 0, spaceAfter: 0 },
    indent: { firstLine: 0, left: 0, right: 0 },
  },
  styles: {
    'Normal': {
      font: { name: 'Times New Roman', size: 12 },
      paragraphFormat: { lineSpacingRule: 'multiple', lineSpacing: 2.0, spaceBefore: 0, spaceAfter: 0 },
    },
    'Quotation': {
      font: { name: 'Times New Roman', size: 10 },
      paragraphFormat: { lineSpacingRule: 'onePtFive', lineSpacing: 1.5, left: 36, right: 36 },
    },
    'Footnote Text': { font: { name: 'Times New Roman', size: 10 } },
  },
  rules: {
    citationBracketStyle: 'square',
    quotedAuthoritySize: 10,
    normalisationOptIn: true,
    protectOrientationSectionBreaks: true,
  },
};

/** UK Court Filing: TNR 12, 1.5-spaced, 25mm margins. */
export const UK_COURT_FILING: StyleProfile = {
  id: 'uk-court-filing',
  name: 'UK Court Filing',
  jurisdiction: 'UK',
  kind: 'court',
  locked: true,
  page: {
    margins: { top: 70.87, bottom: 70.87, left: 70.87, right: 70.87 },
    size: 'A4',
    orientation: 'portrait',
  },
  body: {
    font: { name: 'Times New Roman', size: 12, color: '#000000', bold: false, italic: false },
    spacing: { lineSpacingRule: 'onePtFive', lineSpacing: 1.5, spaceBefore: 0, spaceAfter: 0 },
    indent: { firstLine: 0, left: 0, right: 0 },
  },
  styles: {
    'Normal': {
      font: { name: 'Times New Roman', size: 12 },
      paragraphFormat: { lineSpacingRule: 'onePtFive', lineSpacing: 1.5 },
    },
    'Footnote Text': { font: { name: 'Times New Roman', size: 10 } },
  },
  rules: {
    normalisationOptIn: true,
    protectOrientationSectionBreaks: true,
  },
};

/** Firm Commercial: TNR 11, single-spaced + 6pt spaceAfter, 25.4mm margins. */
export const FIRM_COMMERCIAL: StyleProfile = {
  id: 'firm-commercial',
  name: 'Firm Commercial',
  jurisdiction: 'GENERIC',
  kind: 'commercial',
  locked: true,
  page: {
    margins: { top: 72, bottom: 72, left: 72, right: 72 },
    size: 'A4',
    orientation: 'portrait',
  },
  body: {
    font: { name: 'Times New Roman', size: 11, color: '#000000', bold: false, italic: false },
    spacing: { lineSpacingRule: 'single', lineSpacing: 1.0, spaceBefore: 0, spaceAfter: 6 },
    indent: { firstLine: 0, left: 0, right: 0 },
  },
  styles: {
    'Normal': {
      font: { name: 'Times New Roman', size: 11 },
      paragraphFormat: { lineSpacingRule: 'single', spaceAfter: 6 },
    },
  },
  rules: {
    normalisationOptIn: false,
    protectOrientationSectionBreaks: true,
  },
};

/** All shipped profiles, keyed by id. */
const SHIPPED_PROFILES: Readonly<Record<string, StyleProfile>> = {
  'sg-court-filing': SG_COURT_FILING,
  'uk-court-filing': UK_COURT_FILING,
  'firm-commercial': FIRM_COMMERCIAL,
};

/** Retrieve a shipped profile by id. Returns undefined if not found. */
export function getProfile(id: string): StyleProfile | undefined {
  return SHIPPED_PROFILES[id];
}

/** All available profile ids. */
export function listProfileIds(): string[] {
  return Object.keys(SHIPPED_PROFILES);
}

/** Unit hint — pass when your margin values are in mm; loader converts to pt. */
export type MarginUnit = 'pt' | 'mm';

/**
 * Clone a profile for user editing. Cloned profiles are never locked.
 * Does NOT register the clone — callers store it themselves.
 */
export function cloneProfile(source: StyleProfile, newId: string, newName: string): StyleProfile {
  const clone: StyleProfile = JSON.parse(JSON.stringify(source));
  clone.id = newId;
  clone.name = newName;
  clone.locked = false;
  return clone;
}

/**
 * Convert margin values from mm to pt in a profile copy.
 * Returns a new profile object; does not mutate the source.
 */
export function convertMarginsToPoints(profile: StyleProfile, unit: MarginUnit): StyleProfile {
  if (unit === 'pt') return profile;
  const factor = MM_TO_PT;
  const m = profile.page.margins;
  return {
    ...profile,
    page: {
      ...profile.page,
      margins: {
        top: m.top * factor,
        bottom: m.bottom * factor,
        left: m.left * factor,
        right: m.right * factor,
      },
    },
  };
}
