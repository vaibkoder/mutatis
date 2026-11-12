/**
 * Shared lexicons for the Defined Terms checker's used-but-not-defined allowlist.
 *
 * These arrays are **static English-language / global reference data**, not
 * per-matter configuration: months, weekdays, country and city names, company
 * forms, legal boilerplate, the definitional-verb set, etc. They are identical
 * across jurisdictions and therefore live in ONE place, composed into each
 * `JurisdictionProfile` rather than copied inline — see the reuse-before-build rule ("two
 * implementations of the same thing is a defect") and ADR-0015.
 *
 * Genuinely jurisdiction-specific entries (SG vs UK regulators / institutions)
 * are kept in their own constants below and spread on top of the shared base.
 * Per-matter tuning is still possible: a matter may extend any profile's
 * exclusion arrays at runtime; this module only supplies the baseline.
 */

// ---------------------------------------------------------------------------
// Definition recognition (shared — ADR-0013 verb set)
// ---------------------------------------------------------------------------

export const DEFINITION_VERBS: readonly string[] = [
  'means', 'shall mean', 'means and includes', 'shall mean and include',
  'includes', 'refers to',
  'has the meaning given to it in', 'has the meaning given in',
  'has the meaning set out in', 'has the meaning specified in',
  'has the meaning ascribed to it in',
  'shall have the meaning given to it in', 'shall have the meaning given in',
  'shall have the meaning set out in', 'shall have the meaning specified in',
  'shall have the meaning ascribed to it in',
  'is defined in',
];

export const DETERMINERS: readonly string[] = [
  'the', 'a', 'an', 'this', 'that', 'these', 'those',
  'such', 'its', 'their', 'any', 'each', 'no',
];

// ---------------------------------------------------------------------------
// Exclusion allowlist — shared lexical baseline
// ---------------------------------------------------------------------------

export const PARTY_ROLES: readonly string[] = [
  'Buyer', 'Seller', 'Purchaser', 'Vendor', 'Parties', 'Party',
  'Disclosing Party', 'Receiving Party', 'Lender', 'Borrower', 'Guarantor',
  'Company', 'Shareholder', 'Shareholders', 'Director', 'Directors',
  // Financing / security / agency roles (additive — ADR-0016 / feature 09).
  // SG/UK commercial roles only; no US-specific roles.
  'Agent', 'Trustee', 'Chargor', 'Mortgagor', 'Sponsor', 'Obligor',
  'Indemnifier', 'Beneficiary',
];

export const CONVENTIONAL_TERMS: readonly string[] = [
  'Board', 'Business Day', 'Business Days', 'Working Day', 'Working Days',
  'Schedule', 'Clause', 'Recital', 'Annex', 'Appendix', 'Part',
  'Group', 'Effective Date', 'Completion', 'Closing',
];

/** Tax / data / corporate acronyms common to SG/UK commercial practice. */
export const COMMON_ACRONYMS: readonly string[] = [
  'VAT', 'GST', 'PDPA', 'IRAS', 'ACRA', 'CPF',
];

/** Company entity suffixes — appear capitalised inside party names, never bespoke terms. */
export const COMPANY_FORM_ACRONYMS: readonly string[] = [
  'LLC', 'LP', 'GP', 'PC', 'SA', 'NV', 'BV', 'AG', 'PTY', 'SDN', 'BHD', 'PTE', 'SARL',
];

/** Legal/contractual boilerplate all-caps (e.g. "NOW, THEREFORE, in consideration of …"). */
export const LEGAL_BOILERPLATE_ACRONYMS: readonly string[] = [
  'THEREFORE', 'WHEREAS', 'WITNESSETH', 'HEREBY', 'HEREWITH', 'THEREOF',
  'THEREIN', 'THERETO', 'THEREWITH', 'WHEREOF', 'PURSUANT', 'NOW',
];

export const STATUTORY_WORDS: readonly string[] = [
  'the Act', 'the Regulations', 'the Companies Act',
];

export const STATUTORY_CITATION: RegExp =
  /\b(?:[A-Z][A-Za-z]*\s+)*(?:Act|Regulations)\b(?:\s*\(Amendment\))?(?:\s+\d{4})?/;

/**
 * Proper nouns that are never bespoke defined terms: months, weekdays, country
 * and territory names, major commercial cities, geographic regions, nationality
 * / language adjectives, and bare court references. Shared by both jurisdictions.
 */
export const UNIVERSAL_PROPER_NOUNS: readonly string[] = [
  // Months
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
  // Days
  'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday',
  // Countries / territories — Asia-Pacific
  'Singapore', 'Malaysia', 'Indonesia', 'Thailand', 'Philippines', 'Vietnam',
  'Cambodia', 'Myanmar', 'Brunei', 'Laos', 'Japan', 'China', 'Hong Kong',
  'Macau', 'Macao', 'Taiwan', 'South Korea', 'Korea', 'India', 'Bangladesh',
  'Sri Lanka', 'Pakistan', 'Nepal', 'Australia', 'New Zealand',
  // British Isles & Europe
  'United Kingdom', 'England', 'Wales', 'England and Wales', 'Scotland', 'Northern Ireland',
  'Britain', 'Great Britain', 'Ireland',
  'France', 'Germany', 'Netherlands', 'Belgium', 'Luxembourg',
  'Switzerland', 'Austria', 'Spain', 'Italy', 'Portugal',
  'Sweden', 'Denmark', 'Norway', 'Finland', 'Poland',
  'Czech Republic', 'Hungary', 'Romania', 'Bulgaria', 'Greece', 'Cyprus',
  'European Union',
  // Americas
  'United States', 'United States of America', 'Canada',
  // Middle East & Africa
  'United Arab Emirates', 'UAE', 'Dubai', 'Abu Dhabi',
  'Qatar', 'Bahrain', 'Kuwait', 'Saudi Arabia',
  'South Africa', 'Nigeria', 'Kenya',
  // Other
  'Israel', 'Turkey', 'Brazil', 'Mexico',
  // Major commercial cities
  'London', 'New York', 'Tokyo', 'Beijing', 'Shanghai', 'Shenzhen',
  'Seoul', 'Sydney', 'Melbourne', 'Toronto', 'Paris', 'Frankfurt',
  'Amsterdam', 'Zurich', 'Dublin', 'Mumbai', 'Delhi', 'Bangalore',
  // Geographic regions
  'Asia', 'Europe', 'Americas', 'Africa', 'Oceania',
  'Middle East', 'Southeast Asia', 'Asia Pacific', 'Asia-Pacific',
  'East Asia', 'South Asia', 'Western Europe', 'Eastern Europe', 'Pacific',
  // Nationality / language adjectives
  'Singaporean', 'British', 'English', 'Welsh', 'Scottish', 'Irish',
  'Malaysian', 'Indonesian', 'Thai', 'Filipino', 'Vietnamese',
  'Cambodian', 'Japanese', 'Chinese', 'Korean', 'Indian', 'Australian',
  'New Zealander', 'Canadian', 'American', 'French', 'German', 'Dutch',
  'Belgian', 'Swiss', 'Austrian', 'Spanish', 'Italian', 'Portuguese',
  'Swedish', 'Danish', 'Norwegian', 'Finnish', 'Polish', 'Greek',
  'European', 'Asian', 'Malay', 'Tamil', 'Mandarin', 'Cantonese',
  // Common institutional/court references
  'Court', 'Courts',
];

// ---------------------------------------------------------------------------
// Jurisdiction-specific entries (the genuine SG ≠ UK split)
// ---------------------------------------------------------------------------

/** Singapore regulators / statutory bodies that appear capitalised but undefined. */
export const SG_REGULATOR_ACRONYMS: readonly string[] = [
  'MAS', 'SGX', 'MOM', 'IMDA', 'EDB', 'HDB', 'PSA',
];

export const SG_INSTITUTIONS: readonly string[] = [
  'Monetary Authority of Singapore', 'Singapore Exchange',
  'Inland Revenue Authority of Singapore',
  'Accounting and Corporate Regulatory Authority',
  'Economic Development Board',
];

/** UK (England & Wales) regulators / statutory bodies. */
export const UK_REGULATOR_ACRONYMS: readonly string[] = [
  'HMRC', 'FCA', 'PRA', 'PAYE', 'NIC', 'FRC', 'FOS',
];

export const UK_INSTITUTIONS: readonly string[] = [
  'Companies House', 'HM Land Registry', 'Bank of England',
  'Financial Conduct Authority', 'Prudential Regulation Authority',
];

// ---------------------------------------------------------------------------
// Composed per-jurisdiction allowlists
// ---------------------------------------------------------------------------

const BASE_ACRONYMS: readonly string[] = [
  ...COMMON_ACRONYMS, ...COMPANY_FORM_ACRONYMS, ...LEGAL_BOILERPLATE_ACRONYMS,
];

export const SG_ACRONYMS: readonly string[] = [...BASE_ACRONYMS, ...SG_REGULATOR_ACRONYMS];
export const UK_ACRONYMS: readonly string[] = [...BASE_ACRONYMS, ...UK_REGULATOR_ACRONYMS];

export const SG_PROPER_NOUNS: readonly string[] = [...UNIVERSAL_PROPER_NOUNS, ...SG_INSTITUTIONS];
export const UK_PROPER_NOUNS: readonly string[] = [...UNIVERSAL_PROPER_NOUNS, ...UK_INSTITUTIONS];
