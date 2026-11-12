/**
 * Feature 07-A — Table of Contents: switch derivation.
 *
 * `buildTocSwitches(profile, opts)` is a PURE function — no Office JS, no
 * storage, no side-effects. It derives the Word TOC field-code switch string
 * from a TocProfile according to spec §A.4.1.
 *
 * Switch derivation rules (canonical):
 *  \o "1-3"   — ONLY when styleLevelMap is empty
 *  \t "…"     — when styleLevelMap is non-empty (supersedes \o)
 *  \h         — when useHyperlinks is true
 *  \b Name    — when opts.bookmark is provided (scoped TOC)
 *  \n "1-3"   — when hidePageNumbers is true
 *
 * Tab leaders are NOT a field switch — they are applied via Word TOC styles.
 * This function therefore never emits tab-leader directives.
 *
 * Spec reference: spec/features/07-toc-precedent-templates.md §A.4.1.
 */

import type { TocProfile, BuildSwitchesOptions } from './types.js';

/**
 * Build the TOC field-code switch string for a given profile.
 *
 * The returned string is the portion that follows "TOC" in the field code,
 * e.g. ` \o "1-3" \h` or ` \t "Heading 1,1;Heading 2,2" \h`.
 * The caller is responsible for prepending "TOC" and setting `field.code`.
 *
 * @param profile  Active TOC profile (firm default or doc override).
 * @param opts     Optional modifiers (bookmark for scoped TOCs).
 * @returns        Switch string beginning with a space, e.g. ` \o "1-3" \h`.
 */
export function buildTocSwitches(
  profile: TocProfile,
  opts: BuildSwitchesOptions = {},
): string {
  const parts: string[] = [];

  // Style/level mapping: \t supersedes \o when styleLevelMap is non-empty.
  if (profile.styleLevelMap.length > 0) {
    // \t "Style1,Level;Style2,Level;…"
    const mapping = profile.styleLevelMap
      .map(entry => `${entry.style},${entry.level}`)
      .join(';');
    parts.push(`\\t "${mapping}"`);
  } else {
    // \o "1-3" — standard heading-outline levels.
    parts.push(`\\o "${profile.switches.levels}"`);
  }

  // \h — hyperlink entries.
  if (profile.switches.useHyperlinks) {
    parts.push('\\h');
  }

  // \b BookmarkName — scoped TOC (multiple-TOC case).
  if (opts.bookmark) {
    parts.push(`\\b ${opts.bookmark}`);
  }

  // \n "levels" — hide page numbers.
  if (profile.switches.hidePageNumbers) {
    parts.push(`\\n "${profile.switches.levels}"`);
  }

  // Join with spaces; leading space ensures "TOC" + switches is well-formed.
  return ' ' + parts.join(' ');
}

// ---------------------------------------------------------------------------
// Schedule Index switch string (Strategy 1)
// ---------------------------------------------------------------------------

/**
 * Build the switch string for an Index of Schedules field (Strategy 1).
 *
 * Uses ONLY \t for the supplied schedule styles and \h (no \o — body heading
 * styles must be excluded). This is Strategy 1 from spec §A.5.4.
 *
 * @param scheduleStyles  Named styles to include, e.g. ["Schedule Heading","Schedule Title"].
 * @param level           TOC level to assign each style (typically 1).
 * @param useHyperlinks   Whether to emit \h (defaults to true).
 * @returns               Switch string, e.g. ` \t "Schedule Heading,1;Schedule Title,1" \h`.
 */
export function buildScheduleIndexSwitches(
  scheduleStyles: string[],
  level: number,
  useHyperlinks = true,
): string {
  if (scheduleStyles.length === 0) {
    throw new Error('buildScheduleIndexSwitches: at least one schedule style must be provided');
  }

  const parts: string[] = [];

  // \t only — deliberately NO \o so body headings are excluded.
  const mapping = scheduleStyles.map(s => `${s},${level}`).join(';');
  parts.push(`\\t "${mapping}"`);

  // \h — hyperlink entries.
  if (useHyperlinks) {
    parts.push('\\h');
  }

  return ' ' + parts.join(' ');
}
