/**
 * tests/unit/toc.switches.test.ts
 *
 * Unit tests for A1: buildTocSwitches() and buildScheduleIndexSwitches().
 * These are pure functions — no mocking needed.
 *
 * Acceptance checks for the TOC feature:
 *  - \o "1-3" emitted ONLY when styleLevelMap is empty
 *  - \t emitted from styleLevelMap when non-empty (supersedes \o)
 *  - \h emitted when useHyperlinks is true
 *  - \b BookmarkName emitted when opts.bookmark is provided
 *  - \n "1-3" emitted when hidePageNumbers is true
 *  - Schedule index \t excludes \o (body headings excluded)
 */

import { describe, it, expect } from 'vitest';
import { buildTocSwitches, buildScheduleIndexSwitches } from '@features/content/toc/switches.js';
import type { TocProfile } from '@features/content/toc/types.js';
import { HOUSE_DEFAULT_PROFILE } from '@features/content/toc/config.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makeProfile(overrides: Partial<TocProfile> = {}): TocProfile {
  return {
    id: 'test',
    label: 'Test',
    switches: {
      levels: '1-3',
      useHyperlinks: true,
      hidePageNumbers: false,
      rightAlignPageNo: true,
      preserveTabAndLeader: true,
    },
    styleLevelMap: [],
    tabLeader: 'dots',
    ...overrides,
  };
}

// ---------------------------------------------------------------------------
// buildTocSwitches — \o rule
// ---------------------------------------------------------------------------

describe('buildTocSwitches — \\o vs \\t', () => {
  it('emits \\o "1-3" when styleLevelMap is empty', () => {
    const result = buildTocSwitches(makeProfile({ styleLevelMap: [] }));
    expect(result).toContain('\\o "1-3"');
    expect(result).not.toContain('\\t');
  });

  it('does NOT emit \\o when styleLevelMap is non-empty', () => {
    const result = buildTocSwitches(makeProfile({
      styleLevelMap: [{ style: 'Heading 1', level: 1 }],
    }));
    expect(result).not.toContain('\\o');
  });

  it('emits \\t "Style,Level" when styleLevelMap has one entry', () => {
    const result = buildTocSwitches(makeProfile({
      styleLevelMap: [{ style: 'Heading 1', level: 1 }],
    }));
    expect(result).toContain('\\t "Heading 1,1"');
  });

  it('emits \\t "Style1,L1;Style2,L2" when styleLevelMap has multiple entries', () => {
    const result = buildTocSwitches(makeProfile({
      styleLevelMap: [
        { style: 'Heading 1', level: 1 },
        { style: 'Heading 2', level: 2 },
        { style: 'Clause Heading', level: 1 },
      ],
    }));
    expect(result).toContain('\\t "Heading 1,1;Heading 2,2;Clause Heading,1"');
  });
});

// ---------------------------------------------------------------------------
// buildTocSwitches — \h rule
// ---------------------------------------------------------------------------

describe('buildTocSwitches — \\h (hyperlinks)', () => {
  it('emits \\h when useHyperlinks is true', () => {
    const result = buildTocSwitches(makeProfile({ switches: {
      levels: '1-3', useHyperlinks: true, hidePageNumbers: false,
      rightAlignPageNo: true, preserveTabAndLeader: true,
    } }));
    expect(result).toContain('\\h');
  });

  it('does NOT emit \\h when useHyperlinks is false', () => {
    const result = buildTocSwitches(makeProfile({ switches: {
      levels: '1-3', useHyperlinks: false, hidePageNumbers: false,
      rightAlignPageNo: true, preserveTabAndLeader: true,
    } }));
    expect(result).not.toContain('\\h');
  });
});

// ---------------------------------------------------------------------------
// buildTocSwitches — \b rule
// ---------------------------------------------------------------------------

describe('buildTocSwitches — \\b (scoped TOC bookmark)', () => {
  it('does NOT emit \\b when no bookmark is passed', () => {
    const result = buildTocSwitches(makeProfile());
    expect(result).not.toContain('\\b');
  });

  it('emits \\b BookmarkName when opts.bookmark is provided', () => {
    const result = buildTocSwitches(makeProfile(), { bookmark: 'TOC_SCOPE_part1_1' });
    expect(result).toContain('\\b TOC_SCOPE_part1_1');
  });
});

// ---------------------------------------------------------------------------
// buildTocSwitches — \n rule
// ---------------------------------------------------------------------------

describe('buildTocSwitches — \\n (hide page numbers)', () => {
  it('does NOT emit \\n when hidePageNumbers is false', () => {
    const result = buildTocSwitches(makeProfile());
    expect(result).not.toContain('\\n');
  });

  it('emits \\n "1-3" when hidePageNumbers is true', () => {
    const result = buildTocSwitches(makeProfile({ switches: {
      levels: '1-3', useHyperlinks: true, hidePageNumbers: true,
      rightAlignPageNo: true, preserveTabAndLeader: true,
    } }));
    expect(result).toContain('\\n "1-3"');
  });
});

// ---------------------------------------------------------------------------
// buildTocSwitches — house-default profile (A-T1 switch coverage)
// ---------------------------------------------------------------------------

describe('buildTocSwitches — HOUSE_DEFAULT_PROFILE (A-T1)', () => {
  it('house-default emits \\t (not \\o) plus \\h', () => {
    const result = buildTocSwitches(HOUSE_DEFAULT_PROFILE);
    expect(result).toContain('\\t');
    expect(result).toContain('\\h');
    expect(result).not.toContain('\\o');
    expect(result).not.toContain('\\n');
  });

  it('house-default \\t includes Heading 1,1 and Clause Heading,1', () => {
    const result = buildTocSwitches(HOUSE_DEFAULT_PROFILE);
    expect(result).toContain('Heading 1,1');
    expect(result).toContain('Clause Heading,1');
  });

  it('house-default scoped insert appends \\b', () => {
    const result = buildTocSwitches(HOUSE_DEFAULT_PROFILE, { bookmark: 'TOC_SCOPE_part1_1' });
    expect(result).toContain('\\b TOC_SCOPE_part1_1');
  });
});

// ---------------------------------------------------------------------------
// buildTocSwitches — result format
// ---------------------------------------------------------------------------

describe('buildTocSwitches — result format', () => {
  it('returns a string beginning with a space (well-formed for field.code = "TOC" + switches)', () => {
    const result = buildTocSwitches(makeProfile());
    expect(result.startsWith(' ')).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// buildScheduleIndexSwitches — A-T4 acceptance check
// ---------------------------------------------------------------------------

describe('buildScheduleIndexSwitches (Index of Schedules, Strategy 1)', () => {
  it('emits \\t for schedule styles only — NO \\o', () => {
    const result = buildScheduleIndexSwitches(
      ['Schedule Heading', 'Schedule Title'], 1,
    );
    expect(result).toContain('\\t "Schedule Heading,1;Schedule Title,1"');
    expect(result).not.toContain('\\o');
  });

  it('includes \\h by default', () => {
    const result = buildScheduleIndexSwitches(['Schedule Heading'], 1);
    expect(result).toContain('\\h');
  });

  it('omits \\h when useHyperlinks is false', () => {
    const result = buildScheduleIndexSwitches(['Schedule Heading'], 1, false);
    expect(result).not.toContain('\\h');
  });

  it('correctly encodes a single schedule style', () => {
    const result = buildScheduleIndexSwitches(['Schedule Heading'], 1);
    expect(result).toContain('\\t "Schedule Heading,1"');
  });

  it('throws when called with an empty styles array', () => {
    expect(() => buildScheduleIndexSwitches([], 1)).toThrow();
  });

  it('body-heading styles are excluded (no \\o emitted)', () => {
    // Verifies A-T4: a separate Index of Schedules field excludes body headings.
    const result = buildScheduleIndexSwitches(['Schedule Heading'], 1);
    // The switch string must not ask Word to also pick up outline-level headings.
    expect(result).not.toContain('\\o');
    // It must address Schedule Heading specifically via \t.
    expect(result).toContain('Schedule Heading');
  });
});
