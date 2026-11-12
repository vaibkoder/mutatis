/**
 * Defined-terms dictionary builder for Interactive Reference Lookup (feature 04).
 *
 * Composes the existing Defined-Terms harvester (buildDefinitionsMap +
 * mergeTableDefinitions — feature 01) and projects each DefinitionRecord
 * into a DefinedTerm, grouped by key into TermsDictionary.
 *
 * Reuse contract:
 *  - NO new definition-detection logic.  All detection is in the harvester.
 *  - This file ONLY projects DefinitionsMap → TermsDictionary.
 */

import type { ParaNode } from '@shared/doc-model/index.js';
import type { JurisdictionProfile } from '@shared/config/index.js';
import { getLookupConfig } from '@shared/config/index.js';
import {
  buildDefinitionsMap,
  mergeTableDefinitions,
} from '@features/defined-terms/index.js';
import { buildAlternation } from '@features/defined-terms/normalise.js';
import type { DefinedTerm, TermsDictionary } from './types.js';

// ── Build ─────────────────────────────────────────────────────────────────────

/**
 * Build the in-memory lookup dictionary from the current paragraph snapshot.
 * Mirrors spec §5.1 build procedure; called once at task-pane open and after
 * each debounced DocModel invalidation.
 *
 * @param bodyParas      Body + table-cell paragraphs from DocModel.allParagraphs.
 * @param footnoteParas  Footnote paragraphs from DocModel.allParagraphs.
 * @param profile        Active jurisdiction profile.
 */
export function buildTermsDictionary(
  bodyParas: ReadonlyArray<ParaNode>,
  footnoteParas: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
): TermsDictionary {
  // Store up to the *expanded* cap so the expanded card can show a fuller view;
  // collapsed list rows / Online single-preview trim further at render time.
  const { expandedPreviewMaxChars } = getLookupConfig(profile);

  // Build a quick lookup: paragraphId → ParaNode.
  const paraById = new Map<string, ParaNode>();
  for (const p of bodyParas) paraById.set(p.id, p);
  for (const p of footnoteParas) paraById.set(p.id, p);

  // Harvest definitions — reuse feature-01 logic entirely.
  const defMap = buildDefinitionsMap(bodyParas, footnoteParas, profile);
  // Table-cell definitions (tableCell containers live inside bodyParas).
  mergeTableDefinitions(defMap, bodyParas, profile);

  // Project DefinitionsMap → TermsDictionary.
  const byKey = new Map<string, DefinedTerm[]>();

  for (const entry of defMap.values()) {
    const terms: DefinedTerm[] = [];

    for (const record of entry.definitions) {
      const para = paraById.get(record.anchor.paragraphId);

      const defKind: DefinedTerm['defKind'] = record.isCrossDoc
        ? 'external'
        : record.isCrossRef
          ? 'by-cross-ref'
          : 'inline';

      const rawText = para?.text ?? '';
      const previewText = rawText.length > expandedPreviewMaxChars
        ? rawText.slice(0, expandedPreviewMaxChars) + '…'  // …
        : rawText;

      const forwardRef = record.isCrossRef && record.crossRefTarget
        ? record.crossRefTarget
        : undefined;
      terms.push({
        termKey: entry.key,
        display: entry.canonical,
        previewText,
        defKind,
        ...(forwardRef !== undefined ? { forwardRef } : {}),
        sourceParagraphIndex: para?.index ?? -1,
        anchor: record.anchor,
      });
    }

    if (terms.length > 0) {
      byKey.set(entry.key, terms);
    }
  }

  // Pre-build the longest-match-first alternation over every surface form so a
  // bare cursor sitting *inside* a term (not selecting it whole) can still be
  // located within the sentence window (spec §4.2 / T3). Built once here, never
  // on the per-selection lookup critical path.
  const surfaceForms: string[] = [];
  for (const terms of byKey.values()) {
    for (const t of terms) surfaceForms.push(t.display);
  }
  const termMatcher = buildAlternation(surfaceForms);

  return { byKey, dirty: false, termMatcher };
}
