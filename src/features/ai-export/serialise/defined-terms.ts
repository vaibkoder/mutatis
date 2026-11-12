/**
 * Defined-terms glossary serialiser (feature 11 — spec §4.2, rules F2/F5).
 *
 * Projects the harvested `DefinitionsMap` (+ `UsageIndex` for variants) into a
 * clean glossary. Cross-reference definitions are resolved inline against the
 * Section Index (F2): the referent clause's body text is appended so the LLM
 * never has to chase a pointer. No new scanning — reads in-memory records only.
 */
import type { DefinitionsMap, DefinitionEntry, UsageIndex } from '@features/defined-terms/index.js';
import type { SectionIndex } from '@shared/section-index/index.js';
import type { DefinedTermsExport, TermEntry } from '../types.js';
import type { Citer } from '../cite.js';

export interface DefinedTermsDeps {
  usage: UsageIndex;
  index: SectionIndex | null;
  /** Enclosing-clause citation for an anchor paragraph (F3). */
  cite: Citer;
  /** Full text of a paragraph by id — the definition body lives in the document. */
  paraText: (paragraphId: string) => string;
}

export function serialiseDefinedTerms(map: DefinitionsMap, deps: DefinedTermsDeps): DefinedTermsExport {
  const glossary: TermEntry[] = [];
  for (const entry of map.values()) {
    glossary.push(toTermEntry(entry, deps));
  }
  return { glossary };
}

function toTermEntry(entry: DefinitionEntry, deps: DefinedTermsDeps): TermEntry {
  // Primary record: prefer the first in-document definition over a cross-doc one.
  const primary = entry.definitions.find(d => !d.isCrossDoc) ?? entry.definitions[0]!;

  const baseText = deps.paraText(primary.anchor.paragraphId).trim();
  let definition = baseText;

  // F2 — resolve a same-document cross-reference inline.
  if (primary.isCrossRef && primary.crossRefTarget) {
    const resolved = resolveReferent(primary.crossRefTarget, deps);
    definition = `${baseText || `has the meaning given in ${primary.crossRefTarget}`} ${resolved}`;
  }

  const entryOut: TermEntry = {
    term: entry.canonical,
    definition,
    defined_at: deps.cite(primary.anchor.paragraphId),
    source: primary.source,
  };

  if (entry.definitions.length > 1) entryOut.defined_multiple_times = true;

  const variants = collectVariants(entry, deps.usage);
  if (variants.length > 0) entryOut.variants = variants;

  return entryOut;
}

/** Find the referent section by display label and append its body text (F2). */
function resolveReferent(target: string, deps: DefinedTermsDeps): string {
  const { index, paraText } = deps;
  if (!index) return '[→ target not found in this document]';

  const want = target.trim().toLowerCase();
  const node = index.nodes.find(n => n.displayLabel.trim().toLowerCase() === want);
  if (!node) return '[→ target not found in this document]';

  const body = paraText(node.paraId).trim();
  if (!body) return '[→ target not found in this document]';
  return `[→ "${body}"]`;
}

/** Observed surface variants that differ from the canonical term (F5). */
function collectVariants(entry: DefinitionEntry, usage: UsageIndex): string[] {
  const uses = usage.get(entry.key);
  if (!uses) return [];
  const out = new Set<string>();
  for (const u of uses) {
    const surface = u.surface.trim();
    if (surface && surface !== entry.canonical) out.add(surface);
  }
  return [...out];
}
