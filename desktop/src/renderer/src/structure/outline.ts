import { readDocx, type DocxPara } from './read-docx.js';
import { analyseStructure } from './analyse.js';
import type { ContractStructure, NodeKind, StructureNode } from './model.js';
import { typedNumber } from './labels.js';

/**
 * The structure of one document in a form that crosses the worker boundary and
 * reads well when copied: one entry per node, in document order, with its depth,
 * its citation and a short snippet of its first paragraph.
 *
 * Built in the compare worker from the document's own bytes, beside the redline.
 * It never feeds back into the comparison: the redline shows every change
 * whatever this analysis makes of the document (rule 4).
 */

export interface OutlineEntry {
  readonly id: string;
  readonly kind: NodeKind;
  readonly depth: number;
  readonly parentId: string | null;
  /** What a lawyer would cite: "14.10(b)", "Schedule 2 Part A para 3(a)", "Recital B". */
  readonly citation: string | null;
  readonly heading: string | null;
  readonly term: string | null;
  /** Paragraph range in the analyser's own reading of the document (0-based, inclusive). */
  readonly start: number;
  readonly end: number;
  readonly snippet: string;
}

export interface StructureOutline {
  readonly entries: readonly OutlineEntry[];
  /** The citation for each paragraph of the analyser's reading: its node's, or the nearest ancestor's. */
  readonly paragraphCitations: readonly (string | null)[];
  /** Each paragraph's text, normalised for matching against the engine's paragraph list. */
  readonly paragraphKeys: readonly string[];
  readonly counts: Readonly<Record<string, number>>;
}

const SNIPPET = 120;

export function buildOutline(bytes: ArrayBuffer | Uint8Array): StructureOutline {
  const read = readDocx(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
  return outlineOf(read.paragraphs, analyseStructure(read.paragraphs));
}

export function outlineOf(paras: readonly DocxPara[], structure: ContractStructure): StructureOutline {
  const byId = new Map(structure.nodes.map((n) => [n.id, n]));
  const depthOf = (n: StructureNode): number => {
    let d = 0;
    for (let p = n.parentId; p; p = byId.get(p)?.parentId ?? null) d += 1;
    return d;
  };
  const citationOf = (n: StructureNode | undefined): string | null => {
    for (let x = n; x; x = x.parentId ? byId.get(x.parentId) : undefined) {
      if (x.fullLabel) return x.fullLabel;
    }
    return null;
  };

  const entries: OutlineEntry[] = [...structure.nodes]
    .sort((a, b) => a.start - b.start)
    .map((n) => ({
      id: n.id,
      kind: n.kind,
      depth: depthOf(n),
      parentId: n.parentId,
      citation: n.fullLabel,
      heading: n.heading,
      term: n.term,
      start: n.start,
      end: n.end,
      snippet: snippetFrom(paras, n),
    }));

  const counts: Record<string, number> = {};
  for (const n of structure.nodes) counts[n.kind] = (counts[n.kind] ?? 0) + 1;

  return {
    entries,
    paragraphCitations: structure.paragraphs.map((p) => {
      const node = p.nodeId ? byId.get(p.nodeId) : undefined;
      if (p.fullLabel) return p.fullLabel;
      return citationOf(node);
    }),
    paragraphKeys: paras.map((p) => matchKey(p.text)),
    counts,
  };
}

function snippetFrom(paras: readonly DocxPara[], n: StructureNode): string {
  for (let i = n.start; i <= n.end && i < paras.length; i += 1) {
    const raw = paras[i]!.text;
    const t = (i === n.start ? (typedNumber(raw)?.rest ?? raw) : raw).replace(/[\s\u00a0]+/g, ' ').trim();
    if (t !== '') return t.length > SNIPPET ? `${t.slice(0, SNIPPET - 1)}…` : t;
  }
  return '';
}

/**
 * Text reduced far enough that the engine's normalised paragraph and ours agree:
 * no numbering prefix, quotes and dashes folded, case and punctuation dropped.
 */
export function matchKey(text: string): string {
  const typed = typedNumber(text);
  return (typed ? typed.rest : text)
    .toLowerCase()
    .replace(/[‘’“”"'`]/g, '')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
    .slice(0, 160);
}

// ---------------------------------------------------------------------------
// Copyable text
// ---------------------------------------------------------------------------

const KIND_NAMES: Partial<Record<NodeKind, string>> = {
  front_matter: 'Front matter',
  parties: 'Parties',
  execution: 'Execution',
  definitions: 'Definitions',
};

/** One line's label: the citation, else a name for the kind. */
export function entryTitle(e: OutlineEntry): string {
  if (e.kind === 'definition') return e.term ? `“${e.term}”` : 'Definition';
  if (e.citation) return e.citation;
  if (e.kind === 'heading') return e.heading ?? 'Heading';
  return KIND_NAMES[e.kind] ?? e.kind;
}

/** What follows the label: a heading, else the opening words. */
export function entryDetail(e: OutlineEntry): string {
  if (e.kind === 'heading') return '';
  if (e.kind === 'definition') return '';
  if (e.heading) return e.heading;
  if (e.kind === 'front_matter' || e.kind === 'execution' || e.kind === 'parties') return '';
  return e.snippet;
}

export function outlineToMarkdown(outline: StructureOutline, name: string): string {
  const lines = [
    `# Contract structure — ${name}`,
    '',
    '_Read from the document by Mutatis’s structure analyser. Heuristic: check a citation ' +
      'against the document before relying on it._',
    '',
  ];
  for (const e of outline.entries) {
    const detail = entryDetail(e);
    const title = e.kind === 'heading' ? `_${entryTitle(e)}_` : `**${entryTitle(e)}**`;
    lines.push(`${'  '.repeat(e.depth)}- ${title}${detail ? ` — ${detail}` : ''}`);
  }
  if (outline.entries.length === 0) lines.push('_No structure found._');
  return lines.join('\n') + '\n';
}

export function outlineToJson(outline: StructureOutline, name: string): string {
  return (
    JSON.stringify(
      {
        document: name,
        counts: outline.counts,
        nodes: outline.entries.map((e) => ({
          id: e.id,
          kind: e.kind,
          parent: e.parentId,
          depth: e.depth,
          citation: e.citation,
          heading: e.heading,
          term: e.term,
          paragraphs: [e.start, e.end],
          snippet: e.snippet,
        })),
      },
      null,
      2,
    ) + '\n'
  );
}
