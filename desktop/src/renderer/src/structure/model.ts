/**
 * The contract-structure analyser's output: a clause tree over the document's
 * paragraphs, in which every non-empty paragraph has exactly one owner.
 *
 * The node kinds and the citation rules follow the contract-structure benchmark
 * the analyser is graded against (front matter, parties, recitals, headings,
 * clauses and subclauses, definitions, schedules and their parts, execution),
 * because those are the units a lawyer actually cites and navigates by.
 */

export type NodeKind =
  | 'front_matter'
  | 'parties'
  | 'recital'
  | 'heading'
  | 'clause'
  | 'subclause'
  | 'definitions'
  | 'definition'
  | 'schedule'
  | 'part'
  | 'annex'
  | 'appendix'
  | 'execution'
  | 'other';

export type Region = 'front' | 'body' | 'schedule' | 'execution' | 'back';

/** Where a paragraph's number came from: Word's list numbering, typed text, or none. */
export type LabelSource = 'word' | 'typed' | 'none';

export interface StructureNode {
  readonly id: string;
  kind: NodeKind;
  /** The number as painted or typed at this level ("4.", "(b)", "Schedule 2"). */
  label: string | null;
  /** The citation a lawyer would write ("14.10(b)(ii)", "Schedule 2 Part A para 3(a)"). */
  fullLabel: string | null;
  heading: string | null;
  /** Defined term, without quotes, for a definition (or a numbered paragraph that defines one). */
  term: string | null;
  parentId: string | null;
  readonly region: Region;
  readonly start: number;
  end: number;
  readonly numberingSource: LabelSource;
  /** A number painted in a table row's empty first cell: it starts its citation afresh. */
  readonly rowNumber?: boolean;
}

export interface ParagraphStructure {
  readonly label: string | null;
  readonly fullLabel: string | null;
  readonly labelSource: LabelSource;
  /** The node that owns this paragraph; null only for empty paragraphs before any node. */
  readonly nodeId: string | null;
}

export interface ContractStructure {
  readonly nodes: readonly StructureNode[];
  readonly paragraphs: readonly ParagraphStructure[];
}
