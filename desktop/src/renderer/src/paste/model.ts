/**
 * The format-neutral shape pasted text is reduced to.
 *
 * Three inputs (clipboard HTML, Markdown, plain text) all become `Block[]`, and
 * that is the ONLY thing the rest of the feature knows about: the preview renders
 * it as React elements and the docx builder serialises it. Pasted HTML is never
 * put in the DOM — it can come from opposing counsel's document, so it is read
 * for text and a handful of formatting facts and then discarded.
 */

export interface Run {
  readonly text: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly strike?: boolean;
  readonly sup?: boolean;
  readonly sub?: boolean;
  /** A line break inside the paragraph; `text` is ignored. */
  readonly br?: boolean;
  /** A tab; `text` is ignored. */
  readonly tab?: boolean;
}

export type Align = 'left' | 'center' | 'right' | 'both';

export interface Para {
  readonly kind: 'p';
  readonly runs: readonly Run[];
  /** 1-6 for a heading. */
  readonly heading?: number;
  readonly align?: Align;
  /** Left indent in twentieths of a point. */
  readonly indent?: number;
  /** True for a paragraph that carries a list marker in its text. */
  readonly list?: boolean;
}

export interface Table {
  readonly kind: 'table';
  /** rows → cells → paragraphs */
  readonly rows: readonly (readonly (readonly Para[])[])[];
}

export type Block = Para | Table;

export type PasteFormat = 'formatted' | 'markdown' | 'plain';

export function runIsEmpty(run: Run): boolean {
  return !run.br && !run.tab && run.text === '';
}

export function paraText(p: Para): string {
  return p.runs.map((r) => (r.br ? '\n' : r.tab ? '\t' : r.text)).join('');
}

export function blocksText(blocks: readonly Block[]): string {
  const out: string[] = [];
  for (const b of blocks) {
    if (b.kind === 'p') out.push(paraText(b));
    else for (const row of b.rows) out.push(row.map((cell) => cell.map(paraText).join(' ')).join('\t'));
  }
  return out.join('\n');
}

export function hasContent(blocks: readonly Block[]): boolean {
  return blocksText(blocks).trim() !== '';
}

export function wordCount(blocks: readonly Block[]): number {
  const text = blocksText(blocks).trim();
  return text === '' ? 0 : text.split(/\s+/).length;
}

/** Does this carry anything a plain-text paste would have lost? */
export function hasFormatting(blocks: readonly Block[]): boolean {
  const para = (p: Para): boolean =>
    p.heading !== undefined ||
    p.list === true ||
    p.align !== undefined ||
    p.runs.some((r) => r.bold || r.italic || r.underline || r.strike || r.sup || r.sub);
  return blocks.some((b) => (b.kind === 'p' ? para(b) : true));
}
