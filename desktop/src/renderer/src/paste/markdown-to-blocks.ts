import { unified } from 'unified';
import remarkParse from 'remark-parse';
import remarkGfm from 'remark-gfm';
import type { PhrasingContent, RootContent, Table as MdTable, List } from 'mdast';
import type { Block, Para, Run } from './model.js';

/**
 * Markdown → blocks, through the same parser the chat renderer already uses.
 *
 * Lists become literal markers in the paragraph text ("• ", "1. ") rather than
 * Word numbering: a pasted comparison should show a renumbered list as the text
 * change it is, and real numbering definitions are a lot of machinery for text
 * that was never in a numbered document. Links keep their text only; raw HTML
 * is kept as text, never interpreted.
 */

type Marks = Pick<Run, 'bold' | 'italic' | 'strike'>;

const INDENT_PER_LEVEL = 720;

function inline(nodes: readonly PhrasingContent[], marks: Marks = {}): Run[] {
  const out: Run[] = [];
  for (const node of nodes) {
    switch (node.type) {
      case 'text':
        out.push({ text: node.value, ...marks });
        break;
      case 'strong':
        out.push(...inline(node.children, { ...marks, bold: true }));
        break;
      case 'emphasis':
        out.push(...inline(node.children, { ...marks, italic: true }));
        break;
      case 'delete':
        out.push(...inline(node.children, { ...marks, strike: true }));
        break;
      case 'inlineCode':
        out.push({ text: node.value, ...marks });
        break;
      case 'link':
      case 'linkReference':
        out.push(...inline(node.children, marks));
        break;
      case 'image':
      case 'imageReference':
        if (node.alt) out.push({ text: node.alt, ...marks });
        break;
      case 'break':
        out.push({ text: '', br: true });
        break;
      case 'html':
        out.push({ text: node.value, ...marks });
        break;
      case 'footnoteReference':
        out.push({ text: `[${node.label ?? node.identifier}]`, ...marks });
        break;
    }
  }
  return out;
}

function tableBlock(node: MdTable): Block {
  return {
    kind: 'table',
    rows: node.children.map((row, rowIndex) =>
      row.children.map((cell) => {
        const runs = inline(cell.children, rowIndex === 0 ? { bold: true } : {});
        return [{ kind: 'p' as const, runs }];
      }),
    ),
  };
}

function listBlocks(node: List, level: number, out: Block[]): void {
  let n = node.start ?? 1;
  for (const item of node.children) {
    const marker = node.ordered ? `${n}. ` : '• ';
    n += 1;
    let first = true;
    for (const child of item.children) {
      if (child.type === 'list') {
        listBlocks(child, level + 1, out);
        continue;
      }
      const start = out.length;
      blocks([child], level + 1, out);
      if (first && out[start]?.kind === 'p') {
        const p = out[start] as Para;
        out[start] = { ...p, runs: [{ text: marker }, ...p.runs], list: true };
      }
      first = false;
    }
    if (first) out.push({ kind: 'p', runs: [{ text: marker.trimEnd() }], list: true, indent: INDENT_PER_LEVEL * (level + 1) });
  }
}

function blocks(nodes: readonly RootContent[], level: number, out: Block[]): void {
  const indent = level > 0 ? { indent: INDENT_PER_LEVEL * level } : {};
  for (const node of nodes) {
    switch (node.type) {
      case 'heading':
        out.push({ kind: 'p', runs: inline(node.children), heading: node.depth, ...indent });
        break;
      case 'paragraph':
        out.push({ kind: 'p', runs: inline(node.children), ...indent });
        break;
      case 'list':
        listBlocks(node, level, out);
        break;
      case 'blockquote':
        blocks(node.children, level + 1, out);
        break;
      case 'code':
        for (const line of node.value.split('\n')) {
          out.push({ kind: 'p', runs: line === '' ? [] : [{ text: line }], ...indent });
        }
        break;
      case 'table':
        out.push(tableBlock(node));
        break;
      case 'thematicBreak':
        out.push({ kind: 'p', runs: [{ text: '———' }], align: 'center' });
        break;
      case 'html':
        out.push({ kind: 'p', runs: [{ text: node.value }], ...indent });
        break;
      case 'definition':
      case 'footnoteDefinition':
      case 'yaml':
        break;
      default:
        break;
    }
  }
}

export function markdownToBlocks(source: string): Block[] {
  const tree = unified().use(remarkParse).use(remarkGfm).parse(source.replace(/\r\n?/g, '\n'));
  const out: Block[] = [];
  blocks(tree.children, 0, out);
  return out;
}

/**
 * Does this plain text look like Markdown on purpose? Deliberately conservative:
 * a false positive silently turns "* * *" or "1. Definitions" typed in a contract
 * into formatting, so it needs a clear signal.
 */
export function looksLikeMarkdown(text: string): boolean {
  return (
    /^#{1,6}\s+\S/m.test(text) ||
    /\*\*[^*\n]+\*\*/.test(text) ||
    /(^|[^_\w])__[^_\n]+__([^_\w]|$)/.test(text) ||
    /^\s*\|.+\|\s*\n\s*\|?\s*:?-{3,}:?\s*\|/m.test(text) ||
    /^\s*[-*+]\s+\S/m.test(text) && /\n\s*[-*+]\s+\S/.test(text)
  );
}
