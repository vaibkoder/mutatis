import type { Align, Block, Para, Run, Table } from './model.js';

/**
 * Clipboard HTML (Word, Google Docs, a web page, a rich-text editor) → blocks.
 *
 * Reads what the document says about itself: bold/italic/underline/strike/sup/sub
 * from tags AND from inline styles (Google Docs and Word state bold as
 * `font-weight:700` on a span, and wrap everything in `<b style="font-weight:normal">`),
 * headings, lists, alignment, indent and tables. Colours, fonts and sizes are
 * deliberately not carried: they would show up as "formatting changes" between two
 * pastes that look identical to a reader.
 *
 * The input is parsed with DOMParser, which builds an INERT document — no script
 * runs and no image or stylesheet is fetched — and nothing from it is ever
 * inserted into the page. Only text and the facts above come out.
 */

const SKIP = new Set(['script', 'style', 'head', 'title', 'meta', 'link', 'template', 'xml', 'o:p', 'noscript', 'object', 'iframe', 'svg', 'img', 'canvas']);
const BLOCK = new Set(['p', 'div', 'li', 'blockquote', 'pre', 'section', 'article', 'header', 'footer', 'aside', 'nav', 'main', 'figure', 'figcaption', 'address', 'dd', 'dt', 'dl', 'ul', 'ol', 'caption', 'center', 'form', 'fieldset', 'details', 'summary']);
const TABLE_PARTS = new Set(['table', 'thead', 'tbody', 'tfoot', 'tr', 'td', 'th', 'colgroup', 'col']);
const INDENT_PER_LEVEL = 720;

interface Marks {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  sup?: boolean;
  sub?: boolean;
}

interface Ctx {
  marks: Marks;
  pre: boolean;
  /** Nesting of ul/ol, with the running number for ordered lists. */
  lists: { ordered: boolean; n: number }[];
  /** True inside a table cell: nested tables flatten to paragraphs. */
  inCell: boolean;
  /** Left indent inherited from blockquotes, in twips. */
  quote: number;
}

function styleOf(el: Element): Map<string, string> {
  const map = new Map<string, string>();
  const raw = el.getAttribute('style');
  if (!raw) return map;
  for (const decl of raw.split(';')) {
    const i = decl.indexOf(':');
    if (i === -1) continue;
    map.set(decl.slice(0, i).trim().toLowerCase(), decl.slice(i + 1).trim().toLowerCase());
  }
  return map;
}

function toTwips(value: string): number | null {
  const m = /^(-?\d*\.?\d+)\s*(in|pt|px|cm|mm)?$/.exec(value);
  if (!m) return null;
  const n = Number(m[1]);
  switch (m[2]) {
    case 'in': return Math.round(n * 1440);
    case 'pt': return Math.round(n * 20);
    case 'cm': return Math.round((n * 1440) / 2.54);
    case 'mm': return Math.round((n * 1440) / 25.4);
    case 'px': return Math.round(n * 15);
    default: return n === 0 ? 0 : null;
  }
}

function applyMarks(el: Element, tag: string, base: Marks): Marks {
  const marks: Marks = { ...base };
  switch (tag) {
    case 'b': case 'strong': marks.bold = true; break;
    case 'i': case 'em': case 'cite': case 'dfn': marks.italic = true; break;
    case 'u': case 'ins': marks.underline = true; break;
    case 's': case 'strike': case 'del': marks.strike = true; break;
    case 'sup': marks.sup = true; marks.sub = false; break;
    case 'sub': marks.sub = true; marks.sup = false; break;
  }
  const style = styleOf(el);
  const weight = style.get('font-weight');
  if (weight !== undefined) {
    if (weight === 'bold' || weight === 'bolder' || Number(weight) >= 600) marks.bold = true;
    else if (weight === 'normal' || weight === 'lighter' || Number(weight) <= 500) marks.bold = false;
  }
  const fontStyle = style.get('font-style');
  if (fontStyle === 'italic' || fontStyle === 'oblique') marks.italic = true;
  else if (fontStyle === 'normal') marks.italic = false;
  const deco = `${style.get('text-decoration') ?? ''} ${style.get('text-decoration-line') ?? ''}`;
  if (/underline/.test(deco)) marks.underline = true;
  if (/line-through/.test(deco)) marks.strike = true;
  if (/\bnone\b/.test(deco) && !/underline|line-through/.test(deco)) {
    marks.underline = false;
    marks.strike = false;
  }
  const va = style.get('vertical-align');
  if (va === 'super') { marks.sup = true; marks.sub = false; }
  else if (va === 'sub') { marks.sub = true; marks.sup = false; }
  return marks;
}

function alignOf(style: Map<string, string>): Align | undefined {
  switch (style.get('text-align')) {
    case 'center': return 'center';
    case 'right': return 'right';
    case 'justify': return 'both';
    default: return undefined;
  }
}

/** Word writes a list bullet as one Symbol/Wingdings character. Show it as a bullet. */
function normaliseMarker(text: string): string {
  const t = text.replace(/\s+/g, ' ');
  const core = t.trim();
  if (core.length <= 1 && !/[\p{L}\p{N}]/u.test(core) && core !== '') return '• ';
  if (core === 'o' || core === '·' || core === '§') return '• ';
  return t;
}

export function htmlToBlocks(html: string): Block[] {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const blocks: Block[] = [];

  /** The paragraph being built, for one container (the page, or a table cell). */
  const makeBuilder = (sink: Para[] | Block[]) => {
    let runs: Run[] = [];
    let props: Partial<Pick<Para, 'heading' | 'align' | 'indent' | 'list'>> = {};
    let force = false;

    const lastChar = (): string => {
      const last = runs[runs.length - 1];
      return last && !last.br && !last.tab ? last.text.slice(-1) : '';
    };

    return {
      begin(p: typeof props, forceEmit: boolean) {
        props = { ...props, ...p };
        if (forceEmit) force = true;
      },
      text(raw: string, marks: Marks, pre: boolean) {
        let text = raw;
        if (!pre) {
          text = text.replace(/[\s\u00a0]+/g, ' ');
          if (text.startsWith(' ') && (runs.length === 0 || lastChar() === ' ' || runs[runs.length - 1]?.br)) {
            text = text.slice(1);
          }
        }
        if (text === '') return;
        runs.push({ text, ...marks });
      },
      br() {
        runs.push({ text: '', br: true });
      },
      flush(extra: typeof props = {}) {
        while (runs.length > 0 && runs[runs.length - 1]?.br) runs.pop();
        const last = runs[runs.length - 1];
        if (last && !last.br && !last.tab) {
          const trimmed = last.text.replace(/ +$/, '');
          if (trimmed === '') runs.pop();
          else runs[runs.length - 1] = { ...last, text: trimmed };
        }
        if (runs.length > 0 || force) {
          const merged = { ...props, ...extra };
          sink.push({ kind: 'p', runs, ...merged } as Para & Block);
        }
        runs = [];
        props = {};
        force = false;
      },
      get open() {
        return runs.length > 0 || force;
      },
    };
  };

  const walkInto = (
    root: Node,
    sink: Para[] | Block[],
    ctx0: Ctx,
  ): void => {
    const b = makeBuilder(sink);

    const walk = (node: Node, ctx: Ctx): void => {
      if (node.nodeType === 3) {
        b.text(node.nodeValue ?? '', ctx.marks, ctx.pre);
        return;
      }
      if (node.nodeType !== 1) return;
      const el = node as Element;
      const tag = el.tagName.toLowerCase();
      if (SKIP.has(tag)) return;

      if (tag === 'br') {
        b.br();
        return;
      }
      if (tag === 'hr') {
        b.flush();
        b.begin({ align: 'center' }, true);
        b.text('———', {}, false);
        b.flush();
        return;
      }

      const style = styleOf(el);
      const heading = /^h[1-6]$/.test(tag) ? Number(tag[1]) : undefined;
      const marks = applyMarks(el, tag, ctx.marks);

      // Word marks the bullet or number of a list paragraph with mso-list:Ignore.
      if (/mso-list:\s*ignore/.test(el.getAttribute('style')?.toLowerCase() ?? '')) {
        b.text(normaliseMarker(el.textContent ?? ''), marks, false);
        return;
      }

      if (tag === 'table' && !ctx.inCell) {
        b.flush();
        const table = buildTable(el);
        if (table) (sink as Block[]).push(table);
        return;
      }

      const isBlock =
        BLOCK.has(tag) || heading !== undefined || (ctx.inCell && TABLE_PARTS.has(tag));

      if (!isBlock) {
        for (const child of Array.from(el.childNodes)) walk(child, { ...ctx, marks });
        return;
      }

      let next: Ctx = { ...ctx, marks };
      if (tag === 'pre') next = { ...next, pre: true };
      if (tag === 'blockquote') next = { ...next, quote: ctx.quote + INDENT_PER_LEVEL };
      if (tag === 'ul' || tag === 'ol') {
        b.flush();
        next = { ...next, lists: [...ctx.lists, { ordered: tag === 'ol', n: Number(el.getAttribute('start') ?? 1) || 1 }] };
        const level = next.lists[next.lists.length - 1]!;
        for (const child of Array.from(el.childNodes)) {
          if (child.nodeType === 1 && (child as Element).tagName.toLowerCase() === 'li') {
            walk(child, next);
            level.n += 1;
          } else walk(child, next);
        }
        b.flush();
        return;
      }

      b.flush();
      const msoList = /mso-list/.test(el.getAttribute('style')?.toLowerCase() ?? '');
      const margin = toTwips(style.get('margin-left') ?? '');
      const level = ctx.lists.length;
      const indentTwips =
        (tag === 'li' ? level * INDENT_PER_LEVEL : 0) + ctx.quote + (margin && margin > 0 ? margin : 0);
      b.begin(
        {
          ...(heading !== undefined ? { heading } : {}),
          ...(alignOf(style) ? { align: alignOf(style) as Align } : {}),
          ...(indentTwips > 0 ? { indent: indentTwips } : {}),
          ...(msoList || tag === 'li' ? { list: true } : {}),
        },
        tag === 'p' || tag === 'li' || heading !== undefined,
      );
      if (tag === 'li' && ctx.lists.length > 0) {
        const current = ctx.lists[ctx.lists.length - 1]!;
        b.text(current.ordered ? `${current.n}. ` : '• ', {}, false);
      }

      if (next.pre) {
        const text = el.textContent ?? '';
        const lines = text.replace(/\r\n?/g, '\n').split('\n');
        lines.forEach((line, i) => {
          if (i > 0) { b.flush(); b.begin({}, true); }
          b.text(line, marks, true);
        });
      } else {
        for (const child of Array.from(el.childNodes)) walk(child, next);
      }
      b.flush();
    };

    for (const child of Array.from(root.childNodes)) walk(child, ctx0);
    b.flush();
  };

  const buildTable = (tableEl: Element): Table | null => {
    const rows: Para[][][] = [];
    const rowEls: Element[] = [];
    const collect = (el: Element): void => {
      for (const child of Array.from(el.children)) {
        const t = child.tagName.toLowerCase();
        if (t === 'tr') rowEls.push(child);
        else if (t === 'thead' || t === 'tbody' || t === 'tfoot') collect(child);
      }
    };
    collect(tableEl);
    for (const tr of rowEls) {
      const cells: Para[][] = [];
      for (const cell of Array.from(tr.children)) {
        const t = cell.tagName.toLowerCase();
        if (t !== 'td' && t !== 'th') continue;
        const paras: Para[] = [];
        walkInto(cell, paras, {
          marks: t === 'th' ? { bold: true } : {},
          pre: false,
          lists: [],
          inCell: true,
          quote: 0,
        });
        cells.push(paras.length > 0 ? paras : [{ kind: 'p', runs: [] }]);
      }
      if (cells.length > 0) rows.push(cells);
    }
    return rows.length > 0 ? { kind: 'table', rows } : null;
  };

  walkInto(doc.body, blocks, { marks: {}, pre: false, lists: [], inCell: false, quote: 0 });
  return blocks;
}
