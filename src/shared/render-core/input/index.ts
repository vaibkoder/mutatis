/**
 * Build a resolved `RenderModel` from `authoredBodyXml` + styles + footnotes +
 * numbering — host-free, pure. The single entry point for `input/` parsing.
 */
import { type PoNode, parsePreserveOrder, tagOf, childrenOf, attrOf } from '@shared/ooxml/po-xml.js';
import {
  renderListNumbers, type NumberingScheme, type ParaNumPr,
} from '@shared/numbering/index.js';
import type {
  RenderModel, RcFootnote, PageGeometry, RcBlock, RcRun, RcChangeClass,
} from '../types.js';
import { parseStyles } from './styles.js';
import { newCtx, parseBlocks, findSectPr } from './parse-body.js';
import { geometryFromSectPr } from './sectpr.js';

export interface RenderInput {
  /** Merged `<w:body>` with native revision markup (the Compare authoredBodyXml). */
  authoredBodyXml: string;
  stylesXml?: string;
  footnotesXml?: string;
  /** Pre-parsed numbering scheme (from `parseNumberingScheme`); enables real labels. */
  numbering?: NumberingScheme;
  /**
   * Change-class annotations keyed by native revision `@w:id` (the Compare
   * `classByRevId` side channel). Applied as styling-only `RcRun.changeClasses`
   * (ADR-0008: never filters/drops a revision — every native ins/del/move still
   * renders whether or not a class is attached).
   */
  classByRevId?: ReadonlyMap<string, readonly RcChangeClass[]>;
  /** Override page geometry (else read from sectPr, else A4 default). */
  geometry?: PageGeometry;
}

/** Walk every run in a block tree (paragraphs, table cells, nested) in order. */
function forEachRun(blocks: RcBlock[], visit: (run: RcRun) => void): void {
  for (const block of blocks) {
    if (block.kind === 'paragraph') {
      for (const run of block.paragraph.runs) visit(run);
    } else {
      for (const row of block.table.rows) {
        for (const cell of row.cells) forEachRun(cell.blocks, visit);
      }
    }
  }
}

/**
 * Attach `changeClasses` to every run sitting inside a native revision whose
 * `@w:id` is present in `classByRevId`. Styling annotation only — it never
 * removes a run or a revision (ADR-0008); a run with no mapped class simply
 * keeps `changeClasses` unset and still renders + decorates by its revision type.
 */
function applyClassAnnotations(
  blocks: RcBlock[],
  footnotes: Map<string, RcFootnote>,
  classByRevId: ReadonlyMap<string, readonly RcChangeClass[]>,
): void {
  const annotate = (run: RcRun): void => {
    const id = run.revision?.changeId;
    if (id === undefined) return;
    const classes = classByRevId.get(id);
    if (classes && classes.length) run.changeClasses = classes;
  };
  forEachRun(blocks, annotate);
  for (const fn of footnotes.values()) forEachRun(fn.blocks, annotate);
}

/** Recursively find a `<w:body>` element's children (handles raw body or w:document>body). */
function findBodyChildren(roots: PoNode[]): PoNode[] {
  for (const n of roots) {
    if (tagOf(n) === 'w:body') return childrenOf(n);
  }
  for (const n of roots) {
    const kids = childrenOf(n);
    if (kids.length) {
      const found = findBodyChildren(kids);
      if (found.length) return found;
    }
  }
  // No explicit body wrapper — treat the roots themselves as block-level content.
  return roots;
}

export function buildRenderModel(input: RenderInput): RenderModel {
  const styles = parseStyles(input.stylesXml);
  const ctx = newCtx(styles);

  const bodyChildren = findBodyChildren(parsePreserveOrder(input.authoredBodyXml));
  const blocks = parseBlocks(bodyChildren, ctx);
  const sectPr = findSectPr(bodyChildren);
  const geometry = input.geometry ?? geometryFromSectPr(sectPr);

  // Footnote bodies (parsed with a separate ctx; footnote-body numbering is deferred).
  const footnotes = new Map<string, RcFootnote>();
  if (input.footnotesXml) {
    const fnRoots = parsePreserveOrder(input.footnotesXml);
    const fnRoot = fnRoots.find((n) => tagOf(n) === 'w:footnotes');
    if (fnRoot) {
      for (const fn of childrenOf(fnRoot)) {
        if (tagOf(fn) !== 'w:footnote') continue;
        const type = attrOf(fn, '@_w:type'); // separator / continuationSeparator
        if (type === 'separator' || type === 'continuationSeparator') continue;
        const id = attrOf(fn, '@_w:id');
        if (id === undefined) continue;
        const fnCtx = newCtx(styles);
        footnotes.set(id, { id, blocks: parseBlocks(childrenOf(fn), fnCtx) });
      }
    }
  }

  // Numbering labels via the shared resolver (document order).
  if (input.numbering && ctx.numberingRefs.length) {
    const refs: ParaNumPr[] = ctx.numberingRefs.map((r) => ({
      paraId: r.paraId, numId: r.numId, ilvl: r.ilvl,
    }));
    const labels = renderListNumbers(input.numbering, refs);
    for (const r of ctx.numberingRefs) {
      const ln = labels.get(r.paraId);
      if (ln) {
        r.para.numberLabel = ln.listString;
        r.para.numberLevel = ln.listLevel;
      }
    }
  }

  // Class annotations (styling only; never filters — ADR-0008).
  if (input.classByRevId && input.classByRevId.size) {
    applyClassAnnotations(blocks, footnotes, input.classByRevId);
  }

  return { blocks, footnotes, geometry };
}

export type { RenderInput as RenderModelInput };
