import type { EffectivePPr, EffectiveRPr } from '@shared/render-core/types.js';
import type { CSSProperties } from 'react';

/**
 * Map the engine's effective OOXML properties onto CSS.
 *
 * These are applied as React style objects, which React writes through the CSSOM
 * (`node.style.setProperty`) rather than as parsed HTML attributes. That matters
 * under our CSP: `style-src 'self'` (which style-src-attr falls back to) polices
 * style ATTRIBUTES found when parsing HTML — it does not police CSSOM writes from
 * script. A client-rendered app never trips it.
 */

/**
 * Word fonts mapped to metric-compatible open equivalents, then to a generic.
 *
 * The engine vendors the actual OFL clones (Carlito, Caladea, Liberation) for
 * its paginated and PDF output, where exact metrics decide line breaking. Here
 * the browser reflows the text, so naming them is a best-effort nicety and the
 * generic tail is what actually has to be right.
 */
const FONT_STACKS: Record<string, string> = {
  calibri: 'Carlito, Calibri, ui-sans-serif, system-ui, sans-serif',
  cambria: 'Caladea, Cambria, ui-serif, Georgia, serif',
  arial: '"Liberation Sans", Arial, ui-sans-serif, system-ui, sans-serif',
  helvetica: '"Liberation Sans", Helvetica, ui-sans-serif, system-ui, sans-serif',
  'times new roman': '"Liberation Serif", "Times New Roman", ui-serif, Times, serif',
  georgia: 'Georgia, ui-serif, serif',
  garamond: 'Garamond, ui-serif, Georgia, serif',
  'courier new': '"Liberation Mono", "Courier New", ui-monospace, monospace',
  verdana: 'Verdana, ui-sans-serif, system-ui, sans-serif',
};

export function fontStack(family: string): string {
  const key = family.trim().toLowerCase();
  return FONT_STACKS[key] ?? `"${family}", ui-serif, Georgia, serif`;
}

const JUSTIFICATION: Record<string, CSSProperties['textAlign']> = {
  left: 'left',
  right: 'right',
  center: 'center',
  both: 'justify',
};

/**
 * Line spacing.
 *
 * OOXML encodes this two different ways in one pair of fields, and getting it
 * wrong is the difference between a readable document and a cramped one:
 *   - `auto`: lineValue is a multiple in 240ths (240 = single, 360 = 1.5).
 *   - `exact` / `atLeast`: lineValue is twips (twentieths of a point).
 * A lineValue of 0 means "single", not "zero height".
 */
function lineHeight(ppr: EffectivePPr): CSSProperties['lineHeight'] {
  if (ppr.lineValue === 0) return undefined; // inherit the comfortable default
  if (ppr.lineRule === 'auto') return ppr.lineValue / 240;
  return `${ppr.lineValue / 20}pt`;
}

export function paragraphStyle(ppr: EffectivePPr): CSSProperties {
  const style: CSSProperties = {
    textAlign: JUSTIFICATION[ppr.justification] ?? 'left',
    marginTop: `${ppr.spaceBeforePt}pt`,
    marginBottom: `${ppr.spaceAfterPt}pt`,
    marginRight: ppr.indentRightPt > 0 ? `${ppr.indentRightPt}pt` : undefined,
  };

  const lh = lineHeight(ppr);
  if (lh !== undefined) style.lineHeight = lh;

  // A hanging indent is the first line pulled LEFT out of an indented block, so
  // it needs both halves: extra left padding to hold the block in, and a
  // negative text-indent to pull line one back out. firstLinePt is the simple
  // opposite. The two are mutually exclusive in OOXML.
  if (ppr.hangingPt > 0) {
    style.marginLeft = `${ppr.indentLeftPt}pt`;
    style.paddingLeft = `${ppr.hangingPt}pt`;
    style.textIndent = `-${ppr.hangingPt}pt`;
  } else {
    style.marginLeft = `${ppr.indentLeftPt}pt`;
    if (ppr.firstLinePt > 0) style.textIndent = `${ppr.firstLinePt}pt`;
  }

  return style;
}

export function runStyle(rpr: EffectiveRPr): CSSProperties {
  const style: CSSProperties = {
    fontFamily: fontStack(rpr.fontFamily),
    fontSize: `${rpr.sizePt}pt`,
  };
  if (rpr.bold) style.fontWeight = 700;
  if (rpr.italic) style.fontStyle = 'italic';
  // Underline is left to the revision decoration when the run is an insertion —
  // see RedlineRun. Applying both would make an underlined insertion
  // indistinguishable from an underlined unchanged run.
  return style;
}

/** Vertical alignment is a separate element concern (sup/sub), not a style. */
export function verticalTag(rpr: EffectiveRPr): 'sup' | 'sub' | null {
  if (rpr.superscript) return 'sup';
  if (rpr.subscript) return 'sub';
  return null;
}
