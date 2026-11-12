/**
 * Minimal sfnt (TrueType/OpenType) metric reader — host-free, DOM-free, pure.
 *
 * Reads only what the layout engine needs for Word line-height: `unitsPerEm`
 * (head), and the OS/2 win/typo ascender/descender + USE_TYPO_METRICS flag.
 * Shaping/advances come from HarfBuzz; PDF subsetting from fontkit (inside
 * pdf-lib) — so we deliberately do NOT pull a full font library into layout.
 *
 * Word "single" line height (Spike A, calibration policy):
 *   height = ((useTypoMetrics ? (typoAsc - typoDesc + typoGap)
 *                              : (winAsc + winDesc)) / unitsPerEm) × sizePt
 */

export interface LineMetrics {
  unitsPerEm: number;
  useTypoMetrics: boolean;
  winAscent: number;
  winDescent: number;
  typoAscender: number;
  typoDescender: number;
  typoLineGap: number;
}

function tag(view: DataView, off: number): string {
  return String.fromCharCode(
    view.getUint8(off),
    view.getUint8(off + 1),
    view.getUint8(off + 2),
    view.getUint8(off + 3),
  );
}

/** Locate sfnt table offsets by tag. Returns a map tag → byte offset. */
function tableDirectory(view: DataView): Map<string, number> {
  const out = new Map<string, number>();
  const numTables = view.getUint16(4);
  let p = 12;
  for (let i = 0; i < numTables; i++, p += 16) {
    out.set(tag(view, p), view.getUint32(p + 8));
  }
  return out;
}

/**
 * Read line metrics from raw sfnt bytes (TTF/OTF). Throws if `head` is missing;
 * falls back to win≈typo when `OS/2` is absent (older fonts).
 */
export function readLineMetrics(bytes: Uint8Array): LineMetrics {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const dir = tableDirectory(view);

  const head = dir.get('head');
  if (head === undefined) throw new Error('sfnt: missing head table');
  const unitsPerEm = view.getUint16(head + 18);

  const os2 = dir.get('OS/2');
  if (os2 === undefined) {
    // No OS/2 — approximate with em box (rare for our fonts).
    return {
      unitsPerEm,
      useTypoMetrics: false,
      winAscent: Math.round(unitsPerEm * 0.8),
      winDescent: Math.round(unitsPerEm * 0.2),
      typoAscender: Math.round(unitsPerEm * 0.8),
      typoDescender: -Math.round(unitsPerEm * 0.2),
      typoLineGap: 0,
    };
  }

  const typoAscender = view.getInt16(os2 + 68);
  const typoDescender = view.getInt16(os2 + 70);
  const typoLineGap = view.getInt16(os2 + 72);
  const winAscent = view.getUint16(os2 + 74);
  const winDescent = view.getUint16(os2 + 76);
  const fsSelection = view.getUint16(os2 + 62);
  const useTypoMetrics = (fsSelection & 0x80) !== 0; // bit 7

  return {
    unitsPerEm,
    useTypoMetrics,
    winAscent,
    winDescent,
    typoAscender,
    typoDescender,
    typoLineGap,
  };
}

/** Word "single" line height in points for a given font size. */
export function singleLineHeightPt(m: LineMetrics, sizePt: number): number {
  const fu = m.useTypoMetrics
    ? m.typoAscender - m.typoDescender + m.typoLineGap
    : m.winAscent + m.winDescent;
  return (fu / m.unitsPerEm) * sizePt;
}

/** Ascent (baseline offset from line-box top) in points. */
export function ascentPt(m: LineMetrics, sizePt: number): number {
  const asc = m.useTypoMetrics ? m.typoAscender : m.winAscent;
  return (asc / m.unitsPerEm) * sizePt;
}
