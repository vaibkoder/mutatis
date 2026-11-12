/**
 * Read-only side-by-side change view.
 *
 * PDF-vs-PDF has no openable tracked-changes document — PDFs have no redline
 * target — so a change must be *shown* in the task pane rather than written back.
 * This component renders one {@link DiffOp} as two read-only columns (Original |
 * Revised) with inline ins/del/eq runs, reusing the {@link WordDiff} token model
 * the Word path already produces.
 *
 * ADR-0008: NOTHING is hidden. Every `del` token appears on the Original side and
 * every `ins` token on the Revised side; a low-confidence op is rendered with a
 * badge, never suppressed. Pure presentational — no Office JS, no OOXML.
 */
import type { DiffOp, WordDiff } from '../model.js';

/**
 * Split a replace op's word diffs into the two sides. Original keeps equal + del
 * runs (what the base said); Revised keeps equal + ins runs (what it now says).
 * Pure — exported for tests.
 */
export function splitSides(wordDiffs: WordDiff[]): { original: WordDiff[]; revised: WordDiff[] } {
  const original: WordDiff[] = [];
  const revised: WordDiff[] = [];
  for (const wd of wordDiffs) {
    if (wd.op === 'eq') {
      original.push(wd);
      revised.push(wd);
    } else if (wd.op === 'del') {
      original.push(wd);
    } else {
      revised.push(wd);
    }
  }
  return { original, revised };
}

interface SideBySideDiffProps {
  op: DiffOp;
  /** Resolved plain text of the Original paragraph (for pure delete / move / equal). */
  originalText?: string | undefined;
  /** Resolved plain text of the Revised paragraph (for pure insert / move / equal). */
  revisedText?: string | undefined;
}

export function SideBySideDiff({ op, originalText, revisedText }: SideBySideDiffProps) {
  const lowConfidence = op.lowConfidence === true || op.fuzzyAnchored === true;
  return (
    <div
      class={`dp-compare__sbs${lowConfidence ? ' dp-compare__sbs--low-confidence' : ''}`}
      role="group"
      aria-label="Change detail"
    >
      {lowConfidence && (
        <p class="dp-compare__sbs-confidence" role="note">
          Low-confidence match — please verify against the source PDFs.
        </p>
      )}
      <div class="dp-compare__sbs-columns">
        <SideColumn label="Original" side="original" op={op} fallbackText={originalText} />
        <SideColumn label="Revised" side="revised" op={op} fallbackText={revisedText} />
      </div>
    </div>
  );
}

interface SideColumnProps {
  label: string;
  side: 'original' | 'revised';
  op: DiffOp;
  fallbackText?: string | undefined;
}

function SideColumn({ label, side, op, fallbackText }: SideColumnProps) {
  return (
    <div class={`dp-compare__sbs-col dp-compare__sbs-col--${side}`}>
      <span class="dp-compare__sbs-col-label">{label}</span>
      <p class="dp-compare__sbs-col-text" lang="en-GB">
        {renderSide(side, op, fallbackText)}
      </p>
    </div>
  );
}

/** Render one side's content: inline tokens for a replace, plain text otherwise. */
function renderSide(side: 'original' | 'revised', op: DiffOp, fallbackText?: string | undefined) {
  // Replace: render the word-diff tokens for this side inline.
  if (op.wordDiffs && op.wordDiffs.length > 0) {
    const tokens = side === 'original' ? splitSides(op.wordDiffs).original : splitSides(op.wordDiffs).revised;
    if (tokens.length > 0) return tokens.map((wd, i) => <Token key={i} wd={wd} />);
  }

  // Pure insert/delete/move/equal: the whole paragraph belongs (or not) to a side.
  const text = fallbackText ?? '';
  if (op.type === 'insert') {
    return side === 'revised' ? <ins class="dp-compare__tok dp-compare__tok--ins">{text}</ins> : <Absent />;
  }
  if (op.type === 'delete') {
    return side === 'original' ? <del class="dp-compare__tok dp-compare__tok--del">{text}</del> : <Absent />;
  }
  // move / equal: same text both sides (unchanged content).
  if (text !== '') return <span class="dp-compare__tok dp-compare__tok--eq">{text}</span>;
  return <Absent />;
}

function Token({ wd }: { wd: WordDiff }) {
  if (wd.op === 'ins') return <ins class="dp-compare__tok dp-compare__tok--ins">{wd.text}</ins>;
  if (wd.op === 'del') return <del class="dp-compare__tok dp-compare__tok--del">{wd.text}</del>;
  return <span class="dp-compare__tok dp-compare__tok--eq">{wd.text}</span>;
}

/** Placeholder for a side where the paragraph is absent (pure insert/delete). */
function Absent() {
  return <span class="dp-compare__sbs-absent" aria-label="not present on this side">—</span>;
}
