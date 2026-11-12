/**
 * Markdown-table detection + parse for the Text-vs-text Compare path.
 * Spec: spec/features/10-compare-additional-sources.md §4.1 · ADR-0027.
 *
 * Splits a pasted plain-text blob into ordered BLOCKS — `paragraph` lines and
 * GitHub-style pipe `table` grids. Detection is fully DETERMINISTIC (no library,
 * no ML, ADR-0008/0027): a `table` is a header row containing an unescaped pipe,
 * IMMEDIATELY followed by a delimiter row (`---|---`), then zero or more body
 * rows containing a pipe. Anything that does not match that exact shape stays a
 * `paragraph` line — including a lone `---|---` with no header above it (MT7).
 *
 * Faithful, never-suppress (ADR-0008): ragged rows are PADDED to the widest row
 * so no cell is dropped; the delimiter row is consumed as table structure (it is
 * never emitted as content, and never silently turns surrounding text into a
 * table). A literal pipe inside a cell is written `\|` (mirrors the AI-export
 * emitter convention, src/features/ai-export/formats/markdown.ts).
 *
 * Pure: no Office JS, no Date.now(), no Math.random().
 */

/** One block of a parsed paste: a single text line, or a pipe table. */
export type Block =
  | { kind: 'paragraph'; text: string }
  /** rows[0] is the header; rows[1..] are body rows. All rows padded to the same width. */
  | { kind: 'table'; rows: string[][] };

/**
 * A delimiter cell: optional leading/trailing alignment colon around ≥1 dash.
 * Alignment colons are tolerated but not acted on in v1 (no column-align styling).
 */
const DELIM_CELL = /^:?-+:?$/;

/**
 * Split a Markdown table row into trimmed cell strings.
 *
 * - `\|` is a literal pipe inside a cell (not a delimiter); `\\` is a literal
 *   backslash. Every other `|` separates cells.
 * - A single optional leading and trailing outer pipe is dropped (so both
 *   `| a | b |` and `a | b` yield `['a','b']`); genuinely empty interior cells
 *   (`a || b`) are preserved.
 */
export function splitRow(line: string): string[] {
  const cells: string[] = [];
  let cur = '';
  for (let i = 0; i < line.length; i++) {
    const ch = line[i]!;
    if (ch === '\\' && i + 1 < line.length) {
      const next = line[i + 1]!;
      // Preserve the escape semantics for pipe/backslash; pass others through.
      if (next === '|' || next === '\\') {
        cur += next;
        i++;
        continue;
      }
      cur += ch;
      continue;
    }
    if (ch === '|') {
      cells.push(cur);
      cur = '';
      continue;
    }
    cur += ch;
  }
  cells.push(cur);

  // Drop one empty edge cell produced by an optional outer pipe.
  if (cells.length > 1 && cells[0]!.trim() === '') cells.shift();
  if (cells.length > 1 && cells[cells.length - 1]!.trim() === '') cells.pop();

  return cells.map((c) => c.trim());
}

/** True if `line` contains an unescaped pipe (i.e. is shaped like a table row). */
function hasUnescapedPipe(line: string): boolean {
  for (let i = 0; i < line.length; i++) {
    const ch = line[i]!;
    if (ch === '\\') {
      i++;
      continue;
    }
    if (ch === '|') return true;
  }
  return false;
}

/** True if `line` is a Markdown table delimiter row (`---`, `:--`, `--:|--`, …). */
function isDelimiterRow(line: string): boolean {
  if (!hasUnescapedPipe(line) && !DELIM_CELL.test(line.trim())) return false;
  const cells = splitRow(line);
  if (cells.length === 0) return false;
  return cells.every((c) => DELIM_CELL.test(c));
}

/** Pad every row to the width of the widest row so no cell is ever dropped. */
function padRows(rows: string[][]): string[][] {
  let width = 0;
  for (const r of rows) if (r.length > width) width = r.length;
  return rows.map((r) => (r.length === width ? r : [...r, ...Array(width - r.length).fill('')]));
}

/**
 * Parse a plain-text blob into ordered blocks. Newlines are normalised
 * (CRLF/CR → LF) first. Blank lines become empty `paragraph` blocks (the
 * faithful, categorise-never-suppress representation, matching the existing
 * `synthesise-docx` `blankLines:'keep'` default).
 */
export function parseBlocks(text: string): Block[] {
  const lines = text.replace(/\r\n?/g, '\n').split('\n');
  // A terminal newline yields a trailing '' — a line terminator, not a paragraph.
  if (lines.length > 1 && lines[lines.length - 1] === '') lines.pop();

  const blocks: Block[] = [];
  let i = 0;
  while (i < lines.length) {
    const line = lines[i]!;
    const next = i + 1 < lines.length ? lines[i + 1]! : undefined;

    // Table = header row (has a pipe) + delimiter row immediately below.
    if (hasUnescapedPipe(line) && next !== undefined && isDelimiterRow(next)) {
      const rows: string[][] = [splitRow(line)];
      i += 2; // consume header + delimiter (delimiter is structure, never content)
      while (i < lines.length && lines[i]!.trim() !== '' && hasUnescapedPipe(lines[i]!)) {
        rows.push(splitRow(lines[i]!));
        i++;
      }
      blocks.push({ kind: 'table', rows: padRows(rows) });
      continue;
    }

    blocks.push({ kind: 'paragraph', text: line });
    i++;
  }

  return blocks;
}
