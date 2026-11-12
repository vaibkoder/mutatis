import type { RedlineContext } from './context.js';

/**
 * The deviation matrix — COMPUTED, never generated.
 *
 * By design the matrix is computed deterministically and only
 * NARRATED by the model. Everything in this file is a pure function of the
 * change sets already produced by the engine, so the same set of comparisons
 * yields byte-identical rows every run. The model never sees a blank space it
 * could fill with a plausible number.
 *
 * The output is rendered as its own tab on the right, beside the redlines. The
 * assistant's narration sits in the chat column, to the side of it — never in
 * front of it, and never as a substitute for it.
 */

export interface DeviationCell {
  readonly count: number;
  /** Change ids behind this cell, so a narration can be checked against the redline. */
  readonly ids: readonly string[];
}

export interface DeviationRow {
  readonly classification: string;
  readonly cells: readonly DeviationCell[];
  readonly total: number;
  /** Documents in which this classification appears at all. */
  readonly documents: number;
}

export interface DeviationColumn {
  readonly title: string;
  readonly baseline: string;
  readonly revised: string;
  readonly changes: number;
  /** Changes the context payload omitted for this comparison. Disclosed, not hidden. */
  readonly omitted: number;
}

export interface DeviationMatrix {
  readonly generated: string;
  readonly columns: readonly DeviationColumn[];
  readonly rows: readonly DeviationRow[];
  /** Classifications present in EVERY column — the house position, in effect. */
  readonly universal: readonly string[];
  /** Classifications present in exactly one column — the outliers. */
  readonly unique: readonly string[];
  readonly notes: readonly string[];
}

export interface MatrixInput {
  readonly id: string;
  readonly title: string;
  readonly context: RedlineContext;
}

export function buildDeviationMatrix(
  inputs: readonly MatrixInput[],
  generated: string,
): DeviationMatrix {
  const columns: DeviationColumn[] = inputs.map((input) => ({
    title: input.title,
    baseline: input.context.baseline,
    revised: input.context.revised,
    changes: input.context.totals.total,
    omitted: input.context.omitted,
  }));

  // Sorted, so the row order is a property of the data and not of Map insertion
  // order — which would make two runs over the same comparisons disagree.
  const classifications = [
    ...new Set(
      inputs.flatMap((input) => input.context.changes.map((change) => change.classification)),
    ),
  ].sort();

  const rows: DeviationRow[] = classifications.map((classification) => {
    const cells: DeviationCell[] = inputs.map((input) => {
      const ids = input.context.changes
        .filter((change) => change.classification === classification)
        .map((change) => change.id);
      return { count: ids.length, ids };
    });
    return {
      classification,
      cells,
      total: cells.reduce((sum, cell) => sum + cell.count, 0),
      documents: cells.filter((cell) => cell.count > 0).length,
    };
  });

  // Most-used rows first, then alphabetically — a stable order a reader can scan.
  const ordered = [...rows].sort(
    (a, b) => b.total - a.total || a.classification.localeCompare(b.classification),
  );

  return {
    generated,
    columns,
    rows: ordered,
    universal: ordered
      .filter((row) => columns.length > 0 && row.documents === columns.length)
      .map((row) => row.classification),
    unique: ordered.filter((row) => row.documents === 1).map((row) => row.classification),
    notes: notesFor(inputs, columns),
  };
}

function notesFor(
  inputs: readonly MatrixInput[],
  columns: readonly DeviationColumn[],
): string[] {
  const notes: string[] = [];

  notes.push(
    'Every number here is counted from the comparison engine\'s own change set. ' +
      'It is not an estimate and does not come from a language model.',
  );

  const truncated = columns.filter((column) => column.omitted > 0);
  if (truncated.length > 0) {
    notes.push(
      `${truncated.length} comparison(s) had changes omitted from the assistant's context ` +
        `budget (${truncated.map((c) => `${c.title}: ${c.omitted}`).join(', ')}). Those changes ` +
        'are NOT counted in this matrix. The redline tabs remain complete.',
    );
  }

  const jurisdictions = [...new Set(inputs.map((input) => input.context.jurisdiction))];
  if (jurisdictions.length > 1) {
    notes.push(
      `These comparisons were run under different jurisdiction profiles ` +
        `(${jurisdictions.join(', ')}), which changes what counts as a change. ` +
        'Compare the columns with that in mind.',
    );
  }

  notes.push(
    'Classifications are the engine\'s categories, not a commercial judgement. ' +
      'A "substantive" change is not necessarily important and a "formatting" change is not ' +
      'necessarily trivial.',
  );

  return notes;
}

/** The matrix as the text the model narrates. Deterministic, like the matrix. */
export function matrixAsText(matrix: DeviationMatrix): string {
  const lines: string[] = [];
  lines.push(`Deviation matrix across ${matrix.columns.length} comparison(s).`);
  lines.push('');
  lines.push(`Columns: ${matrix.columns.map((c) => c.title).join(' | ')}`);
  lines.push('');
  lines.push(`classification\t${matrix.columns.map((c) => c.title).join('\t')}\ttotal`);
  for (const row of matrix.rows) {
    lines.push(
      `${row.classification}\t${row.cells.map((cell) => cell.count).join('\t')}\t${row.total}`,
    );
  }
  lines.push('');
  lines.push(
    `In every comparison: ${matrix.universal.length > 0 ? matrix.universal.join(', ') : 'none'}`,
  );
  lines.push(
    `In only one comparison: ${matrix.unique.length > 0 ? matrix.unique.join(', ') : 'none'}`,
  );
  lines.push('');
  for (const note of matrix.notes) lines.push(`NOTE: ${note}`);
  return lines.join('\n');
}
