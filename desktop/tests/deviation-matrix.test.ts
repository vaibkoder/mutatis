import { describe, it, expect } from 'vitest';
import { buildDeviationMatrix, matrixAsText } from '../src/renderer/src/ai/deviation-matrix.js';
import type { RedlineContext } from '../src/renderer/src/ai/context.js';

/**
 * The matrix is the thing that must never be free-formed by a
 * model: "the output must be consistent run to run". These tests are that
 * promise, written down.
 */

function context(
  classifications: string[],
  overrides: Partial<RedlineContext> = {},
): RedlineContext {
  return {
    generated: '2026-01-01T00:00:00.000Z',
    jurisdiction: 'UK',
    baseline: 'a.docx',
    revised: 'b.docx',
    totals: { total: classifications.length, byCategory: {}, moved: 0, spellingVariants: 0 },
    changes: classifications.map((classification, index) => ({
      id: `op-${index}`,
      type: 'replace' as const,
      classification: classification as never,
      clause: '',
    })),
    omitted: 0,
    blindSpots: [],
    ...overrides,
  };
}

const AT = '2026-01-01T00:00:00.000Z';

describe('buildDeviationMatrix', () => {
  it('counts classifications per comparison', () => {
    const matrix = buildDeviationMatrix(
      [
        { id: 't1', title: 'One', context: context(['substantive', 'substantive', 'formatting']) },
        { id: 't2', title: 'Two', context: context(['substantive']) },
      ],
      AT,
    );

    const substantive = matrix.rows.find((row) => row.classification === 'substantive');
    expect(substantive?.cells.map((cell) => cell.count)).toEqual([2, 1]);
    expect(substantive?.total).toBe(3);
    expect(substantive?.documents).toBe(2);

    const formatting = matrix.rows.find((row) => row.classification === 'formatting');
    expect(formatting?.cells.map((cell) => cell.count)).toEqual([1, 0]);
  });

  it('is byte-identical across runs over the same input', () => {
    const inputs = [
      { id: 't1', title: 'One', context: context(['b', 'a', 'b']) },
      { id: 't2', title: 'Two', context: context(['c', 'a']) },
    ];
    expect(JSON.stringify(buildDeviationMatrix(inputs, AT))).toBe(
      JSON.stringify(buildDeviationMatrix(inputs, AT)),
    );
  });

  it('orders rows by total then name, not by encounter order', () => {
    const matrix = buildDeviationMatrix(
      [
        { id: 't1', title: 'One', context: context(['zebra', 'apple', 'apple']) },
        { id: 't2', title: 'Two', context: context(['apple', 'mango']) },
      ],
      AT,
    );
    expect(matrix.rows.map((row) => row.classification)).toEqual(['apple', 'mango', 'zebra']);
  });

  it('separates what appears everywhere from what appears once', () => {
    const matrix = buildDeviationMatrix(
      [
        { id: 't1', title: 'One', context: context(['shared', 'only-here']) },
        { id: 't2', title: 'Two', context: context(['shared']) },
      ],
      AT,
    );
    expect(matrix.universal).toEqual(['shared']);
    expect(matrix.unique).toEqual(['only-here']);
  });

  it('carries the change ids behind every cell, so a claim can be checked', () => {
    const matrix = buildDeviationMatrix(
      [{ id: 't1', title: 'One', context: context(['substantive', 'substantive']) }],
      AT,
    );
    expect(matrix.rows[0]?.cells[0]?.ids).toEqual(['op-0', 'op-1']);
  });

  it('DISCLOSES omitted changes rather than presenting the counts as complete', () => {
    const matrix = buildDeviationMatrix(
      [
        { id: 't1', title: 'One', context: context(['substantive'], { omitted: 12 }) },
        { id: 't2', title: 'Two', context: context(['substantive']) },
      ],
      AT,
    );
    expect(matrix.notes.some((note) => note.includes('One: 12'))).toBe(true);
    expect(matrix.notes.some((note) => note.includes('NOT counted'))).toBe(true);
  });

  it('flags a mixed-jurisdiction matrix — the profiles disagree about what a change is', () => {
    const matrix = buildDeviationMatrix(
      [
        { id: 't1', title: 'One', context: context(['substantive']) },
        { id: 't2', title: 'Two', context: context(['substantive'], { jurisdiction: 'SG' }) },
      ],
      AT,
    );
    expect(matrix.notes.some((note) => /different jurisdiction profiles/.test(note))).toBe(true);
  });

  it('always says the numbers are counted, not generated', () => {
    const matrix = buildDeviationMatrix(
      [{ id: 't1', title: 'One', context: context([]) }],
      AT,
    );
    expect(matrix.notes[0]).toMatch(/does not come from a language model/);
  });

  it('handles a comparison with no changes without inventing a row', () => {
    const matrix = buildDeviationMatrix([{ id: 't1', title: 'One', context: context([]) }], AT);
    expect(matrix.rows).toEqual([]);
    expect(matrix.universal).toEqual([]);
  });
});

describe('matrixAsText', () => {
  it('renders every row and every note for the model to narrate', () => {
    const text = matrixAsText(
      buildDeviationMatrix(
        [
          { id: 't1', title: 'One', context: context(['substantive', 'formatting']) },
          { id: 't2', title: 'Two', context: context(['substantive']) },
        ],
        AT,
      ),
    );
    expect(text).toContain('substantive\t1\t1\t2');
    expect(text).toContain('formatting\t1\t0\t1');
    expect(text).toContain('In every comparison: substantive');
    expect(text).toContain('In only one comparison: formatting');
    expect(text).toContain('NOTE:');
  });
});
