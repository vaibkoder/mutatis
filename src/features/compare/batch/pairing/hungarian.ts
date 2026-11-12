/**
 * Optimal assignment on a square cost matrix — O(n³) Kuhn–Munkres with
 * potentials (the shortest-augmenting-path formulation).
 *
 * Why an exact solver rather than "each document takes its best match": greedy
 * is not symmetric and not order-independent. If the SPA's best candidate is
 * also the disclosure letter's best candidate, greedy gives the loser whatever
 * is left over, and which one loses depends on which was visited first. That
 * fails acceptance M4 (shuffling both bundles must give an identical
 * assignment) and, worse, produces a *plausible* wrong pairing — which is the
 * failure this whole feature is built to prevent. At M, N ≤ 50 the cubic cost
 * is irrelevant: a 100×100 matrix is a million operations.
 *
 * The implementation is the standard one, kept close to its textbook form on
 * purpose — this is the module where cleverness is least welcome. Internally
 * 1-indexed (the algorithm's sentinel row/column 0 is what makes the augmenting
 * path terminate); the exported function is 0-indexed.
 *
 * Pure — no Office JS, no DOM, no network.
 */

/**
 * Minimum-cost perfect assignment of rows to columns.
 *
 * @param cost `cost[i][j]` — finite. Must be square or wider than tall.
 * @returns `rowToCol[i]` — the column assigned to row `i`.
 */
export function hungarian(cost: readonly (readonly number[])[]): number[] {
  const n = cost.length;
  if (n === 0) return [];
  const m = (cost[0] as readonly number[]).length;
  if (m < n) {
    throw new Error(`hungarian: needs at least as many columns as rows (got ${n}×${m})`);
  }

  // `u`/`v` are the dual potentials; `p[j]` is the row currently assigned to
  // column `j`; `way[j]` records the augmenting path back to column 0.
  const u = new Array<number>(n + 1).fill(0);
  const v = new Array<number>(m + 1).fill(0);
  const p = new Array<number>(m + 1).fill(0);
  const way = new Array<number>(m + 1).fill(0);

  for (let i = 1; i <= n; i += 1) {
    p[0] = i;
    let j0 = 0;
    const minv = new Array<number>(m + 1).fill(Infinity);
    const used = new Array<boolean>(m + 1).fill(false);

    do {
      used[j0] = true;
      const i0 = p[j0] as number;
      let delta = Infinity;
      let j1 = 0;

      for (let j = 1; j <= m; j += 1) {
        if (used[j]) continue;
        const cur = (cost[i0 - 1] as readonly number[])[j - 1] as number
          - (u[i0] as number)
          - (v[j] as number);
        if (cur < (minv[j] as number)) {
          minv[j] = cur;
          way[j] = j0;
        }
        if ((minv[j] as number) < delta) {
          delta = minv[j] as number;
          j1 = j;
        }
      }

      for (let j = 0; j <= m; j += 1) {
        if (used[j]) {
          u[p[j] as number] = (u[p[j] as number] as number) + delta;
          v[j] = (v[j] as number) - delta;
        } else {
          minv[j] = (minv[j] as number) - delta;
        }
      }

      j0 = j1;
    } while (p[j0] !== 0);

    // Walk the augmenting path back, flipping each edge.
    do {
      const j1 = way[j0] as number;
      p[j0] = p[j1] as number;
      j0 = j1;
    } while (j0 !== 0);
  }

  const rowToCol = new Array<number>(n).fill(-1);
  for (let j = 1; j <= m; j += 1) {
    const row = p[j] as number;
    if (row > 0) rowToCol[row - 1] = j - 1;
  }
  return rowToCol;
}
