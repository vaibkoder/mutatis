/**
 * Longest Common Subsequence on string keys.
 * Spec: spec/features/08-compare.md §5.1 (segmented LCS), §9.3 (reused for tables).
 *
 * Classic O(n·m) dynamic-programming LCS. Used on paragraph-hash sequences
 * within anchor-bounded segments (≤ ~1000 items per segment), and on table
 * row/cell hashes. Pure.
 */

/**
 * Compute the LCS of two arrays by a string key projection.
 * Returns the matched index pairs `[aIndex, bIndex]` in ascending order.
 *
 * @param a    first sequence.
 * @param b    second sequence.
 * @param key  projects an element to the string compared for equality.
 */
export function computeLCS<T>(a: T[], b: T[], key: (x: T) => string): Array<[number, number]> {
  const n = a.length;
  const m = b.length;
  if (n === 0 || m === 0) return [];

  const aKeys = a.map(key);
  const bKeys = b.map(key);

  // dp[i][j] = LCS length of a[i..] and b[j..]. Dimensions (n+1)×(m+1).
  // Use a flat typed array for the table to keep allocation tight.
  const width = m + 1;
  const dp = new Int32Array((n + 1) * width);

  // Int32Array reads are always in-bounds here; `?? 0` only quiets
  // noUncheckedIndexedAccess (TypedArray indexing is typed `number | undefined`).
  const at = (k: number): number => dp[k] ?? 0;

  for (let i = n - 1; i >= 0; i--) {
    const rowBase = i * width;
    const nextRowBase = (i + 1) * width;
    for (let j = m - 1; j >= 0; j--) {
      if (aKeys[i] === bKeys[j]) {
        dp[rowBase + j] = at(nextRowBase + j + 1) + 1;
      } else {
        const down = at(nextRowBase + j);
        const right = at(rowBase + j + 1);
        dp[rowBase + j] = down >= right ? down : right;
      }
    }
  }

  // Backtrack from (0,0) to recover the matched pairs in order.
  const pairs: Array<[number, number]> = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (aKeys[i] === bKeys[j]) {
      pairs.push([i, j]);
      i++;
      j++;
    } else if (at((i + 1) * width + j) >= at(i * width + j + 1)) {
      i++;
    } else {
      j++;
    }
  }

  return pairs;
}
