/**
 * What a zero-text object change (ADR-0045 § Amendment) looks like to a
 * surface that prints it: the change-list row and the exported report.
 * Spec: spec/features/08-compare.md §8.3.
 *
 * Pure, and free of the XML stack, so a UI component can import it.
 */
import type { PropertyChange, ZeroTextObject } from '../model.js';

/**
 * The zero-text objects one side has and the other lacks, by sort — what the
 * change-list row names (ADR-0045 § Amendment). Aligned on sort and place, so
 * a MOVED object is in both lists: struck where it was, inserted where it is,
 * which is how the redline shows it.
 */
export function objectEdits(change: PropertyChange): { removed: ZeroTextObject[]; added: ZeroTextObject[] } {
  const from = change.from.objects ?? [];
  const to = change.to.objects ?? [];
  const keyOf = (x: ZeroTextObject): string => `${x.sort}@${x.at}/${x.within ?? ''}`;
  const toKeys = to.map(keyOf);
  // Longest common subsequence on sort@place; what is left over changed.
  const len = from.map(() => new Array<number>(to.length + 1).fill(0));
  len.push(new Array<number>(to.length + 1).fill(0));
  for (let i = from.length - 1; i >= 0; i--) {
    for (let j = to.length - 1; j >= 0; j--) {
      len[i]![j] = keyOf(from[i]!) === toKeys[j] ? len[i + 1]![j + 1]! + 1 : Math.max(len[i + 1]![j]!, len[i]![j + 1]!);
    }
  }
  const removed: ZeroTextObject[] = [];
  const added: ZeroTextObject[] = [];
  let i = 0;
  let j = 0;
  while (i < from.length && j < to.length) {
    if (keyOf(from[i]!) === toKeys[j]) {
      i++;
      j++;
    } else if (len[i + 1]![j]! >= len[i]![j + 1]!) {
      removed.push(from[i++]!);
    } else {
      added.push(to[j++]!);
    }
  }
  removed.push(...from.slice(i));
  added.push(...to.slice(j));
  return { removed, added };
}

/**
 * A paragraph's text with each zero-text object written in at its place as
 * `[footnote reference]`, `[picture]` — for a surface that prints both sides
 * of a change whose words are identical (the exported report).
 */
export function textWithObjects(text: string, objects: readonly ZeroTextObject[]): string {
  let out = '';
  let at = 0;
  for (const x of objects) {
    // A negative place is in the manual number, which `normText` has stripped.
    const to = Math.min(Math.max(x.at, 0), text.length);
    out += `${text.slice(at, to)}[${x.sort}]`;
    at = Math.max(at, to);
  }
  return out + text.slice(at);
}
