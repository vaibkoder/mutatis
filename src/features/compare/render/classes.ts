/**
 * ChangeClass → RcChangeClass mapping (feature 16 / ADR-0024, Phase 2 B3).
 *
 * Compare's classifier emits its own `ChangeClass` union (model.ts); RenderCore
 * owns an identically-named `RcChangeClass` union (render-core/types.ts) so the
 * engine never imports `@features/*` (the `shared ↛ features` wall). The two are
 * a **1:1 name map** — this module is the single place Compare maps into the
 * engine's annotation type, with an exhaustive switch so adding a class on either
 * side fails the build until the map is updated.
 *
 * Annotations are **styling only and never filter** (ADR-0008): mapping classes
 * onto runs changes how a revision is coloured, never whether it renders. A
 * change with no class still renders by its native `<w:ins>/<w:del>/<w:move*>`
 * markup.
 */
import type { ChangeClass } from '../model.js';
import type { RcChangeClass } from '@shared/render-core/index.js';

/** Map one Compare class to its RenderCore counterpart (exhaustive — 1:1 by name). */
export function toRcChangeClass(c: ChangeClass): RcChangeClass {
  switch (c) {
    case 'structural':
    case 'obligationLanguage':
    case 'numericalValue':
    case 'definedTerm':
    case 'crossRefRenumber':
    case 'substantiveText':
    case 'whitespacePunct':
    case 'formatting':
      return c;
    default: {
      // Exhaustiveness guard: a new ChangeClass with no mapping is a compile error.
      const never: never = c;
      return never;
    }
  }
}

/**
 * Convert the driver's `classByRevId` side channel (native `@w:id` → Compare
 * classes) into the engine's `RenderInput.classByRevId` (→ RcChangeClass). Drops
 * nothing — every revision id is carried through; empty class lists pass through
 * as-is (the engine treats a missing/empty list as "render by revision type").
 */
export function toRcClassMap(
  byId: ReadonlyMap<string, readonly ChangeClass[]>,
): ReadonlyMap<string, readonly RcChangeClass[]> {
  const out = new Map<string, readonly RcChangeClass[]>();
  for (const [id, classes] of byId) {
    out.set(id, classes.map(toRcChangeClass));
  }
  return out;
}
