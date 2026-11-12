/**
 * Re-export shim — the numbering scheme parser/resolver moved to the shared
 * numbering module (`@shared/numbering`) when RenderCore (feature 16) made it
 * cross-cutting (one home for shared logic). Host-web keeps
 * importing it from here unchanged. Host-neutral: no Office or Word globals.
 */
export * from '@shared/numbering/numbering-scheme.js';
