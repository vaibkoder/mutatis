/**
 * apps/web — the minimal in-app document view (spec/features/12 §4 hard-part #3).
 *
 * Renders the loaded paragraphs as a scrollable text view, each element keyed by
 * `data-para-id` so `WebHostPort.reveal(loc)` can scroll + highlight the paragraph
 * a finding points at. This is the "minimum viable view" for analyse-and-export: a
 * located, scrollable text rendering — not a faithful Word render.
 */
import type { ParaNode } from '@shared/doc-model/index.js';

interface DocumentViewProps {
  paragraphs: ReadonlyArray<ParaNode>;
}

export function DocumentView({ paragraphs }: DocumentViewProps) {
  return (
    <div class="dp-web-doc" role="document" aria-label="Imported document">
      {paragraphs.map(para => (
        <p
          key={para.id}
          data-para-id={para.id}
          class={`dp-web-doc__para${para.container === 'footnote' ? ' dp-web-doc__para--footnote' : ''}`}
        >
          {/* Non-breaking space keeps empty paragraphs visible/clickable. */}
          {para.text === '' ? ' ' : para.text}
        </p>
      ))}
    </div>
  );
}
