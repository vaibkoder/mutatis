import { signal } from '@preact/signals';

export interface PreviewContent {
  title: string;
  body: string;
}

/** Shared signal for the preview region. Interactive Lookup sets this; shell renders it. */
export const previewContent = signal<PreviewContent | null>(null);

export function PreviewRegion() {
  const content = previewContent.value;
  if (!content) return null;

  return (
    <aside
      class="dp-preview-region"
      aria-label={`Preview: ${content.title}`}
    >
      <header class="dp-preview-region__header">
        <h4 class="dp-preview-region__title">{content.title}</h4>
        <button
          class="dp-preview-region__close"
          aria-label="Close preview"
          onClick={() => { previewContent.value = null; }}
          type="button"
        >
          ✕
        </button>
      </header>
      <div class="dp-preview-region__body" aria-readonly="true">
        {content.body}
      </div>
    </aside>
  );
}
