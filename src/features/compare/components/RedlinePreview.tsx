/**
 * Feature 16 (RenderCore, Phase 2 — C2): the in-pane paginated redline preview.
 *
 * Upgrades the feature-14 inline token-DOM preview (ADR-0019) to RenderCore's
 * paginated layout: the SAME render source that backs the vector-PDF snapshot is
 * laid out off the main thread (worker) and painted as page frames. Read-only —
 * a complement to the editable `.docx`, never a replacement.
 *
 * ADR-0008 (never-suppress): the layout renders every native revision; class
 * annotations are styling only. Nothing is hidden in the preview.
 *
 * The mount is async (layout runs in the worker), so it lives in a `useEffect`
 * with a cancellation guard + `destroy()` cleanup. `client` is injectable so unit
 * tests drive the real pipeline in-process without a real Worker.
 */
import { useEffect, useRef } from 'preact/hooks';
import type { RenderClient, RenderTheme, PreviewHandle } from '@shared/render-core/index.js';
import type { RenderSource } from '../ui-types.js';

interface RedlinePreviewProps {
  /** The surfaced render source. Null until a comparison has produced one. */
  src: RenderSource | null;
  /** Injected render client (tests). Production uses the shared session worker. */
  client?: RenderClient;
  /** Optional theme passthrough. */
  theme?: RenderTheme;
}

export function RedlinePreview({ src, client, theme }: RedlinePreviewProps) {
  const hostRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const el = hostRef.current;
    if (!el || !src) return;

    let handle: PreviewHandle | null = null;
    let cancelled = false;

    const opts = {
      ...(client !== undefined ? { client } : {}),
      ...(theme !== undefined ? { theme } : {}),
    };
    // Lazy-import the render path (engine + clone-font bytes + worker) so it is
    // code-split out of the main bundle — the preview only loads it on demand (D1).
    void import('../render/index.js')
      .then(({ mountRedlinePreview }) => mountRedlinePreview(el, src, opts))
      .then(h => {
        // Effect re-ran / unmounted before layout finished — discard this mount.
        if (cancelled) { h.destroy(); return; }
        handle = h;
      })
      .catch(() => {
        /* Preview is a best-effort read-only view; the PDF export surfaces errors. */
      });

    return () => {
      cancelled = true;
      if (handle) handle.destroy();
      // Drop any frames already appended (cleanup before the next mount).
      el.replaceChildren();
    };
  }, [src, client, theme]);

  return (
    <section class="dp-compare__preview" aria-label="Rendered redline preview" lang="en-GB">
      <h3 class="dp-compare__preview-title">Redline preview</h3>
      <div ref={hostRef} class="dp-compare__preview-host" />
    </section>
  );
}
