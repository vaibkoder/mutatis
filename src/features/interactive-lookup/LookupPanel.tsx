/**
 * Preview panel for Interactive Reference Lookup (feature 04).
 * Renders every LookupResult state (spec §7).
 *
 * Jump-to is the ONLY action that moves the cursor — explicit and user-initiated.
 * The cursor never moves on its own (spec §7.3).
 */

import { useRef, useEffect } from 'preact/hooks';
import { DocModel, resolveAnchor } from '@shared/doc-model/index.js';
import { wordRun } from '@shared/office/index.js';
import { getActiveProfile, getLookupConfig } from '@shared/config/index.js';
import type {
  LookupResult,
  LookupResolved,
  LookupDefByXref,
  LookupAmbiguous,
  DefinedTerm,
} from './types.js';
import type { SectionNode } from '@shared/section-index/types.js';
import type { TextAnchor } from '@shared/doc-model/index.js';
import {
  definitionList,
  activeDefinitionKey,
  isDesktopHost,
  setActiveDefinition,
} from './driver.js';
import type { DefinitionListEntry } from './definition-list.js';

// ── Jump-to helpers ───────────────────────────────────────────────────────────

async function jumpToNode(node: SectionNode): Promise<void> {
  await wordRun(async (ctx) => {
    const range = node.getRange(ctx);
    await ctx.sync();
    range.select();
    await ctx.sync();
  });
}

async function jumpToAnchor(anchor: TextAnchor): Promise<void> {
  const paras = DocModel.allParagraphs.value;
  await wordRun(async (ctx) => {
    const range = resolveAnchor(ctx, paras, anchor);
    await ctx.sync();
    range.select();
    await ctx.sync();
  });
}

// ── Sub-components ────────────────────────────────────────────────────────────

function JumpToButton({ onClick }: { onClick: () => void }) {
  return (
    <button
      type="button"
      class="dp-lookup__jump-btn"
      onClick={onClick}
      title="Jump to this location in the document"
    >
      Jump to →
    </button>
  );
}

/** Clamp preview text to `maxChars`; `Infinity` (the default) shows it in full. */
function clampPreview(text: string, maxChars: number): string {
  return text.length > maxChars ? text.slice(0, maxChars) + '…' : text;
}

function PreviewBlock({
  label,
  previewText,
  onJump,
  maxChars = Infinity,
}: {
  label: string;
  previewText: string;
  onJump: () => void;
  /** Render-time cap. Omit (Infinity) for the expanded card's fuller view. */
  maxChars?: number;
}) {
  return (
    <div class="dp-lookup__block">
      <div class="dp-lookup__block-header">
        <h3 class="dp-lookup__label">{label}</h3>
        <JumpToButton onClick={onJump} />
      </div>
      <div class="dp-lookup__preview-text" aria-readonly="true">
        {clampPreview(previewText, maxChars)}
      </div>
    </div>
  );
}

function ContainerBadge({ text }: { text: string }) {
  return <span class="dp-lookup__container-badge">{text}</span>;
}

// ── State renderers ───────────────────────────────────────────────────────────

function ResolvedPanel({ result }: { result: LookupResolved }) {
  return (
    <PreviewBlock
      label={result.label}
      previewText={result.previewText}
      onJump={() => { jumpToNode(result.node).catch(() => {}); }}
    />
  );
}

function DefinitionPanel({ term, maxChars = Infinity }: { term: DefinedTerm; maxChars?: number }) {
  return (
    <PreviewBlock
      label={term.display}
      previewText={term.previewText}
      onJump={() => { jumpToAnchor(term.anchor).catch(() => {}); }}
      maxChars={maxChars}
    />
  );
}

function DefByXrefPanel({ result, maxChars = Infinity }: { result: LookupDefByXref; maxChars?: number }) {
  return (
    <div class="dp-lookup__stacked">
      <p class="dp-lookup__section-label">Definition</p>
      <DefinitionPanel term={result.term} maxChars={maxChars} />
      {result.resolvedClause && (
        <>
          <p class="dp-lookup__section-label">Referred clause</p>
          <ResolvedPanel result={result.resolvedClause} />
        </>
      )}
    </div>
  );
}

function AmbiguousPanel({ result, maxChars = Infinity }: { result: LookupAmbiguous; maxChars?: number }) {
  return (
    <div class="dp-lookup__ambiguous">
      <p class="dp-lookup__ambiguous-heading">
        {result.candidates.length} possible targets
      </p>
      {result.candidates.map((c, i) => (
        <div
          key={i}
          class={`dp-lookup__block dp-lookup__block--candidate${i === 0 ? ' dp-lookup__block--expanded' : ''}`}
        >
          <div class="dp-lookup__block-header">
            <h3 class="dp-lookup__label">{c.label}</h3>
            <ContainerBadge text={c.containerBadge} />
            <JumpToButton
              onClick={() => {
                if (c.kind === 'section') {
                  jumpToNode(c.node).catch(() => {});
                } else {
                  jumpToAnchor(c.term.anchor).catch(() => {});
                }
              }}
            />
          </div>
          {i === 0 && (
            <div class="dp-lookup__preview-text" aria-readonly="true">
              {clampPreview(c.previewText, maxChars)}
            </div>
          )}
        </div>
      ))}
    </div>
  );
}

// ── Accumulating definition list (feature 04b, desktop only) ─────────────────

const COLLAPSED_TEASER_CHARS = 160;

function teaser(s: string): string {
  const flat = s.replace(/\s+/g, ' ').trim();
  return flat.length > COLLAPSED_TEASER_CHARS
    ? flat.slice(0, COLLAPSED_TEASER_CHARS) + '…'
    : flat;
}

/** One card in the definition list: expanded (full view) when active, else a compact row. */
function DefinitionCard({ entry, active }: { entry: DefinitionListEntry; active: boolean }) {
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    // Scroll the active card into view within the panel. Guarded for jsdom.
    if (active) ref.current?.scrollIntoView?.({ block: 'nearest' });
  }, [active]);

  const term = entry.result.term;

  return (
    <div
      ref={ref}
      class={`dp-lookup__card${active ? ' dp-lookup__card--active' : ''}`}
    >
      {active ? (
        // Active card shows the fuller view (no maxChars → full stored text).
        entry.result.type === 'def-by-xref'
          ? <DefByXrefPanel result={entry.result} />
          : <DefinitionPanel term={term} />
      ) : (
        <button
          type="button"
          class="dp-lookup__card-row"
          onClick={() => setActiveDefinition(entry.termKey)}
          title="Expand this definition"
        >
          <span class="dp-lookup__card-term">{entry.display}</span>
          <span class="dp-lookup__card-teaser">{teaser(term.previewText)}</span>
        </button>
      )}
    </div>
  );
}

function DefinitionList({ list, activeKey }: { list: DefinitionListEntry[]; activeKey: string | null }) {
  return (
    <div class="dp-lookup__deflist">
      <p class="dp-lookup__deflist-heading">
        Definitions ({list.length})
      </p>
      {list.map(entry => (
        <DefinitionCard
          key={entry.termKey}
          entry={entry}
          active={entry.termKey === activeKey}
        />
      ))}
    </div>
  );
}

// ── Main panel ────────────────────────────────────────────────────────────────

/** Skeleton shimmer shown while a lookup is in-flight (before first result).
 *  The driver sets lookupResult to null during the initial load window. */
export function LookupSkeleton() {
  return (
    <section class="dp-lookup" aria-label="Reference preview loading">
      <header class="dp-lookup__header">
        <span class="dp-lookup__title">Reference preview</span>
      </header>
      <div class="dp-lookup__content">
        <div class="dp-lookup__skeleton" aria-hidden="true">
          <div class="dp-lookup__skeleton-line dp-lookup__skeleton-line--title" />
          <div class="dp-lookup__skeleton-line" />
          <div class="dp-lookup__skeleton-line dp-lookup__skeleton-line--short" />
        </div>
      </div>
    </section>
  );
}

export function LookupPanel({ result }: { result: LookupResult }) {
  const list = definitionList.value;
  const activeKey = activeDefinitionKey.value;
  const showList = isDesktopHost.value && list.length > 0;
  const { previewMaxChars } = getLookupConfig(getActiveProfile());

  // The result actually on display (unwrap the Updating wrapper).
  const shown = result.type === 'updating' ? result.stale : result;
  const isDefResult = shown.type === 'definition' || shown.type === 'def-by-xref';

  return (
    <section class="dp-lookup" aria-label="Reference preview">
      <header class="dp-lookup__header">
        <span class="dp-lookup__title">Reference preview</span>
        {result.type === 'updating' && (
          <span class="dp-lookup__updating-chip" aria-live="polite">
            updating definitions…
          </span>
        )}
      </header>
      <div class="dp-lookup__content">
        {showList ? (
          <>
            {/* A current cross-ref/external result shows above the list;
                a current definition is already the active card; idle is suppressed. */}
            {!isDefResult && shown.type !== 'idle' && renderResult(result, previewMaxChars)}
            <DefinitionList list={list} activeKey={activeKey} />
          </>
        ) : (
          renderResult(result, previewMaxChars)
        )}
      </div>
    </section>
  );
}

/**
 * Render a single result. `previewMaxChars` caps definition/ambiguous preview text
 * (spec §7.1) — the expanded *active card* bypasses this for its fuller view.
 */
function renderResult(result: LookupResult, previewMaxChars: number): preact.ComponentChild {
  switch (result.type) {
    case 'idle':
      return (
        <p class="dp-lookup__idle">
          Select a clause reference or defined term to preview it.
        </p>
      );

    case 'resolved':
      return <ResolvedPanel result={result} />;

    case 'definition':
      return <DefinitionPanel term={result.term} maxChars={previewMaxChars} />;

    case 'def-by-xref':
      return <DefByXrefPanel result={result} maxChars={previewMaxChars} />;

    case 'external':
      return (
        <p class="dp-lookup__external">
          Points to an <strong>external document</strong>:{' '}
          <em>{result.docName}</em>.
          <br />
          This add-in can only preview references within the current file.
        </p>
      );

    case 'ambiguous':
      return <AmbiguousPanel result={result} maxChars={previewMaxChars} />;

    case 'no-definition':
      return (
        <p class="dp-lookup__no-def">
          No definition found for{' '}
          <em>"{result.phrase}"</em> in this document.
          It may be defined in an external document or not yet defined.
        </p>
      );

    case 'updating':
      return renderResult(result.stale, previewMaxChars);
  }
}
