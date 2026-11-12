/**
 * Assign-source picker (spec §8.3).
 *
 * Searchable list of SectionIndex.ordered entries. For Ambiguous findings,
 * candidateTargets are pre-surfaced at the top.
 *
 * v1: replaces matched range text with the target's displayLabel (static text).
 * Seam for v2 REF-field strategy is present in assign-source.ts but not exposed here.
 */

import { useState } from 'preact/hooks';
import type { SectionNode } from '@shared/section-index/types.js';
import type { TextAnchor } from '@shared/doc-model/anchor.js';
import type { ParaNode } from '@shared/doc-model/types.js';
import { wordRun } from '@shared/office/index.js';
import { applyAssignSource, filterCompatibleTargets } from '@features/cross-reference/assign-source.js';

interface AssignSourcePickerProps {
  /** All index nodes (ordered). */
  nodes: ReadonlyArray<SectionNode>;
  /** Pre-seeded candidates (Ambiguous finding). */
  candidateTargets?: SectionNode[];
  /** Primary unit type — used to surface compatible entries first. */
  unitType?: string | undefined;
  anchor: TextAnchor;
  paras: ReadonlyArray<ParaNode>;
  onComplete?: () => void;
  onCancel?: () => void;
}

export function AssignSourcePicker({
  nodes,
  candidateTargets = [],
  unitType,
  anchor,
  paras,
  onComplete,
  onCancel,
}: AssignSourcePickerProps) {
  const [query, setQuery] = useState('');
  const [selected, setSelected] = useState<SectionNode | null>(
    candidateTargets.length === 1 ? (candidateTargets[0] ?? null) : null,
  );
  const [applying, setApplying] = useState(false);

  const baseList = unitType
    ? filterCompatibleTargets(nodes as SectionNode[], unitType)
    : (nodes as SectionNode[]);

  // Candidates at top, then filtered remainder.
  const candidateKeys = new Set(candidateTargets.map(n => n.key));
  const candidates = candidateTargets;
  const rest = baseList.filter(n => !candidateKeys.has(n.key));

  const filtered = [...candidates, ...rest].filter(n => {
    if (!query) return true;
    return n.displayLabel.toLowerCase().includes(query.toLowerCase());
  });

  async function handleConfirm() {
    if (!selected || applying) return;
    setApplying(true);
    try {
      await wordRun(async ctx => {
        await applyAssignSource(ctx, paras, anchor, selected);
      });
      onComplete?.();
    } catch {
      // Leave open on error so the user can retry.
    } finally {
      setApplying(false);
    }
  }

  return (
    <div class="dp-assign-source" role="dialog" aria-label="Assign source" aria-modal="true">
      <header class="dp-assign-source__header">
        <span class="dp-assign-source__title">Assign source</span>
        <button class="dp-assign-source__close" type="button" onClick={onCancel} aria-label="Cancel">✕</button>
      </header>

      <input
        class="dp-assign-source__search"
        type="search"
        placeholder="Search…"
        value={query}
        onInput={(e) => setQuery((e.target as HTMLInputElement).value)}
        aria-label="Search numbered units"
      />

      <ul class="dp-assign-source__list" role="listbox" aria-label="Numbered units">
        {filtered.map(node => (
          <li
            key={node.key}
            class={`dp-assign-source__item${selected?.key === node.key ? ' dp-assign-source__item--selected' : ''}${candidateKeys.has(node.key) ? ' dp-assign-source__item--candidate' : ''}`}
            role="option"
            aria-selected={selected?.key === node.key}
            onClick={() => setSelected(node)}
          >
            <span class="dp-assign-source__label">{node.displayLabel}</span>
          </li>
        ))}
        {filtered.length === 0 && (
          <li class="dp-assign-source__empty">No matches</li>
        )}
      </ul>

      <footer class="dp-assign-source__footer">
        <button
          class="dp-assign-source__confirm"
          type="button"
          disabled={!selected || applying}
          onClick={handleConfirm}
          aria-busy={applying}
        >
          {applying ? 'Applying…' : 'Replace with selected'}
        </button>
        <button class="dp-assign-source__cancel" type="button" onClick={onCancel}>
          Cancel
        </button>
      </footer>
    </div>
  );
}
