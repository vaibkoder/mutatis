/**
 * Feature 07 — Content tab root panel.
 *
 * Hosts three sub-features behind a sub-tab switcher:
 *   - Contents      → TOC + Index of Schedules  (sub-feature A)
 *   - Clause Library → vetted precedent library  (sub-feature B)
 *   - Templates     → tagged-variable templates  (sub-feature C)
 *
 * Each sub-panel is imported from its own subdir barrel so the three
 * sub-features can be built in isolated worktrees without touching this
 * file.
 */

import { signal } from '@preact/signals';
import { TocPanel } from './toc/index.js';
import { LibraryPanel } from './library/index.js';
import { TemplatesPanel } from './templates/index.js';

type ContentSubTab = 'toc' | 'library' | 'templates';

const _activeSubTab = signal<ContentSubTab>('toc');

const SUB_TABS: { id: ContentSubTab; label: string }[] = [
  { id: 'toc', label: 'Contents' },
  { id: 'library', label: 'Clause Library' },
  { id: 'templates', label: 'Templates' },
];

export function ContentPanel() {
  const active = _activeSubTab.value;

  return (
    <div class="dp-content" lang="en-GB">
      <nav
        class="dp-content__subtabs"
        role="tablist"
        aria-label="Content tools"
      >
        {SUB_TABS.map(tab => {
          const isActive = tab.id === active;
          return (
            <button
              key={tab.id}
              type="button"
              role="tab"
              aria-selected={isActive}
              aria-controls={`dp-content-panel-${tab.id}`}
              id={`dp-content-subtab-${tab.id}`}
              class={`dp-content__subtab${isActive ? ' dp-content__subtab--active' : ''}`}
              onClick={() => { _activeSubTab.value = tab.id; }}
            >
              {tab.label}
            </button>
          );
        })}
      </nav>

      <div
        class="dp-content__panel"
        role="tabpanel"
        id={`dp-content-panel-${active}`}
      >
        {active === 'toc'
          ? <TocPanel />
          : active === 'library'
            ? <LibraryPanel />
            : <TemplatesPanel />}
      </div>
    </div>
  );
}
