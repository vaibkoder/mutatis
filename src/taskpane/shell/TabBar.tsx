import { Badge } from '../components/Badge.js';
import type { TabId } from '@shared/findings/index.js';
import { tabCounts } from '@shared/findings/index.js';

interface Tab {
  id: TabId;
  label: string;
}

const TABS: Tab[] = [
  { id: 'proof', label: 'Proof' },
  { id: 'review', label: 'Review' },
  { id: 'repair', label: 'Repair' },
  { id: 'numbering', label: 'Numbering' },
  { id: 'content', label: 'Content' },
  { id: 'compare', label: 'Compare' },
  { id: 'settings', label: 'Settings' },
];

interface TabBarProps {
  activeTab: TabId;
  onTabChange: (tab: TabId) => void;
}

export function TabBar({ activeTab, onTabChange }: TabBarProps) {
  const counts = tabCounts.value;

  return (
    <nav
      class="dp-tab-bar"
      role="tablist"
      aria-label="Mutatis features"
    >
      {TABS.map(tab => {
        const count = counts[tab.id] ?? 0;
        const isActive = tab.id === activeTab;
        return (
          <button
            key={tab.id}
            role="tab"
            aria-selected={isActive}
            aria-controls={`dp-panel-${tab.id}`}
            id={`dp-tab-${tab.id}`}
            class={`dp-tab${isActive ? ' dp-tab--active' : ''}`}
            onClick={() => onTabChange(tab.id)}
            type="button"
          >
            <span class="dp-tab__label">{tab.label}</span>
            <Badge count={count} />
          </button>
        );
      })}
    </nav>
  );
}
