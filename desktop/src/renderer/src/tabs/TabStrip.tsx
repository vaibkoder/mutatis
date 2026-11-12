import {
  AlertTriangle,
  Bot,
  FileDiff,
  GitCompareArrows,
  Loader2,
  Plus,
  Table2,
  X,
} from 'lucide-react';
import type { Tab, TabId, TabState } from './model.js';

interface TabStripProps {
  state: TabState;
  onSelect: (id: TabId) => void;
  onClose: (id: TabId) => void;
}

export function TabStrip({ state, onSelect, onClose }: TabStripProps) {
  return (
    <div
      role="tablist"
      aria-label="Open comparisons"
      className="flex h-9 shrink-0 items-stretch gap-px overflow-x-auto border-b border-border bg-surface-sunken"
    >
      {state.tabs.map((tab) => (
        <TabButton
          key={tab.id}
          tab={tab}
          active={tab.id === state.activeId}
          onSelect={() => onSelect(tab.id)}
          onClose={() => onClose(tab.id)}
        />
      ))}
    </div>
  );
}

function TabButton({
  tab,
  active,
  onSelect,
  onClose,
}: {
  tab: Tab;
  active: boolean;
  onSelect: () => void;
  onClose: () => void;
}) {
  const closable = tab.kind !== 'landing';

  return (
    <div
      className={
        'group relative flex min-w-0 items-center gap-1.5 border-r border-border px-3 ' +
        (active ? 'bg-surface' : 'bg-transparent hover:bg-border/40')
      }
    >
      <button
        type="button"
        role="tab"
        aria-selected={active}
        onClick={onSelect}
        className="flex min-w-0 items-center gap-1.5 py-1 text-[12px]"
        title={tab.kind === 'redline' ? `${tab.originalName} → ${tab.revisedName}` : tab.title}
      >
        <TabIcon tab={tab} />
        <span
          className={
            'max-w-[18ch] truncate ' + (active ? 'font-medium text-text' : 'text-text-muted')
          }
        >
          {tab.title}
        </span>
      </button>

      {closable ? (
        <button
          type="button"
          onClick={onClose}
          aria-label={`Close ${tab.title}`}
          className="rounded p-0.5 text-text-faint opacity-0 transition hover:bg-border hover:text-text group-hover:opacity-100 focus-visible:opacity-100"
        >
          <X size={12} aria-hidden />
        </button>
      ) : null}

      {active ? (
        <span aria-hidden className="absolute inset-x-0 bottom-0 h-0.5 bg-accent" />
      ) : null}
    </div>
  );
}

function TabIcon({ tab }: { tab: Tab }) {
  if (tab.kind === 'landing') return <Plus size={12} className="text-text-faint" aria-hidden />;
  if (tab.kind === 'matrix') return <Table2 size={12} className="text-text-faint" aria-hidden />;
  if (tab.kind === 'self-redline') {
    return tab.status.state === 'running' ? (
      <Loader2 size={12} className="animate-spin text-accent" aria-hidden />
    ) : tab.status.state === 'failed' ? (
      <AlertTriangle size={12} className="text-danger" aria-hidden />
    ) : (
      <Bot size={12} className="text-accent" aria-hidden />
    );
  }
  if (tab.kind === 'pairing') {
    return tab.status === 'building' ? (
      <Loader2 size={12} className="animate-spin text-accent" aria-hidden />
    ) : tab.status === 'failed' ? (
      <AlertTriangle size={12} className="text-danger" aria-hidden />
    ) : (
      <GitCompareArrows size={12} className="text-text-faint" aria-hidden />
    );
  }
  switch (tab.status.state) {
    case 'running':
      return <Loader2 size={12} className="animate-spin text-accent" aria-hidden />;
    case 'failed':
      return <AlertTriangle size={12} className="text-danger" aria-hidden />;
    case 'cancelled':
      return <X size={12} className="text-text-faint" aria-hidden />;
    case 'ready':
      return <FileDiff size={12} className="text-text-faint" aria-hidden />;
  }
}
