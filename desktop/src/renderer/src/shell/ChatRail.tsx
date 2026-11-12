import { MessageSquare, PanelLeftClose, PanelLeftOpen, Sparkles } from 'lucide-react';
import { ChatPanel, type ChatPanelProps } from '../chat/ChatPanel.js';

interface ChatRailProps extends ChatPanelProps {
  collapsed: boolean;
  onToggle: () => void;
}

/**
 * Left column. The chrome and the collapse behaviour live here; the chat itself
 * is `ChatPanel`, which reaches the provider only through the preload bridge.
 */
export function ChatRail({ collapsed, onToggle, ...chat }: ChatRailProps) {
  if (collapsed) {
    return (
      <div className="flex h-full flex-col items-center gap-2 border-r border-border bg-surface-sunken py-2">
        <RailButton label="Expand chat panel" onClick={onToggle}>
          <PanelLeftOpen size={15} aria-hidden />
        </RailButton>
        <div className="mt-1 flex h-7 w-7 items-center justify-center text-text-faint">
          <MessageSquare size={15} aria-hidden />
        </div>
      </div>
    );
  }

  return (
    <div className="flex h-full flex-col border-r border-border bg-surface-sunken">
      <div className="flex h-9 shrink-0 items-center gap-2 border-b border-border px-2">
        <Sparkles size={13} className="text-accent" aria-hidden />
        <span className="text-[11px] font-semibold tracking-wide uppercase text-text-muted">
          Assistant
        </span>
        <div className="ml-auto">
          <RailButton label="Collapse chat panel" onClick={onToggle}>
            <PanelLeftClose size={15} aria-hidden />
          </RailButton>
        </div>
      </div>

      <ChatPanel {...chat} />
    </div>
  );
}

function RailButton({
  label,
  onClick,
  children,
}: {
  label: string;
  onClick: () => void;
  children: React.ReactNode;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      aria-label={label}
      title={label}
      className="flex h-7 w-7 items-center justify-center rounded text-text-muted transition-colors hover:bg-border hover:text-text"
    >
      {children}
    </button>
  );
}
