import { ThemeToggle } from '../renderer/src/theme/ThemeToggle.js';
import { WebSettingsDialog } from './WebSettingsDialog.js';

/** AppShell's title bar without the macOS traffic-light inset or the drag region. */
export function WebTitleBar({ subtitle }: { subtitle?: string | undefined }) {
  return (
    <header className="flex h-11 shrink-0 items-center gap-3 border-b border-border bg-surface-raised px-3 select-none">
      <div className="flex min-w-0 items-baseline gap-2">
        <span className="text-[13px] font-semibold tracking-tight">Mutatis</span>
        {subtitle ? (
          <span className="truncate text-[11px] text-text-faint">{subtitle}</span>
        ) : null}
      </div>
      <div className="ml-auto flex items-center gap-2">
        <ThemeToggle />
        <WebSettingsDialog />
      </div>
    </header>
  );
}
