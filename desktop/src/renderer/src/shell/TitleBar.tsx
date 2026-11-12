import { SettingsDialog } from '../settings/SettingsDialog.js';
import { ThemeToggle } from '../theme/ThemeToggle.js';

/**
 * The titlebar doubles as the window drag region.
 *
 * `-webkit-app-region: drag` swallows ALL pointer events in its box, so every
 * interactive child needs `.no-drag` or it will simply never fire. Never put a
 * context menu here.
 */
// `| undefined` is required, not noise: exactOptionalPropertyTypes is on (it
// mirrors Mutatis's tsconfig so upstream source compiles clean here), and
// under it an optional prop does NOT implicitly accept undefined.
export function TitleBar({ subtitle }: { subtitle?: string | undefined }) {
  return (
    <header
      className="drag-region flex h-11 shrink-0 items-center gap-3 border-b border-border bg-surface-raised pr-3 pl-20 select-none"
    >
      <div className="flex min-w-0 items-baseline gap-2">
        <span className="text-[13px] font-semibold tracking-tight">Mutatis Desktop</span>
        {subtitle ? (
          <span className="truncate text-[11px] text-text-faint">{subtitle}</span>
        ) : null}
      </div>
      <div className="ml-auto flex items-center gap-2">
        <ThemeToggle />
        <SettingsDialog />
      </div>
    </header>
  );
}
