/**
 * Briefing-pack composite button (feature 11 — spec §6.2).
 *
 * Lives in the Proof tab header. Thin wrapper over {@link CopyButton} that
 * builds the composite briefing-pack envelope (#14). Enabled when at least one
 * source checker has current results.
 */
import type { AiExportEnvelope } from '../types.js';
import { CopyButton } from './CopyButton.js';

export interface BriefingPackButtonProps {
  build: () => AiExportEnvelope | null;
  disabledReason?: string;
}

export function BriefingPackButton({ build, disabledReason }: BriefingPackButtonProps) {
  return (
    <div class="dp-ai-export__briefing">
      <CopyButton
        build={build}
        fileLabel="briefing-pack"
        copyLabel="Copy Briefing Pack"
        {...(disabledReason !== undefined ? { disabledReason } : {})}
      />
    </div>
  );
}
