import { render } from 'preact';
import { Shell } from './shell/index.js';
import { waitForOffice } from '@shared/office/index.js';
import { setActiveHost } from '@shared/host/index.js';
import { OfficeHostPort } from '@host-office/index.js';
import '../styles/shell.css';

/** Apply the live Office theme to <html data-office-theme>. */
function applyOfficeTheme(): void {
  try {
    const theme = Office.context.officeTheme;
    const bg = theme?.bodyBackgroundColor ?? '#ffffff';
    const isDark = isDarkColour(bg);
    document.documentElement.setAttribute('data-office-theme', isDark ? 'dark' : 'light');
  } catch {
    document.documentElement.setAttribute('data-office-theme', 'light');
  }
}

function isDarkColour(hex: string): boolean {
  const m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(hex);
  if (!m) return false;
  const r = parseInt(m[1]!, 16);
  const g = parseInt(m[2]!, 16);
  const b = parseInt(m[3]!, 16);
  const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  return luminance < 0.5;
}

waitForOffice()
  .then(() => {
    // Install the Office adapter as the active HostPort. UI/feature code reaches
    // the host via getActiveHost() and never imports an adapter directly (ADR-0022).
    setActiveHost(new OfficeHostPort());
    applyOfficeTheme();
    try {
      Office.context.document.addHandlerAsync(
        Office.EventType.OfficeThemeChanged ?? ('officeThemeChanged' as unknown as Office.EventType),
        applyOfficeTheme
      );
    } catch {
      /* theme-changed event not supported on this host */
    }

    const root = document.getElementById('root');
    if (root) {
      render(<Shell />, root);
    }
  })
  .catch((err: unknown) => {
    const root = document.getElementById('root');
    if (root) {
      root.innerHTML = `<pre style="color:red;padding:8px;font-size:11px;white-space:pre-wrap">${String(err)}</pre>`;
    }
  });
