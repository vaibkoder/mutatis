/**
 * apps/web entry (ADR-0022, spec/features/12 phase 2).
 *
 * The standalone counterpart to `src/taskpane/main.tsx`: it installs the BROWSER
 * host adapter (`WebHostPort`) as the active `HostPort` and renders the web app.
 * There is NO Office.js here — no `waitForOffice`, no office.js script in the HTML,
 * no network. The app is a static client-side SPA; the confidentiality invariant
 * holds by construction (ADR-0001/0003).
 */
import { render } from 'preact';
import { setActiveHost } from '@shared/host/index.js';
import { WebHostPort } from '@host-web/index.js';
import { WebApp } from './WebApp.js';
import '../styles/web.css';

const host = new WebHostPort();
setActiveHost(host);

const root = document.getElementById('root');
if (root) {
  render(<WebApp host={host} />, root);
}
