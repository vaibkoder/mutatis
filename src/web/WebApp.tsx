/**
 * apps/web — the standalone web app shell (ADR-0022, spec/features/12 §5 phase 3b).
 *
 * Pick a `.docx` → parse it into the shared DocSnapshot via the host adapter (which,
 * since phase 3a, also renders list numbers from numbering.xml) → run the UNCHANGED
 * checkers → render feature-grouped findings (`WebFindings`, reusing `FindingRow`) →
 * export the document back as a `.docx` download.
 *
 * Checkers lit up here (all client-side):
 *   - Factual Consistency (phase 2 PoC).
 *   - Defined Terms (`analyseDefinedTerms`) — pure analysis.
 *   - Cross-Reference + Static-Numbering flag (`analyseCrossReference(..., null)`):
 *     `ctx=null` is the ADR-0014 graceful-degradation path — no REF-field/OOXML
 *     probes (those modules are dynamic-imported only with a live ctx), so the web
 *     bundle stays Office-free. Yields both the cross-reference findings AND the
 *     static-numbering ("manually-typed clause number") findings.
 *   - Party-Role Consistency (`analysePartyRole`) — pure analysis; ports as-is
 *     (spec/features/12 §2.1 Tier ✅). No banner / no capability degradation.
 *
 * The whole pipeline runs client-side; nothing leaves the machine (ADR-0001/0003).
 */
import { useState } from 'preact/hooks';
import type { ParaNode } from '@shared/doc-model/index.js';
import type { Finding } from '@shared/findings/index.js';
import type { JurisdictionProfile } from '@shared/config/index.js';
import { getActiveProfile } from '@shared/config/index.js';
import { runFactualConsistency } from '@features/factual-consistency/index.js';
import { analyseDefinedTerms, type DefinedTermsAnalysis } from '@features/defined-terms/index.js';
import { analyseCrossReference, type CrossReferenceAnalysis } from '@features/cross-reference/index.js';
import { analysePartyRole } from '@features/party-role/index.js';
import type { WebHostPort } from '@host-web/index.js';
import { WebBundleCompare } from './WebBundleCompare.js';
import { DocumentView } from './DocumentView.js';
import { WebFindings } from './WebFindings.js';
import { WebCompare } from './WebCompare.js';
import { WebBatchCompare } from './WebBatchCompare.js';
import { ThemeToggle } from './ThemeToggle.js';
import { IconDocument } from './icons.js';

interface WebAppProps {
  host: WebHostPort;
}

type Status =
  | { kind: 'idle' }
  | { kind: 'loading' }
  | { kind: 'ready' }
  | { kind: 'error'; message: string };

/**
 * Top-level view: proof one document, compare two, or compare one template
 * against many, and bundle-against-bundle (ADR-0037 — web target only; the
 * add-in has no batch surface, and ADR-0003's two-document ceiling is why).
 */
type View = 'proof' | 'compare' | 'batch' | 'bundles';

/** Held in one place because the header renders it twice — visibly and as `title`. */
const TAGLINE =
  "Run Mutatis's checks on a document without Word — fully on your machine, nothing uploaded.";

/**
 * Run the standalone checkers over one snapshot and return the combined findings
 * plus the two analyses the banners read. Cross-reference runs with a null context
 * (ADR-0014 degraded mode — no live Word host on web). Party-role is pure analysis
 * with no banner, so only its findings are surfaced.
 */
async function runAllChecks(
  paras: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
): Promise<{ findings: Finding[]; defined: DefinedTermsAnalysis; xref: CrossReferenceAnalysis }> {
  const factual = runFactualConsistency(paras, profile);
  const defined = analyseDefinedTerms(paras, profile);
  const partyRole = analysePartyRole(paras, profile);
  const xref = await analyseCrossReference(paras, profile, null);
  const findings = [
    ...xref.xrefFindings,
    ...xref.staticNumFindings,
    ...defined.findings,
    ...partyRole.findings,
    ...factual,
  ];
  return { findings, defined, xref };
}

export function WebApp({ host }: WebAppProps) {
  const [paragraphs, setParagraphs] = useState<ReadonlyArray<ParaNode>>([]);
  const [findings, setFindings] = useState<Finding[]>([]);
  const [defined, setDefined] = useState<DefinedTermsAnalysis | null>(null);
  const [xref, setXref] = useState<CrossReferenceAnalysis | null>(null);
  const [fileName, setFileName] = useState<string>('');
  const [status, setStatus] = useState<Status>({ kind: 'idle' });
  const [view, setView] = useState<View>('proof');

  function applyChecks(result: { findings: Finding[]; defined: DefinedTermsAnalysis; xref: CrossReferenceAnalysis }) {
    setFindings(result.findings);
    setDefined(result.defined);
    setXref(result.xref);
  }

  async function handleFile(e: Event) {
    const input = e.currentTarget as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;
    setStatus({ kind: 'loading' });
    try {
      const buffer = await file.arrayBuffer();
      host.ingest(buffer, file.name);
      const snapshot = await host.loadDocument();
      const all = [...snapshot.paragraphs, ...snapshot.footnotes];
      setParagraphs(all);
      setFileName(file.name);
      // Analyse-on-open — run every checker.
      applyChecks(await runAllChecks(all, getActiveProfile()));
      setStatus({ kind: 'ready' });
    } catch (err) {
      setStatus({ kind: 'error', message: err instanceof Error ? err.message : String(err) });
    }
  }

  async function handleRecheck() {
    if (paragraphs.length === 0) return;
    applyChecks(await runAllChecks(paragraphs, getActiveProfile()));
  }

  async function handleDownload() {
    try {
      const bundle = await host.readDocx();
      await host.writeBack(bundle);
    } catch {
      // Best-effort; no document or unsupported environment.
    }
  }

  const ready = status.kind === 'ready';

  return (
    <div class="dp-web">
      <header class="dp-web__header">
        <div class="dp-web__brand">
          <h1 class="dp-web__title">Mutatis</h1>
          {/* The tagline truncates on a narrow window, so it carries `title`:
              it states the confidentiality promise, and an ellipsis must never
              be the only copy of that. */}
          <p class="dp-web__tagline" title={TAGLINE}>
            {TAGLINE}
          </p>
        </div>
        <ThemeToggle />
      </header>

      {/* Top-level view switch: proof one document, or compare two. */}
      <div class="dp-web__view-toggle" role="tablist" aria-label="Mutatis mode">
        <button
          type="button"
          role="tab"
          aria-selected={view === 'proof' ? 'true' : 'false'}
          class={`dp-web__view-btn${view === 'proof' ? ' dp-web__view-btn--active' : ''}`}
          onClick={() => setView('proof')}
        >
          Proof
        </button>
        <button
          type="button"
          role="tab"
          aria-selected={view === 'compare' ? 'true' : 'false'}
          class={`dp-web__view-btn${view === 'compare' ? ' dp-web__view-btn--active' : ''}`}
          onClick={() => setView('compare')}
        >
          Compare
        </button>
        <button
          type="button"
          role="tab"
          aria-selected={view === 'batch' ? 'true' : 'false'}
          class={`dp-web__view-btn${view === 'batch' ? ' dp-web__view-btn--active' : ''}`}
          onClick={() => setView('batch')}
        >
          Compare many
        </button>
        <button
          type="button"
          role="tab"
          aria-selected={view === 'bundles' ? 'true' : 'false'}
          class={`dp-web__view-btn${view === 'bundles' ? ' dp-web__view-btn--active' : ''}`}
          onClick={() => setView('bundles')}
        >
          Compare bundles
        </button>
      </div>

      {view === 'compare' && <WebCompare />}

      {view === 'batch' && <WebBatchCompare />}

      {view === 'bundles' && <WebBundleCompare />}

      {view === 'proof' && (
        <ProofView
          fileName={fileName}
          status={status}
          paragraphs={paragraphs}
          findings={findings}
          defined={defined}
          xref={xref}
          ready={ready}
          onFile={handleFile}
          onRecheck={handleRecheck}
          onDownload={handleDownload}
        />
      )}
    </div>
  );
}

interface ProofViewProps {
  fileName: string;
  status: Status;
  paragraphs: ReadonlyArray<ParaNode>;
  findings: Finding[];
  defined: DefinedTermsAnalysis | null;
  xref: CrossReferenceAnalysis | null;
  ready: boolean;
  onFile: (e: Event) => void;
  onRecheck: () => void;
  onDownload: () => void;
}

/** The single-document proof flow (file → checkers → findings → download). */
function ProofView({
  fileName,
  status,
  paragraphs,
  findings,
  defined,
  xref,
  ready,
  onFile,
  onRecheck,
  onDownload,
}: ProofViewProps) {
  return (
    <>
      <div class="dp-web__toolbar">
        <label class="dp-web__file">
          <span class="dp-web__file-btn">Choose .docx…</span>
          <input
            type="file"
            accept=".docx"
            class="dp-web__file-input"
            onChange={onFile}
            aria-label="Choose a .docx document to check"
          />
        </label>
        {fileName && <span class="dp-web__filename">{fileName}</span>}
        <span class="dp-web__spacer" />
        <button
          type="button"
          class="dp-web__btn"
          onClick={onRecheck}
          disabled={!ready}
        >
          Re-check
        </button>
        <button
          type="button"
          class="dp-web__btn"
          onClick={onDownload}
          disabled={!ready}
        >
          Download .docx
        </button>
      </div>

      {status.kind === 'loading' && <p class="dp-web__status">Reading document…</p>}
      {status.kind === 'error' && (
        <p class="dp-web__status dp-web__status--error">Could not read this file: {status.message}</p>
      )}
      {status.kind === 'idle' && (
        <p class="dp-web__status">Choose a Word <code>.docx</code> file to begin.</p>
      )}

      {ready && (
        <div class="dp-web__panes">
          <section class="dp-web__doc-pane" aria-label="Document">
            {/* The icon repeats the visible label, so it is decorative. */}
            <div class="dp-web__pane-head">
              <IconDocument />
              <h2 class="dp-web__pane-title">Document</h2>
            </div>
            <div class="dp-web__pane-body">
              <DocumentView paragraphs={paragraphs} />
            </div>
          </section>
          <section class="dp-web__findings-pane" aria-label="Findings">
            <WebFindings findings={findings} defined={defined} xref={xref} />
          </section>
        </div>
      )}
    </>
  );
}
