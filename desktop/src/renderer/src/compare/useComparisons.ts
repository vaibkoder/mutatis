import { useCallback, useEffect, useRef, useState } from 'react';
import {
  addTab,
  closeTab,
  initialTabState,
  nextMatrixTabId,
  nextPairingTabId,
  nextSelfRedlineTabId,
  nextTabId,
  pairTitle,
  pairsFor,
  selectTab,
  updatePairingTab,
  updateSelfRedlineTab,
  updateTab,
  type MatrixTab,
  type PairingTab,
  type RedlineTab,
  type SelfRedlineTab,
  type TabId,
  type TabState,
} from '../tabs/model.js';
import { splitTrackedChanges } from './tracked-changes.js';
import { comparisonShape, type DocumentSet, type LoadedDoc } from '../state/documents.js';
import { useSettings } from '../settings/SettingsProvider.js';
import { repoint, type PairingResult } from '../pairing/build-pairing.js';
import { buildDeviationMatrix, type MatrixInput } from '../ai/deviation-matrix.js';
import type { CompareOutcome } from './protocol.js';
import type { ProposedEdit } from '../edits/apply-edits.js';
import {
  CompareCancelledError,
  createCompareClient,
  type CompareClient,
} from './client.js';

/**
 * Owns the tab state and the compare worker.
 *
 * One worker for the whole session, created lazily on the first comparison so
 * the app does not pay worker startup just to show the landing page.
 */
/** What to tell the user about how the two documents were read, side by side. */
function notesFor(original: LoadedDoc, revised: LoadedDoc): string[] {
  const out: string[] = [];
  for (const [label, doc] of [['Original', original], ['Modified', revised]] as const) {
    for (const note of doc.notes ?? []) out.push(`${label} (${doc.name}): ${note}`);
  }
  return out;
}

/** The name the assistant sees; says so when the text was not read straight from a Word file. */
function aiLabel(doc: LoadedDoc): string {
  return doc.aiNote ? `${doc.name} (${doc.aiNote})` : doc.name;
}

export function useComparisons() {
  const [tabs, setTabs] = useState<TabState>(initialTabState);
  const clientRef = useRef<CompareClient | null>(null);
  /** Worker job id per tab, so Cancel knows what to cancel. */
  const jobIds = useRef(new Map<TabId, number>());
  /** The document set a pairing tab was built from, so Confirm can resolve ids. */
  const pairingDocs = useRef(new Map<TabId, DocumentSet>());
  /**
   * The two files behind each redline tab, kept so the assistant can edit one.
   *
   * A `File` is a handle, not a copy — retaining it costs almost nothing, and
   * without it an AI edit would mean asking the user to find the document again
   * for a comparison they are already looking at.
   */
  const sourceDocs = useRef(new Map<TabId, { original: LoadedDoc; revised: LoadedDoc }>());
  /** Latest tabs, readable from callbacks without making them stale. */
  const tabsRef = useRef<TabState>(initialTabState);
  tabsRef.current = tabs;

  /**
   * The compare profile, read through a ref.
   *
   * It is a real correctness setting — UK and SG profiles disagree about what
   * counts as a change — so it must be the value in force when a job STARTS,
   * not the one captured when these callbacks were first created. A ref keeps
   * the callbacks stable without letting them go stale.
   */
  const { jurisdiction } = useSettings();
  const jurisdictionRef = useRef(jurisdiction);
  jurisdictionRef.current = jurisdiction;

  useEffect(() => {
    return () => {
      clientRef.current?.terminate();
      clientRef.current = null;
    };
  }, []);

  const getClient = useCallback((): CompareClient => {
    clientRef.current ??= createCompareClient();
    return clientRef.current;
  }, []);

  /** Run a list of already-decided pairs, one tab each. */
  const runPairs = useCallback(
    async (pairs: ReadonlyArray<{ original: LoadedDoc; revised: LoadedDoc }>) => {
      if (pairs.length === 0) return;

      const client = getClient();

      for (const pair of pairs) {
        const id = nextTabId();
        const tab: RedlineTab = {
          kind: 'redline',
          id,
          title: pairTitle(pair.original.name, pair.revised.name),
          originalName: pair.original.name,
          revisedName: pair.revised.name,
          ...(notesFor(pair.original, pair.revised).length > 0
            ? { notes: notesFor(pair.original, pair.revised) }
            : {}),
          status: { state: 'running', progress: 0, step: 'Reading documents' },
          result: null,
        };
        setTabs((prev) => addTab(prev, tab));
        sourceDocs.current.set(id, { original: pair.original, revised: pair.revised });

        try {
          // Read fresh buffers per job: they are TRANSFERRED to the worker and
          // detached here, so a buffer cannot be reused across comparisons.
          const [originalBytes, revisedBytes] = await Promise.all([
            pair.original.file.arrayBuffer(),
            pair.revised.file.arrayBuffer(),
          ]);

          const job = client.compare({
            originalBytes,
            revisedBytes,
            jurisdiction: jurisdictionRef.current,
            originalName: aiLabel(pair.original),
            revisedName: aiLabel(pair.revised),
            onProgress: (pct, step) =>
              setTabs((prev) =>
                updateTab(prev, id, { status: { state: 'running', progress: pct, step } }),
              ),
          });
          jobIds.current.set(id, job.id);

          const outcome = await job.outcome;
          setTabs((prev) => updateTab(prev, id, { status: { state: 'ready' }, result: outcome }));
        } catch (error) {
          if (error instanceof CompareCancelledError) {
            setTabs((prev) => updateTab(prev, id, { status: { state: 'cancelled' } }));
          } else {
            setTabs((prev) =>
              updateTab(prev, id, {
                status: {
                  state: 'failed',
                  message: error instanceof Error ? error.message : String(error),
                },
              }),
            );
          }
        } finally {
          jobIds.current.delete(id);
        }
      }
    },
    [getClient],
  );

  /**
   * Start work for the current selection.
   *
   * 1:1 and 1:N run immediately — there is nothing to decide. Many-to-many goes
   * through the pairing grid first: which document answers which is a judgement,
   * and guessing it silently produces a redline that looks entirely normal and
   * is against the wrong document.
   */
  const compare = useCallback(
    async (documents: DocumentSet) => {
      const shape = comparisonShape(documents);
      if (shape === 'none') return;

      if (shape === 'tracked-changes') {
        const only = documents.original[0] ?? documents.modified[0]!;
        try {
          const { original, modified } = await splitTrackedChanges(only);
          await runPairs([{ original, revised: modified }]);
        } catch (error) {
          setTabs((prev) =>
            addTab(prev, {
              kind: 'redline',
              id: nextTabId(),
              title: only.name,
              originalName: only.name,
              revisedName: only.name,
              status: {
                state: 'failed',
                message: error instanceof Error ? error.message : String(error),
              },
              result: null,
            }),
          );
        }
        return;
      }

      if (shape !== 'many-to-many') {
        await runPairs(pairsFor(documents));
        return;
      }

      const id = nextPairingTabId();
      const tab: PairingTab = {
        kind: 'pairing',
        id,
        title: `Pair ${documents.original.length} × ${documents.modified.length}`,
        status: 'building',
        result: null,
      };
      setTabs((prev) => addTab(prev, tab));
      pairingDocs.current.set(id, documents);

      try {
        const docs = await Promise.all([
          ...documents.original.map(async (d) => ({
            docId: d.id,
            side: 'a' as const,
            fileName: d.name,
            bytes: await d.file.arrayBuffer(),
          })),
          ...documents.modified.map(async (d) => ({
            docId: d.id,
            side: 'b' as const,
            fileName: d.name,
            bytes: await d.file.arrayBuffer(),
          })),
        ]);

        const result = await getClient().proposePairing({
          docs,
          jurisdiction: jurisdictionRef.current,
        });
        setTabs((prev) => updatePairingTab(prev, id, { status: 'ready', result }));
      } catch (error) {
        setTabs((prev) =>
          updatePairingTab(prev, id, {
            status: 'failed',
            message: error instanceof Error ? error.message : String(error),
          }),
        );
      }
    },
    [getClient, runPairs],
  );

  /** Re-point one row. Pure and local — no bytes are re-read. */
  const repointRow = useCallback((tabId: TabId, rowId: string, b: string | null) => {
    setTabs((prev) => {
      const tab = prev.tabs.find((t) => t.id === tabId);
      if (!tab || tab.kind !== 'pairing' || !tab.result) return prev;
      const next = repoint(tab.result as PairingResult, jurisdictionRef.current, rowId, b);
      return updatePairingTab(prev, tabId, { result: next });
    });
  }, []);

  /** Confirm a pairing and run every runnable pair it contains. */
  const confirmPairing = useCallback(
    async (tabId: TabId) => {
      const state = pairingDocs.current.get(tabId);
      const tab = tabsRef.current.tabs.find((t) => t.id === tabId);
      if (!state || !tab || tab.kind !== 'pairing' || !tab.result) return;

      const { proposal } = tab.result as PairingResult;
      const byId = new Map<string, LoadedDoc>();
      for (const d of [...state.original, ...state.modified]) byId.set(d.id, d);

      const pairs = proposal.rows
        .filter((r) => r.a !== null && r.b !== null && r.runnable)
        .map((r) => ({ original: byId.get(r.a!), revised: byId.get(r.b!) }))
        .filter(
          (p): p is { original: LoadedDoc; revised: LoadedDoc } =>
            p.original !== undefined && p.revised !== undefined,
        );

      await runPairs(pairs);
    },
    [runPairs],
  );

  const cancel = useCallback((tabId: TabId) => {
    const jobId = jobIds.current.get(tabId);
    if (jobId !== undefined) clientRef.current?.cancel(jobId);
  }, []);

  /**
   * Every finished comparison, as the assistant sees it. Recomputed from tab
   * state so it can never drift from what is on screen.
   */
  const readyContexts: readonly MatrixInput[] = tabs.tabs
    .filter((tab): tab is RedlineTab => tab.kind === 'redline' && tab.status.state === 'ready')
    .map((tab) => ({
      id: tab.id,
      title: tab.title,
      context: (tab.result as CompareOutcome).context,
    }));

  /**
   * Open the deviation matrix.
   *
   * Computed here, from the engine's change sets, and opened as its own tab. No
   * model is involved in producing it — a model is only ever asked to narrate
   * what this function already decided.
   */
  const openDeviationMatrix = useCallback((generated: string = new Date().toISOString()) => {
    const inputs = tabsRef.current.tabs
      .filter((tab): tab is RedlineTab => tab.kind === 'redline' && tab.status.state === 'ready')
      .map((tab) => ({
        id: tab.id,
        title: tab.title,
        context: (tab.result as CompareOutcome).context,
      }));
    if (inputs.length === 0) return null;

    const matrix = buildDeviationMatrix(inputs, generated);
    const tab: MatrixTab = {
      kind: 'matrix',
      id: nextMatrixTabId(),
      title: `Deviation matrix (${inputs.length})`,
      result: matrix,
    };
    setTabs((prev) => addTab(prev, tab));
    return matrix;
  }, []);

  /**
   * Apply the assistant's edits to one side of a comparison and redline the
   * result against it.
   *
   * Opens a tab immediately, before the work starts, because this is the moment
   * the user needs to see that something is happening TO THEIR DOCUMENT. It
   * returns a summary for the model, so the next turn knows exactly what was
   * applied and what was refused — and can say so rather than assuming it all
   * worked.
   */
  const applyAiEdits = useCallback(
    async (
      tabId: TabId,
      target: 'original' | 'modified',
      edits: readonly ProposedEdit[],
    ): Promise<{ applied: number; refused: { find: string; refusal: string }[]; error?: string }> => {
      const source = sourceDocs.current.get(tabId);
      if (!source) {
        return { applied: 0, refused: [], error: 'That comparison is no longer open.' };
      }
      const doc = target === 'original' ? source.original : source.revised;

      const id = nextSelfRedlineTabId();
      const tab: SelfRedlineTab = {
        kind: 'self-redline',
        id,
        title: `Assistant → ${doc.name.replace(/\.[^.]+$/, '')}`,
        targetName: doc.name,
        status: { state: 'running', progress: 0, step: 'Applying the proposed edits' },
        result: null,
      };
      setTabs((prev) => addTab(prev, tab));

      try {
        const job = getClient().selfRedline({
          targetBytes: await doc.file.arrayBuffer(),
          targetName: doc.name,
          edits,
          jurisdiction: jurisdictionRef.current,
          onProgress: (pct, step) =>
            setTabs((prev) =>
              updateSelfRedlineTab(prev, id, { status: { state: 'running', progress: pct, step } }),
            ),
        });
        jobIds.current.set(id, job.id);

        const result = await job.result;
        setTabs((prev) =>
          updateSelfRedlineTab(prev, id, { status: { state: 'ready' }, result }),
        );
        return {
          applied: result.applied.length,
          refused: result.refused.map((edit) => ({ find: edit.find, refusal: edit.refusal })),
        };
      } catch (error) {
        const message =
          error instanceof CompareCancelledError
            ? 'The user cancelled this edit.'
            : error instanceof Error
              ? error.message
              : String(error);
        setTabs((prev) =>
          updateSelfRedlineTab(prev, id, { status: { state: 'failed', message } }),
        );
        return { applied: 0, refused: [], error: message };
      }
    },
    [getClient],
  );

  const select = useCallback((id: TabId) => setTabs((prev) => selectTab(prev, id)), []);
  const close = useCallback((id: TabId) => {
    jobIds.current.delete(id);
    pairingDocs.current.delete(id);
    sourceDocs.current.delete(id);
    setTabs((prev) => closeTab(prev, id));
  }, []);

  return {
    tabs,
    compare,
    cancel,
    select,
    close,
    repointRow,
    confirmPairing,
    readyContexts,
    openDeviationMatrix,
    applyAiEdits,
  };
}
