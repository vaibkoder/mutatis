import { useCallback, useState } from 'react';
import { Group, Panel, Separator, useDefaultLayout, usePanelRef } from 'react-resizable-panels';
import { TitleBar } from './TitleBar.js';
import { ChatRail } from './ChatRail.js';
import { WorkPane } from './WorkPane.js';
import { ContextRail } from './ContextRail.js';
import { layoutStorage } from '../lib/layout-storage.js';
import { useComparisons } from '../compare/useComparisons.js';
import { activeTab } from '../tabs/model.js';
import {
  comparisonShape,
  describeShape,
  emptyDocumentSet,
  type DocSide,
  type DocumentSet,
  type LoadedDoc,
} from '../state/documents.js';

/**
 * The three-column shell.
 *
 * EVERY size below carries an explicit unit. In react-resizable-panels v4 a bare
 * number means PIXELS and a bare string means PERCENT — so v3's `minSize={20}`
 * (20 percent) silently becomes 20 pixels with no type error and no runtime
 * warning. Explicit units make the intent unambiguous and immune to that.
 *
 * `Separator` elements are DIRECT children of `Group`. The library computes
 * ordering geometrically from offsetLeft/offsetTop, so wrapping one in a div
 * does not throw — it just quietly breaks the layout.
 */

/** Matches ChatRail's collapsed icon strip. */
const CHAT_COLLAPSED_PX = 52;

export function AppShell() {
  const [documents, setDocuments] = useState<DocumentSet>(emptyDocumentSet);
  const [chatCollapsed, setChatCollapsed] = useState(false);
  const chatPanel = usePanelRef();
  const {
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
  } = useComparisons();

  // The comparison the user is actually looking at. The assistant is given this
  // one in full and the others as totals — see ai/prompt.ts.
  const active = activeTab(tabs);
  const activeContext =
    active.kind === 'redline' && active.status.state === 'ready'
      ? (readyContexts.find((entry) => entry.id === active.id) ?? null)
      : null;

  // The assistant can only edit a document that belongs to the comparison in
  // front of the user. Editing something they are not looking at would be a
  // surprise, and "which document" would stop having an obvious answer.
  const editableTab = active.kind === 'redline' ? active : null;
  const names =
    editableTab !== null
      ? { original: editableTab.originalName, modified: editableTab.revisedName }
      : null;

  const { defaultLayout, onLayoutChanged } = useDefaultLayout({
    id: 'mutatis.shell.v1',
    storage: layoutStorage,
    panelIds: ['chat', 'work', 'context'],
    onlySaveAfterUserInteractions: true,
  });

  const addDocs = useCallback((side: DocSide, docs: readonly LoadedDoc[]) => {
    setDocuments((prev) => ({ ...prev, [side]: [...prev[side], ...docs] }));
  }, []);

  const removeDoc = useCallback((side: DocSide, id: string) => {
    setDocuments((prev) => ({ ...prev, [side]: prev[side].filter((d) => d.id !== id) }));
  }, []);

  const toggleChat = useCallback(() => {
    const panel = chatPanel.current;
    if (!panel) return;
    if (panel.isCollapsed()) panel.expand();
    else panel.collapse();
  }, [chatPanel]);

  const shape = comparisonShape(documents);

  return (
    <div className="flex h-full flex-col">
      <TitleBar subtitle={shape === 'none' ? undefined : describeShape(shape)} />

      <Group
        orientation="horizontal"
        className="min-h-0 flex-1"
        defaultLayout={defaultLayout}
        onLayoutChanged={onLayoutChanged}
        // Without this the library injects a <style> element for its drag
        // cursors, which `style-src 'self'` blocks. We supply the cursor rules
        // ourselves in index.css — the same trade Base UI's CSPProvider makes.
        disableCursor
      >
        <Panel
          id="chat"
          panelRef={chatPanel}
          collapsible
          collapsedSize={`${CHAT_COLLAPSED_PX}px`}
          defaultSize="22%"
          minSize="16%"
          maxSize="38%"
          // onCollapse/onExpand were REMOVED in v4 — collapse state is derived
          // from the resize callback instead.
          onResize={(size) => setChatCollapsed(size.inPixels <= CHAT_COLLAPSED_PX + 1)}
        >
          <ChatRail
            collapsed={chatCollapsed}
            onToggle={toggleChat}
            contexts={readyContexts}
            activeContext={activeContext}
            onDeviationMatrix={() => openDeviationMatrix()}
            names={names}
            onApplyEdits={(target, edits) =>
              editableTab !== null
                ? applyAiEdits(editableTab.id, target, edits)
                : Promise.resolve({
                    applied: 0,
                    refused: [],
                    error: 'No comparison is open, so there is no document to edit.',
                  })
            }
          />
        </Panel>

        <Separator className="w-px" />

        <Panel id="work" minSize="30%">
          <WorkPane
            documents={documents}
            tabs={tabs}
            onAdd={addDocs}
            onRemove={removeDoc}
            onCompare={() => void compare(documents)}
            onSelectTab={select}
            onCloseTab={close}
            onCancelTab={cancel}
            onRepoint={repointRow}
            onConfirmPairing={(id) => void confirmPairing(id)}
            onComparePair={(original, modified) =>
              void compare({ original: [original], modified: [modified] })
            }
          />
        </Panel>

        <Separator className="w-px" />

        <Panel id="context" collapsible collapsedSize="0px" defaultSize="20%" minSize="14%" maxSize="32%">
          <ContextRail documents={documents} tabs={tabs} />
        </Panel>
      </Group>
    </div>
  );
}
