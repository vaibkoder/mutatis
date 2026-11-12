import { useCallback, useState } from 'react';
import { Group, Panel, Separator, useDefaultLayout } from 'react-resizable-panels';
import { WebTitleBar } from './WebTitleBar.js';
import { WorkPane } from '../renderer/src/shell/WorkPane.js';
import { ContextRail } from '../renderer/src/shell/ContextRail.js';
import { layoutStorage } from '../renderer/src/lib/layout-storage.js';
import { useComparisons } from '../renderer/src/compare/useComparisons.js';
import {
  comparisonShape,
  describeShape,
  emptyDocumentSet,
  type DocSide,
  type DocumentSet,
  type LoadedDoc,
} from '../renderer/src/state/documents.js';

/**
 * AppShell (renderer/src/shell/AppShell.tsx) without the chat rail: two columns,
 * work and context. Kept as a copy rather than a prop on AppShell so the Electron
 * shell stays untouched; the panel notes below are inherited from it.
 *
 * Original notes:
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

export function WebShell() {
  const [documents, setDocuments] = useState<DocumentSet>(emptyDocumentSet);
  const {
    tabs,
    compare,
    cancel,
    select,
    close,
    repointRow,
    confirmPairing,
  } = useComparisons();

  const { defaultLayout, onLayoutChanged } = useDefaultLayout({
    id: 'mutatis.web.shell.v1',
    storage: layoutStorage,
    panelIds: ['work', 'context'],
    onlySaveAfterUserInteractions: true,
  });

  const addDocs = useCallback((side: DocSide, docs: readonly LoadedDoc[]) => {
    setDocuments((prev) => ({ ...prev, [side]: [...prev[side], ...docs] }));
  }, []);

  const removeDoc = useCallback((side: DocSide, id: string) => {
    setDocuments((prev) => ({ ...prev, [side]: prev[side].filter((d) => d.id !== id) }));
  }, []);

  const shape = comparisonShape(documents);

  return (
    <div className="flex h-full flex-col">
      <WebTitleBar subtitle={shape === 'none' ? undefined : describeShape(shape)} />

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
