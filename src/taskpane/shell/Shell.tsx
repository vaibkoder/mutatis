import { useEffect } from 'preact/hooks';
import { signal } from '@preact/signals';
import { TabBar } from './TabBar.js';
import { FindingsList } from './FindingsList.js';
import { SettingsPanel } from './SettingsPanel.js';
import { PreviewRegion } from './PreviewRegion.js';
import type { TabId } from '@shared/findings/index.js';
import { DocModel } from '@shared/doc-model/index.js';
import { LookupPanel, LookupSkeleton } from '@features/interactive-lookup/LookupPanel.js';
import {
  lookupResult,
  lookupDisabled,
  lookupStatusMessage,
  lookupWordApi13Banner,
  activateLookup,
} from '@features/interactive-lookup/driver.js';
import { NumberingPanel } from '@features/numbering/index.js';
import { RepairPanel } from '@features/repair/index.js';
import { ContentPanel } from '@features/content/index.js';
import {
  ComparePanel,
  compareDriverState,
  textCompareDriverState,
  pdfCompareDriverState,
} from '@features/compare/index.js';
import { AwarenessNotice } from '@features/ai-export/index.js';

const _activeTab = signal<TabId>('proof');

// 'review', 'numbering', 'repair', 'content', and 'compare' are excluded — they render via dedicated branches.
const CONTENT_TABS: TabId[] = ['proof'];

function renderReviewTab() {
  if (lookupDisabled.value) {
    return (
      <p class="dp-lookup__disabled">
        {lookupStatusMessage.value ?? 'Reference preview is unavailable in this version of Word.'}
      </p>
    );
  }
  const result = lookupResult.value;
  return (
    <>
      {lookupWordApi13Banner.value && (
        <aside class="dp-lookup__banner" role="note">
          {lookupWordApi13Banner.value}
        </aside>
      )}
      {result === null
        ? <LookupSkeleton />
        : <LookupPanel result={result} />
      }
    </>
  );
}

export function Shell() {
  const activeTab = _activeTab.value;

  // Kick off initial document load and subscribe to events once, on mount.
  useEffect(() => {
    DocModel.load()
      .then(() => activateLookup())
      .catch(() => { /* will retry on next user action */ });
    DocModel.subscribeToEvents();
  }, []);

  return (
    <div class="dp-shell" lang="en-GB">
      <TabBar
        activeTab={activeTab}
        onTabChange={tab => { _activeTab.value = tab; }}
      />
      <main class="dp-shell__content">
        {activeTab === 'review'
          ? renderReviewTab()
          : activeTab === 'numbering'
            ? <NumberingPanel />
            : activeTab === 'repair'
              ? <RepairPanel />
              : activeTab === 'content'
                ? <ContentPanel />
                : activeTab === 'compare'
                  ? <ComparePanel driver={compareDriverState} textDriver={textCompareDriverState} pdfDriver={pdfCompareDriverState} />
                  : CONTENT_TABS.includes(activeTab)
                    ? <FindingsList tabId={activeTab} />
                    : <SettingsPanel />
        }
      </main>
      <PreviewRegion />
      {/* One-time AI-export awareness toast (feature 11) — global so it shows on any tab. */}
      <AwarenessNotice />
    </div>
  );
}
