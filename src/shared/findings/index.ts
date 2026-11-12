export type { Finding, FindingLocation, SuggestedAction, Severity, TabId } from './types.js';
export { CATEGORY_TAB_MAP, findingToTab } from './types.js';
export {
  allFindings,
  tabCounts,
  addFinding,
  addFindings,
  removeFinding,
  removeByPrefix,
  clearFindings,
  findingsForTab,
  findingsByCategory,
} from './store.js';
