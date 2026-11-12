/**
 * Template feature driver — high-level orchestration.
 *
 * Coordinates the author, fill, conditional, repeating, and dataset
 * sub-modules.  UI panels call driver functions rather than sub-modules
 * directly to keep the panel components thin.
 */

export {
  markVariable,
  markConditionalSection,
  defineRepeatingGroup,
  removeVariable,
  removeConditional,
  removeRepeatingGroup,
  loadSchema,
  saveSchema,
  validateTag,
} from './author.js';

export {
  scanAndReconcile,
  fillQuestionnaire,
} from './fill.js';

export {
  evaluateCondition,
  applyConditionals,
  restoreConditional,
} from './conditional.js';

export {
  fillRepeatingGroup,
} from './repeating.js';

export {
  saveDataSet,
  loadDataSet,
  listDataSets,
  deleteDataSet,
  pushScalarDataSet,
  pushRepeatingDataSet,
  buildFuzzyPreview,
} from './datasets.js';

export {
  getExecutionBlocks,
  renderExecutionBlock,
} from './execution-blocks.js';
