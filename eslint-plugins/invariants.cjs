/**
 * Custom ESLint rules that enforce the three catastrophic invariants.
 * These are deterministic walls — they fail the build if violated.
 *
 * Rule 1 — no-network-egress (confidentiality, ADR-0001/ADR-0003)
 * Rule 2 — no-us-literals    (jurisdiction,    ADR-0002 / CONVENTIONS.md)
 * Rule 3 — no-cut-feature-refs (scope,         ADR-0005)
 * Rule 4 — no-host-imports   (host portability, ADR-0022 / spec/features/12)
 */

// Network primitives that must never appear outside the sanctioned storage adapters.
const NETWORK_PRIMITIVES = new Set([
  'fetch',
  'XMLHttpRequest',
  'WebSocket',
  'EventSource',
]);

const NETWORK_IMPORT_PATTERNS = [
  /^node:https?$/,
  /^axios$/,
  /^got$/,
  /^ky$/,
  /^node-fetch$/,
  /^cross-fetch$/,
  /^isomorphic-fetch$/,
  /^navigator\.sendBeacon/,
];

// Files in the storage adapter path are allow-listed for network calls.
// (The adapters carry only user-authored library data, never document content.)
function isAllowListed(filename) {
  if (!filename) return false;
  return (
    filename.includes('src/shared/storage/adapters') ||
    filename.includes('src/shared/storage/github') ||
    filename.includes('src/shared/storage/sharepoint')
  );
}

// US-centric product-facing literals that must never appear in non-config source.
// These must come from the Config Layer (spec/CONVENTIONS.md).
const US_LITERAL_PATTERNS = [
  { pattern: /\bSection\s+\d/i, msg: 'Use "Clause" not "Section" as default clause noun (ADR-0002 / CONVENTIONS.md §Terminology). Read from the Config Layer.' },
  { pattern: /\bExhibit\s+[A-Z\d]/i, msg: 'Use "Schedule" not "Exhibit" as default annexure noun (ADR-0002 / CONVENTIONS.md §Terminology). Read from the Config Layer.' },
  { pattern: /\bMM\/DD\/YYYY\b/, msg: 'Use DD/MM/YYYY or DD Month YYYY date format (ADR-0002 / CONVENTIONS.md §Dates).' },
  { pattern: /["']USD["'](?!\s*:)/, msg: 'Never default to USD. Bare $ is ambiguous in SG/UK context — use the Config Layer currency table (CONVENTIONS.md §Currency).' },
  { pattern: /Bluebook|OSCOLA|§\s*\d+\s*U\.S\.C/, msg: 'Bluebook/OSCOLA citations are cut (ADR-0005). SG/UK transactional scope only.' },
];

// Identifiers that reference cut features (ADR-0005).
// These should never appear in src/features/** or src/shared/**.
const CUT_FEATURE_IDENTIFIERS = [
  'TableOfAuthorities',
  'tableOfAuthorities',
  'bluebookCite',
  'oscola',
  'executionConsistency',
  'partyConsistency',
  'textboxDetect',
  'splitScreen',
  'realtimeFeedback',
  'liveRefField',
  'perAuthorAccept',
  'oneDriveSchemeSync',
  'autoRenumber',
  'convertToListNumbering',
  'staticToDynamic',
  'mlSemanticEquivalence',
  'semanticEquivalence',
  'pyodide',
];

// The global namespaces that betray a host dependency. They are AMBIENT globals
// from @types/office-js (never imported), so the check is on references to the
// `Word` / `Office` identifier as the root of a member or type-qualified name —
// it deliberately ignores comments and string literals.
const HOST_GLOBALS = new Set(['Word', 'Office']);

// Allow-list for `no-host-imports` (ADR-0022). The host-neutral feature core +
// UI must never touch Office JS; only the host adapters and the as-yet-unported
// write/navigate seams may. This list is a RATCHET — entries are REMOVED as each
// feature is ported behind HostPort (spec/features/12 phases 2+). Do not add to
// it to silence a lint error in core code: make that code host-neutral instead.
//
// Deliberately ABSENT (= protected, must stay host-free):
//   src/shared/host/**           — the port itself
//   src/features/factual-consistency/**  — the standalone proof-of-concept feature
//   src/features/defined-terms/**, party-role/** — already host-free analysis
//   most of src/shared/** (findings/index, section-index/builder, config/index, …)
const HOST_ALLOWED = [
  // host adapter(s)
  'src/host-office/',
  // shared seams still backed by Office JS (ingest / navigation / storage / write)
  'src/shared/office/',
  'src/shared/ooxml/',
  'src/shared/doc-model/',
  'src/shared/storage/doc-settings',
  'src/shared/findings/types',      // SuggestedAction.apply(ctx) — write path, ported later
  'src/shared/section-index/types', // SectionNode.getRange — navigation, ported later
  'src/shared/config/types',
  // unported feature write/navigate paths (Tier ⚠️/host-coupled in spec §2.1)
  // NOTE: src/features/compare/ was REMOVED here in 12f — Compare's three host
  // touchpoints (active-doc read, redline write, goToChange navigation) now route
  // through the HostPort seam (getActiveHost()), so the feature is host-free.
  'src/features/content/',
  'src/features/cross-reference/',
  'src/features/interactive-lookup/',
  'src/features/numbering/',
  'src/features/repair/',
  // task-pane host wiring (the add-in shell that installs/uses the office host)
  'src/taskpane/main',
  'src/taskpane/shell/Shell',
  'src/taskpane/shell/FindingsList',
  'src/taskpane/shell/ProofExports',
];

function isHostAllowed(filename) {
  if (!filename) return false;
  const f = filename.replace(/\\/g, '/');
  return HOST_ALLOWED.some(p => f.includes(p));
}

// ─── Rule: no-network-egress ─────────────────────────────────────────────────

const noNetworkEgress = {
  meta: {
    type: 'problem',
    docs: {
      description:
        'Disallow network primitives outside the storage-adapter allow-list (confidentiality invariant — ADR-0001/ADR-0003). Document content must never leave the client.',
      recommended: true,
    },
    schema: [],
    messages: {
      networkPrimitive:
        'Network call to "{{name}}" is forbidden outside src/shared/storage/adapters* (ADR-0001/ADR-0003). ' +
        'All document content must stay client-side. ' +
        'If this is a storage-adapter carrying library data (not document content), move the file to src/shared/storage/adapters*.',
      networkImport:
        'Import of "{{source}}" introduces a network dependency (ADR-0001/ADR-0003). ' +
        'Use Office JS APIs or the OOXML Service for all document operations; never send document bytes over the network.',
    },
  },
  create(context) {
    const filename = context.getFilename();
    if (isAllowListed(filename)) return {};

    return {
      // fetch(...), XMLHttpRequest(), WebSocket(...)
      NewExpression(node) {
        if (node.callee.type === 'Identifier' && NETWORK_PRIMITIVES.has(node.callee.name)) {
          context.report({ node, messageId: 'networkPrimitive', data: { name: node.callee.name } });
        }
      },
      CallExpression(node) {
        if (node.callee.type === 'Identifier' && NETWORK_PRIMITIVES.has(node.callee.name)) {
          context.report({ node, messageId: 'networkPrimitive', data: { name: node.callee.name } });
        }
        // navigator.sendBeacon(...)
        if (
          node.callee.type === 'MemberExpression' &&
          node.callee.object.type === 'Identifier' &&
          node.callee.object.name === 'navigator' &&
          node.callee.property.type === 'Identifier' &&
          node.callee.property.name === 'sendBeacon'
        ) {
          context.report({ node, messageId: 'networkPrimitive', data: { name: 'navigator.sendBeacon' } });
        }
      },
      // import axios / import 'node:https' etc.
      ImportDeclaration(node) {
        const src = node.source.value;
        if (NETWORK_IMPORT_PATTERNS.some(p => p.test(src))) {
          context.report({ node, messageId: 'networkImport', data: { source: src } });
        }
      },
    };
  },
};

// ─── Rule: no-us-literals ────────────────────────────────────────────────────

const noUsLiterals = {
  meta: {
    type: 'problem',
    docs: {
      description:
        'Disallow US-centric hardcoded strings in product-facing code (jurisdiction invariant — ADR-0002). ' +
        'Terminology and date formats must come from the Config Layer (spec/CONVENTIONS.md).',
      recommended: true,
    },
    schema: [],
    messages: {
      usLiteral: '{{msg}}',
    },
  },
  create(context) {
    const filename = context.getFilename();
    // Config Layer and test fixtures are allowed to define/test these values.
    if (
      filename.includes('src/shared/config') ||
      filename.includes('tests/') ||
      filename.includes('spec/')
    ) {
      return {};
    }

    function checkLiteral(node, value) {
      if (typeof value !== 'string') return;
      for (const { pattern, msg } of US_LITERAL_PATTERNS) {
        if (pattern.test(value)) {
          context.report({ node, messageId: 'usLiteral', data: { msg } });
          break;
        }
      }
    }

    return {
      Literal(node) {
        checkLiteral(node, node.value);
      },
      TemplateLiteral(node) {
        for (const quasi of node.quasis) {
          checkLiteral(node, quasi.value.raw);
        }
      },
    };
  },
};

// ─── Rule: no-cut-feature-refs ───────────────────────────────────────────────

const noCutFeatureRefs = {
  meta: {
    type: 'problem',
    docs: {
      description:
        'Disallow references to cut features (scope invariant — ADR-0005). ' +
        'These features are permanently out of scope for Mutatis. ' +
        'See the scope register §Cut for rationale.',
      recommended: true,
    },
    schema: [],
    messages: {
      cutFeature:
        '"{{name}}" references a cut feature (ADR-0005 / the scope register §Cut). ' +
        'This feature was deliberately removed — do not reintroduce it.',
    },
  },
  create(context) {
    const filename = context.getFilename();
    // ADRs, spec, and test fixtures may reference cut features by name.
    if (
      filename.includes('docs/') ||
      filename.includes('spec/') ||
      filename.includes('tests/') ||
      filename.includes('eslint-plugins/')
    ) {
      return {};
    }

    return {
      Identifier(node) {
        if (CUT_FEATURE_IDENTIFIERS.includes(node.name)) {
          context.report({ node, messageId: 'cutFeature', data: { name: node.name } });
        }
      },
    };
  },
};

// ─── Rule: no-host-imports ───────────────────────────────────────────────────

const noHostImports = {
  meta: {
    type: 'problem',
    docs: {
      description:
        'Disallow references to the Office/Word host globals outside the host adapters ' +
        '(host-portability invariant — ADR-0022). The feature core and UI must depend only ' +
        'on the HostPort seam so they can run in both the Word add-in and the standalone web app.',
      recommended: true,
    },
    schema: [],
    messages: {
      hostRef:
        'Reference to host global "{{name}}" is forbidden here (ADR-0022 / spec/features/12). ' +
        'core/ui must be host-neutral and depend only on the HostPort seam (@shared/host); ' +
        'only host adapters (src/host-office/**) or an as-yet-unported write/navigate path may touch Office JS. ' +
        'Make this code host-neutral — do NOT add the file to HOST_ALLOWED just to silence this.',
    },
  },
  create(context) {
    const filename = context.getFilename();
    if (isHostAllowed(filename)) return {};

    function check(reportNode, rootNode) {
      if (rootNode && rootNode.type === 'Identifier' && HOST_GLOBALS.has(rootNode.name)) {
        context.report({ node: reportNode, messageId: 'hostRef', data: { name: rootNode.name } });
      }
    }

    return {
      // value position: Word.run(...), Office.context.document, ...
      MemberExpression(node) {
        check(node, node.object);
      },
      // type position: Word.RequestContext, Office.File, ...
      TSQualifiedName(node) {
        check(node, node.left);
      },
    };
  },
};

// ─── Plugin export ───────────────────────────────────────────────────────────

module.exports = {
  rules: {
    'no-network-egress': noNetworkEgress,
    'no-us-literals': noUsLiterals,
    'no-cut-feature-refs': noCutFeatureRefs,
    'no-host-imports': noHostImports,
  },
};
