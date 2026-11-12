/**
 * The one custom invariant for Mutatis Desktop.
 *
 * Rule — no-renderer-provider-access
 *
 * A provider API key, or a direct call to a provider, must never be reachable
 * from renderer or preload code. This is the single invariant here that is
 * catastrophic on ONE violation: a key that reaches the renderer is a key in a
 * process that loads remote-ish content and ships devtools, and a provider call
 * from the renderer both leaks the key and defeats the audit story.
 *
 * Deliberately ONE rule. This project does not replicate the add-in's
 * wider lint rules. Everything else here is caught by review or by manual
 * testing in the app.
 *
 * Defence in depth: this lint is the first wall. The second is the CSP's
 * `connect-src 'self'` (src/main/app-protocol.ts), which blocks a provider
 * request from the renderer at runtime even if one is somehow written.
 */

/** Packages that speak directly to a model provider. Never in renderer/preload. */
const PROVIDER_SDK_PATTERNS = [
  /^@ai-sdk\/(?!react$|provider-utils$)/, //  @ai-sdk/react is the UI half and is fine
  /^openai$/,
  /^@anthropic-ai\//,
  /^@google\/genai$/,
  /^@mistralai\//,
  /^cohere-ai$/,
  /^ollama$/,
  /^groq-sdk$/,
];

/** Provider hostnames. A literal containing one of these is a direct call. */
const PROVIDER_HOSTS = [
  'api.deepseek.com',
  'dashscope.aliyuncs.com',
  'dashscope-intl.aliyuncs.com',
  'api.openai.com',
  'api.anthropic.com',
  'generativelanguage.googleapis.com',
  'api.mistral.ai',
  'api.groq.com',
  'openrouter.ai',
];

/** Literal shapes that are almost certainly a real credential. */
const KEY_LITERAL_PATTERNS = [
  /\bsk-[A-Za-z0-9_-]{16,}/,
  /\bsk-ant-[A-Za-z0-9_-]{16,}/,
  /\bAIza[0-9A-Za-z_-]{30,}/,
];

/** A build-time env var whose NAME suggests it carries a secret. */
const SECRET_ENV_NAME = /(KEY|SECRET|TOKEN|PASSWORD|CREDENTIAL)/i;

/** Renderer and preload are the untrusted side of the bridge. */
function isRendererOrPreload(filename) {
  if (!filename) return false;
  const f = filename.replace(/\\/g, '/');
  return f.includes('/src/renderer/') || f.includes('/src/preload/');
}

const noRendererProviderAccess = {
  meta: {
    type: 'problem',
    docs: {
      description:
        'Provider SDKs, provider hostnames and API keys must never be reachable from renderer or preload code.',
    },
    schema: [],
    messages: {
      providerSdk:
        "Provider SDK '{{name}}' imported in renderer/preload. ALL provider HTTP happens in the Electron main process — the renderer talks to main over the IPC bridge. Move this import to src/main/.",
      providerHost:
        "Provider hostname '{{host}}' appears in renderer/preload code. The renderer must never call a provider directly: DeepSeek and Qwen have no browser CORS story, the key would have to be in the renderer to use it, and CSP connect-src 'self' blocks it at runtime anyway.",
      keyLiteral:
        'What looks like a live API key is hard-coded here. Keys live in the Electron main process in safeStorage, and never appear in source.',
      secretEnv:
        "'{{name}}' is a build-time env var whose name suggests a secret. Anything in import.meta.env is INLINED INTO THE RENDERER BUNDLE in plain text. Read secrets in main from safeStorage instead.",
    },
  },

  create(context) {
    const filename = context.filename ?? context.getFilename();
    if (!isRendererOrPreload(filename)) return {};

    function checkLiteral(node, value) {
      if (typeof value !== 'string') return;

      for (const host of PROVIDER_HOSTS) {
        if (value.includes(host)) {
          context.report({ node, messageId: 'providerHost', data: { host } });
          return;
        }
      }
      for (const pattern of KEY_LITERAL_PATTERNS) {
        if (pattern.test(value)) {
          context.report({ node, messageId: 'keyLiteral' });
          return;
        }
      }
    }

    function checkModuleSource(node, name) {
      if (typeof name !== 'string') return;
      for (const pattern of PROVIDER_SDK_PATTERNS) {
        if (pattern.test(name)) {
          context.report({ node, messageId: 'providerSdk', data: { name } });
          return;
        }
      }
    }

    return {
      ImportDeclaration(node) {
        checkModuleSource(node, node.source.value);
      },

      // `await import('@ai-sdk/deepseek')` and `require('openai')`
      ImportExpression(node) {
        if (node.source.type === 'Literal') checkModuleSource(node, node.source.value);
      },
      CallExpression(node) {
        if (
          node.callee.type === 'Identifier' &&
          node.callee.name === 'require' &&
          node.arguments[0]?.type === 'Literal'
        ) {
          checkModuleSource(node, node.arguments[0].value);
        }
      },

      Literal(node) {
        checkLiteral(node, node.value);
      },
      TemplateElement(node) {
        checkLiteral(node, node.value.cooked);
      },

      // import.meta.env.VITE_DEEPSEEK_API_KEY — inlined into the bundle in clear.
      MemberExpression(node) {
        if (node.property.type !== 'Identifier') return;
        const name = node.property.name;
        if (!SECRET_ENV_NAME.test(name)) return;

        const obj = node.object;
        const isEnvRead =
          (obj.type === 'MemberExpression' &&
            obj.property.type === 'Identifier' &&
            obj.property.name === 'env') ||
          (obj.type === 'Identifier' && obj.name === 'env');

        if (isEnvRead) {
          context.report({ node, messageId: 'secretEnv', data: { name } });
        }
      },
    };
  },
};

module.exports = {
  rules: {
    'no-renderer-provider-access': noRendererProviderAccess,
  },
};
