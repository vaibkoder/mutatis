import { describe, it, expect } from 'vitest';
import { RuleTester } from 'eslint';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const invariants = require('../eslint-plugins/invariants.cjs');

const rule = invariants.rules['no-renderer-provider-access'];

/**
 * The one catastrophic invariant: no provider key, secret, or direct provider
 * call reachable from renderer or preload code.
 *
 * These tests exist because the rule is only worth having if it actually fires.
 * A lint rule that silently matches nothing is worse than no rule — it buys
 * false confidence in exactly the place we can least afford it.
 */
const ruleTester = new RuleTester({
  languageOptions: {
    ecmaVersion: 2022,
    sourceType: 'module',
  },
});

const RENDERER = '/repo/src/renderer/src/thing.ts';
const PRELOAD = '/repo/src/preload/index.ts';
const MAIN = '/repo/src/main/providers.ts';

describe('no-renderer-provider-access', () => {
  it('flags violations in renderer and preload, and allows them in main', () => {
    ruleTester.run('no-renderer-provider-access', rule, {
      valid: [
        // Main is where provider code BELONGS. The rule must not fire here.
        { code: `import { createDeepSeek } from '@ai-sdk/deepseek';`, filename: MAIN },
        { code: `const u = 'https://api.deepseek.com/v1';`, filename: MAIN },
        { code: `const k = process.env.DEEPSEEK_API_KEY;`, filename: MAIN },

        // @ai-sdk/react is the UI half — no provider HTTP, no credentials.
        { code: `import { useChat } from '@ai-sdk/react';`, filename: RENDERER },

        // Ordinary renderer code must not trip it.
        { code: `const title = 'Compare documents';`, filename: RENDERER },
        { code: `const n = import.meta.env.VITE_BUILD_NUMBER;`, filename: RENDERER },
        { code: `fetch('./index.html');`, filename: RENDERER },
      ],

      invalid: [
        {
          code: `import { createDeepSeek } from '@ai-sdk/deepseek';`,
          filename: RENDERER,
          errors: [{ messageId: 'providerSdk' }],
        },
        {
          code: `import Anthropic from '@anthropic-ai/sdk';`,
          filename: PRELOAD,
          errors: [{ messageId: 'providerSdk' }],
        },
        {
          code: `const m = await import('openai');`,
          filename: RENDERER,
          errors: [{ messageId: 'providerSdk' }],
        },
        {
          code: `fetch('https://api.deepseek.com/v1/chat/completions');`,
          filename: RENDERER,
          errors: [{ messageId: 'providerHost' }],
        },
        {
          code: 'const u = `https://dashscope-intl.aliyuncs.com/compatible-mode/v1`;',
          filename: RENDERER,
          errors: [{ messageId: 'providerHost' }],
        },
        {
          code: `const k = 'sk-abcdefghijklmnopqrstuvwxyz123456';`,
          filename: RENDERER,
          errors: [{ messageId: 'keyLiteral' }],
        },
        {
          code: `const k = import.meta.env.VITE_DEEPSEEK_API_KEY;`,
          filename: RENDERER,
          errors: [{ messageId: 'secretEnv' }],
        },
        {
          code: `const s = import.meta.env.VITE_SESSION_TOKEN;`,
          filename: PRELOAD,
          errors: [{ messageId: 'secretEnv' }],
        },
      ],
    });
  });

  it('exposes exactly one rule — this project keeps its lint surface deliberately small', () => {
    expect(Object.keys(invariants.rules)).toEqual(['no-renderer-provider-access']);
  });
});
