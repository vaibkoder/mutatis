import { createAlibaba } from '@ai-sdk/alibaba';
import { createDeepSeek } from '@ai-sdk/deepseek';
import type { LanguageModel } from 'ai';
import { getKey } from '../secrets.js';
import type { ModelChoice, ProviderId, ProviderInfo } from '../../common/ipc.js';

/**
 * Provider registry. MAIN PROCESS ONLY.
 *
 * Every provider HTTP call originates here. The renderer cannot reach a
 * provider: it has no key, the CSP forbids the connection, and a lint rule
 * fails the build if anyone writes the import. Main is Node, so browser CORS —
 * which DeepSeek and Qwen have no story for — never applies.
 */

export type { ModelChoice, ProviderInfo };

export const PROVIDERS: readonly ProviderInfo[] = [
  {
    id: 'deepseek',
    label: 'DeepSeek',
    models: [
      { id: 'deepseek-v4-flash', label: 'V4 Flash — fast, cheapest' },
      { id: 'deepseek-v4-pro', label: 'V4 Pro — strongest reasoning' },
    ],
    keyUrl: 'https://platform.deepseek.com/api_keys',
  },
  {
    id: 'alibaba',
    label: 'Qwen',
    models: [
      { id: 'qwen3.7-max', label: '3.7 Max — strongest' },
      { id: 'qwen-plus', label: 'Plus — balanced' },
      { id: 'qwen-flash', label: 'Flash — fast, cheapest' },
    ],
    keyUrl: 'https://modelstudio.console.alibabacloud.com/',
  },
];

export const DEFAULT_MODEL: ModelChoice = {
  provider: 'deepseek',
  model: 'deepseek-v4-flash',
};

export class MissingKeyError extends Error {
  constructor(readonly provider: ProviderId) {
    super(`No API key is configured for ${provider}. Add one in Settings.`);
    this.name = 'MissingKeyError';
  }
}

export function resolveModel(choice: ModelChoice): LanguageModel {
  const key = getKey(choice.provider);
  if (!key) throw new MissingKeyError(choice.provider);

  switch (choice.provider) {
    case 'deepseek':
      return createDeepSeek({ apiKey: key })(choice.model);
    case 'alibaba':
      // The SDK's default base URL is dashscope-INTL, not the mainland-China
      // endpoint. That is the right default here and must not be "corrected":
      // the two are separate services with separate consoles, and a key issued
      // for one returns 401 on the other.
      return createAlibaba({ apiKey: key })(choice.model);
  }
}
