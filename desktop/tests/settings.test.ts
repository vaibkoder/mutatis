import { describe, it, expect } from 'vitest';
import { DEFAULT_SETTINGS, parseSettings } from '../src/renderer/src/settings/parse.js';

/**
 * Stored preferences are read at startup, before anything is on screen. A
 * parse that throws there is a window that never opens, so every malformed
 * shape below must come back as a usable Settings object.
 */
describe('parseSettings', () => {
  it('returns defaults when nothing is stored', () => {
    expect(parseSettings(null)).toEqual(DEFAULT_SETTINGS);
    expect(parseSettings('')).toEqual(DEFAULT_SETTINGS);
  });

  it('returns defaults rather than throwing on corrupt JSON', () => {
    expect(parseSettings('{not json')).toEqual(DEFAULT_SETTINGS);
    expect(parseSettings('null')).toEqual(DEFAULT_SETTINGS);
    expect(parseSettings('[1,2,3]')).toEqual(DEFAULT_SETTINGS);
    expect(parseSettings('"a string"')).toEqual(DEFAULT_SETTINGS);
  });

  it('round-trips a well-formed value', () => {
    const stored = JSON.stringify({
      jurisdiction: 'SG',
      model: { provider: 'alibaba', model: 'qwen3.7-max' },
    });
    expect(parseSettings(stored)).toEqual({
      jurisdiction: 'SG',
      model: { provider: 'alibaba', model: 'qwen3.7-max' },
    });
  });

  it('falls back per field, keeping the half that is valid', () => {
    const stored = JSON.stringify({ jurisdiction: 'SG', model: { provider: 'openai' } });
    expect(parseSettings(stored)).toEqual({
      jurisdiction: 'SG',
      model: DEFAULT_SETTINGS.model,
    });
  });

  it('treats an unknown jurisdiction as UK rather than refusing to start', () => {
    expect(parseSettings(JSON.stringify({ jurisdiction: 'US' })).jurisdiction).toBe('UK');
    expect(parseSettings(JSON.stringify({ jurisdiction: 42 })).jurisdiction).toBe('UK');
  });

  it('rejects a model entry with an empty id', () => {
    const stored = JSON.stringify({ model: { provider: 'deepseek', model: '' } });
    expect(parseSettings(stored).model).toEqual(DEFAULT_SETTINGS.model);
  });

  it('does not carry unknown keys through', () => {
    const stored = JSON.stringify({ jurisdiction: 'UK', apiKey: 'sk-should-never-be-here' });
    expect(Object.keys(parseSettings(stored)).sort()).toEqual(['jurisdiction', 'model']);
  });
});
