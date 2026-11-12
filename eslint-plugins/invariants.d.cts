/**
 * Type surface for the flat-config plugin, so `tests/unit/eslint/*` can import it
 * under `tsconfig.test.json` without an implicit `any`. The plugin itself stays
 * CommonJS because ESLint loads it directly.
 */
import type { Rule } from 'eslint';
declare const plugin: { rules: Record<string, Rule.RuleModule> };
export = plugin;
