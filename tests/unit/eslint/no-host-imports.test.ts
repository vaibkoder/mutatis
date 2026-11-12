import { RuleTester } from 'eslint';
import tsParser from '@typescript-eslint/parser';

import invariants from '../../../eslint-plugins/invariants.cjs';

const rule = (invariants as { rules: Record<string, unknown> }).rules['no-host-imports'];

const ruleTester = new RuleTester({
  languageOptions: {
    parser: tsParser,
    parserOptions: { ecmaVersion: 2020, sourceType: 'module' },
  },
});

ruleTester.run('no-host-imports', rule as never, {
  valid: [
    // Host-free code in the protected port — no Office/Word reference at all.
    { code: 'export const x: number = 1;', filename: 'src/shared/host/types.ts' },
    // Office references are allowed inside a host adapter (allow-listed path).
    {
      code: 'function f(): void { void Word.run; const c: Word.RequestContext = null as never; void c; }',
      filename: 'src/host-office/office-host.ts',
    },
    // Allowed inside an as-yet-unported feature write path.
    {
      code: 'function f(): void { void Office.context; }',
      filename: 'src/features/numbering/apply.ts',
    },
    // A non-host member access is never flagged.
    { code: 'const w = { Word: 1 }; export const v = w.Word;', filename: 'src/features/factual-consistency/x.ts' },
  ],
  invalid: [
    // Value position in a protected (non-allow-listed) feature.
    {
      code: 'async function f(): Promise<void> { await Word.run(async () => {}); }',
      filename: 'src/features/factual-consistency/dates.ts',
      errors: [{ messageId: 'hostRef' }],
    },
    // Type position in the protected port.
    {
      code: 'export function f(c: Word.RequestContext): void { void c; }',
      filename: 'src/shared/host/types.ts',
      errors: [{ messageId: 'hostRef' }],
    },
    // Office namespace in the protected port.
    {
      code: 'export const u: unknown = Office.context.document;',
      filename: 'src/shared/host/active.ts',
      errors: [{ messageId: 'hostRef' }],
    },
  ],
});
