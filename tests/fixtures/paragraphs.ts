import type { ParaNode } from '@shared/doc-model/types.js';

/** Build a ParaNode fixture; getRange throws (navigation not exercised in unit tests). */
export function para(
  index: number,
  text: string,
  opts: { listString?: string; listLevel?: number; styleName?: string; container?: ParaNode['container'] } = {}
): ParaNode {
  const node: ParaNode = {
    id: `body-para-${index}`,
    index,
    text,
    styleName: opts.styleName ?? 'Normal',
    container: opts.container ?? 'body',
    getRange: () => { throw new Error('getRange not used in this test'); },
  };
  if (opts.listString !== undefined) node.listString = opts.listString;
  if (opts.listLevel !== undefined) node.listLevel = opts.listLevel;
  return node;
}
