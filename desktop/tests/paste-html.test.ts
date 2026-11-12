// @vitest-environment happy-dom
import { describe, expect, it } from 'vitest';
import { htmlToBlocks } from '../src/renderer/src/paste/html-to-blocks.js';
import { blocksText, hasFormatting, type Block, type Para } from '../src/renderer/src/paste/model.js';

const paras = (blocks: Block[]) => blocks.filter((b): b is Para => b.kind === 'p');
const run = (p: Para | undefined, text: string) => p?.runs.find((r) => r.text === text);

// Trimmed from what Word puts on the clipboard: conditional comments, MsoNormal
// classes, a numbered list written as mso-list paragraphs, empty o:p elements.
const WORD = `<html xmlns:o="urn:schemas-microsoft-com:office:office"><head><style><!-- p.MsoNormal {margin:0} --></style></head><body>
<!--StartFragment-->
<h1 style='margin-top:12.0pt'><span style='font-weight:normal'>Definitions</span></h1>
<p class=MsoNormal style='text-align:justify'>The <b style='mso-bidi-font-weight:normal'>Seller</b> shall <i>not</i> sell <u>any</u> Shares<o:p></o:p></p>
<p class=MsoNormal><o:p>&nbsp;</o:p></p>
<p class=MsoListParagraph style='margin-left:.5in;text-indent:-.25in;mso-list:l0 level1 lfo1'><![if !supportLists]><span style='mso-list:Ignore'>1.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp; </span></span><![endif]>First term</p>
<p class=MsoListParagraph style='margin-left:.5in;text-indent:-.25in;mso-list:l0 level1 lfo1'><span style='mso-list:Ignore'><span style='font-family:Symbol'>·<span>&nbsp;&nbsp;</span></span></span>Bullet</p>
<!--EndFragment-->
</body></html>`;

// Google Docs wraps everything in a <b style="font-weight:normal"> and states
// bold/italic as inline styles on spans.
const GDOCS = `<meta charset='utf-8'><b style="font-weight:normal;" id="docs-internal-guid-1"><p dir="ltr" style="text-align:center"><span style="font-weight:700;">Heads</span><span style="font-weight:400;"> and </span><span style="font-style:italic;text-decoration:underline;">tails</span><span style="text-decoration:line-through;"> gone</span></p></b>`;

describe('Word clipboard HTML', () => {
  const blocks = htmlToBlocks(WORD);
  const ps = paras(blocks);

  it('reads the heading, and does not treat font-weight:normal as bold', () => {
    expect(ps[0]?.heading).toBe(1);
    expect(ps[0]?.runs.every((r) => !r.bold)).toBe(true);
    expect(blocksText(blocks)).toContain('Definitions');
  });

  it('keeps bold, italic, underline and justification', () => {
    expect(run(ps[1], 'Seller')?.bold).toBe(true);
    expect(run(ps[1], 'not')?.italic).toBe(true);
    expect(run(ps[1], 'any')?.underline).toBe(true);
    expect(ps[1]?.align).toBe('both');
  });

  it('keeps an empty paragraph as a blank one', () => {
    expect(ps[2]?.runs).toEqual([]);
  });

  it('keeps list markers as text, shows a Symbol bullet as a bullet, and indents', () => {
    expect(blocksText([ps[3]!])).toBe('1. First term');
    expect(blocksText([ps[4]!])).toBe('• Bullet');
    expect(ps[3]?.list).toBe(true);
    expect(ps[3]?.indent).toBe(720);
  });

  it('counts as rich', () => {
    expect(hasFormatting(blocks)).toBe(true);
  });
});

describe('Google Docs clipboard HTML', () => {
  const [p] = paras(htmlToBlocks(GDOCS));
  it('reads style-based bold/italic/underline/strike through the normal-weight wrapper', () => {
    expect(run(p, 'Heads')?.bold).toBe(true);
    expect(run(p, ' and ')?.bold).toBeFalsy();
    expect(run(p, 'tails')).toMatchObject({ italic: true, underline: true });
    expect(run(p, ' gone')?.strike).toBe(true);
    expect(p?.align).toBe('center');
  });
});

describe('web and editor HTML', () => {
  it('numbers ordered lists, bullets unordered ones, nests them', () => {
    const t = blocksText(htmlToBlocks('<ol start="3"><li>a</li><li>b<ul><li>c</li></ul></li></ol>'));
    expect(t).toBe('3. a\n4. b\n• c');
  });

  it('reads a table with a bold header row', () => {
    const [t] = htmlToBlocks('<table><thead><tr><th>A</th><th>B</th></tr></thead><tbody><tr><td>1</td><td><b>2</b></td></tr></tbody></table>');
    expect(t?.kind).toBe('table');
    if (t?.kind !== 'table') return;
    expect(t.rows).toHaveLength(2);
    expect(t.rows[0]?.[0]?.[0]?.runs[0]?.bold).toBe(true);
    expect(t.rows[1]?.[1]?.[0]?.runs[0]?.bold).toBe(true);
  });

  it('turns <br> into a line break but drops a trailing one', () => {
    const [p] = paras(htmlToBlocks('<p>one<br>two<br></p>'));
    expect(p?.runs.map((r) => (r.br ? '\n' : r.text)).join('')).toBe('one\ntwo');
  });

  it('collapses source whitespace but keeps <pre> lines', () => {
    expect(blocksText(htmlToBlocks('<p>a \n   b</p>'))).toBe('a b');
    expect(blocksText(htmlToBlocks('<pre>x\n  y</pre>'))).toBe('x\n  y');
  });

  it('is not rich when there is no formatting (an editor copy is just text)', () => {
    const blocks = htmlToBlocks('<meta charset="utf-8"><div style="font-family:Menlo;color:#333"><div><span style="color:red">hello</span></div><div><span>world</span></div></div>');
    expect(blocksText(blocks)).toBe('hello\nworld');
    expect(hasFormatting(blocks)).toBe(false);
  });
});

describe('hostile input', () => {
  it('yields text only: scripts, handlers, images and links contribute nothing executable', () => {
    const blocks = htmlToBlocks(
      '<p onclick="x()">keep <a href="javascript:alert(1)">this</a><script>alert(1)</script><img src="https://evil.example/p.png" onerror="x()"><iframe src="https://evil.example"></iframe></p><style>p{color:red}</style>',
    );
    expect(blocksText(blocks)).toBe('keep this');
  });
});
