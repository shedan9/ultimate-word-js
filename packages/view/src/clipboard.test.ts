// @vitest-environment jsdom
import type { ResolvedRunProps } from '@uw/model';
import { describe, expect, it } from 'vitest';
import { fragmentToHtml, htmlToParagraphs, isTrustedHtml } from './clipboard.ts';

const base: ResolvedRunProps = {
  fonts: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '仿宋', cs: '', hint: 'eastAsia' },
  bold: false,
  boldCs: false,
  italic: false,
  italicCs: false,
  caps: false,
  smallCaps: false,
  strike: false,
  doubleStrike: false,
  hidden: false,
  size: 320,
  sizeCs: 320,
  underline: 'none',
  color: 'auto',
  vertAlign: 'baseline',
  charSpacing: 0,
  scale: 100,
  position: 0,
  kerning: 0,
  snapToGrid: true,
  langEastAsia: 'zh-CN',
};
const parse = (html: string) => htmlToParagraphs(html, new DOMParser());

describe('剪贴板 HTML', () => {
  it('复制出的 HTML 读回来格式与段落不变：字体三个桶、字号、开关、颜色、对齐、连续空格与空段', () => {
    const html = fragmentToHtml({
      paragraphs: [
        {
          justification: 'center',
          runs: [
            { text: '版    本', props: { ...base, bold: true, color: 'FF0000' } },
            {
              text: 'x<&>"',
              props: { ...base, italic: true, underline: 'double', vertAlign: 'superscript' },
            },
          ],
        },
        { justification: 'both', runs: [] },
      ],
    });
    expect(isTrustedHtml(html)).toBe(true);
    const fonts = { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '仿宋' };
    const common = { strike: false, size: 320, fonts };
    expect(parse(html)).toEqual([
      {
        justification: 'center',
        runs: [
          {
            text: '版    本',
            patch: {
              ...common,
              bold: true,
              italic: false,
              underline: 'none',
              vertAlign: 'baseline',
              color: 'FF0000',
            },
          },
          {
            text: 'x<&>"',
            patch: { ...common, bold: false, italic: true, underline: 'double', vertAlign: 'superscript' },
          },
        ],
      },
      { justification: 'both', runs: [] },
    ]);
  });

  it('Word 的 HTML：mso-* 分桶、源码折行不在汉字间插空格、<o:p> 空段、列表编号不进正文', () => {
    const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office"><body><!--StartFragment-->
<p class=MsoNormal align=center style='text-align:center'><span lang=EN-US style='font-size:16.0pt;
font-family:"Times New Roman",serif;mso-fareast-font-family:仿宋_GB2312'>2026</span><span
style='font-size:16.0pt;font-family:仿宋_GB2312;mso-ascii-font-family:"Times New Roman";color:red'>年关于
通知</span></p>
<p class=MsoNormal><span lang=EN-US><o:p>&nbsp;</o:p></span></p>
<p class=MsoListParagraph><![if !supportLists]><span style='mso-list:Ignore'>1.<span
style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp; </span></span><![endif]><b>要点</b></p>
<!--EndFragment--></body></html>`;
    const off = { italic: false, underline: 'none', strike: false, vertAlign: 'baseline' } as const;
    expect(parse(html)).toEqual([
      {
        justification: 'center',
        runs: [
          {
            text: '2026',
            patch: {
              ...off,
              bold: false,
              size: 320,
              fonts: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '仿宋_GB2312' },
            },
          },
          {
            text: '年关于通知',
            patch: {
              ...off,
              bold: false,
              size: 320,
              color: 'FF0000',
              fonts: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '仿宋_GB2312' },
            },
          },
        ],
      },
      { runs: [] },
      { runs: [{ text: '要点', patch: { ...off, bold: true } }] },
    ]);
  });

  it('网页 HTML 只带开关且只加不减，字体字号颜色对齐一律不带；块与 <br> 拆段、空白折叠', () => {
    const html = `<div style="text-align:center"><b>粗</b>  普通
      <span style="font-family:Arial;font-size:14px;color:#00f">字</span> </div><div><br></div>
      <ul><li><i>一</i></li><li><s>二</s><sub>2</sub></li></ul><p>甲<br>乙</p><img src="x.png"><table><tr><td>格</td><td></td></tr></table>`;
    expect(isTrustedHtml(html)).toBe(false);
    expect(parse(html)).toEqual([
      {
        runs: [
          { text: '粗', patch: { bold: true } },
          { text: ' 普通 字', patch: {} },
        ],
      },
      { runs: [] },
      { runs: [{ text: '一', patch: { italic: true } }] },
      {
        runs: [
          { text: '二', patch: { strike: true } },
          { text: '2', patch: { vertAlign: 'subscript' } },
        ],
      },
      { runs: [{ text: '甲', patch: {} }] },
      { runs: [{ text: '乙', patch: {} }] },
      { runs: [{ text: '格', patch: {} }] },
    ]);
  });

  it('pre 保留空白并按换行拆段，只有图片的 HTML 读不出段落', () => {
    expect(parse('<pre>a  b\n\n c </pre>')).toEqual([
      { runs: [{ text: 'a  b', patch: {} }] },
      { runs: [] },
      { runs: [{ text: ' c ', patch: {} }] },
    ]);
    expect(parse('<img src="a.png"><style>p{}</style>')).toEqual([]);
  });
});
