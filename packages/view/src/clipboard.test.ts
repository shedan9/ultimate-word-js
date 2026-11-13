// @vitest-environment jsdom
import { createDiagnosticSink } from '@uw/core';
import type { CascadeContext, ResolvedRunProps, RunProps, Style } from '@uw/model';
import {
  createTextEditor,
  DEFAULT_SECTION_PROPS,
  DEFAULT_SETTINGS,
  EMPTY_NUMBERING,
  EMPTY_THEME,
  parseStyles,
  rangeOfNode,
  resolveRunProps,
  walkParagraphs,
} from '@uw/model';
import { highlightFill, shadingFill } from '@uw/render-dom';
import { describe, expect, it } from 'vitest';
import { fragmentToHtml, htmlToParagraphs, isTrustedHtml } from './clipboard.ts';
import { createEditingController } from './editing.ts';

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
  highlight: 'none',
  shading: undefined,
  emphasis: 'none',
  charSpacing: 0,
  scale: 100,
  position: 0,
  kerning: 0,
  snapToGrid: true,
  langEastAsia: 'zh-CN',
};
const parse = (html: string) => htmlToParagraphs(html, new DOMParser());
const noShading = { pattern: 'nil', color: 'auto', fill: 'auto' } as const;

/** 四种来源都会在删掉直接格式后重新级联；不能只在无样式的空文档里验证粘贴。 */
function decoratedTarget(layer: 'defaults' | 'paragraph' | 'character' | 'direct') {
  const decoration: RunProps = {
    highlight: 'yellow',
    shading: { pattern: 'clear', color: 'auto', fill: 'D9D9D9' },
    emphasis: 'dot',
  };
  const style: Style = {
    id: 'Decorated',
    type: layer === 'paragraph' ? 'paragraph' : 'character',
    name: 'Decorated',
    basedOn: undefined,
    next: undefined,
    isDefault: false,
    paraProps: {},
    runProps: decoration,
    tableProps: {},
    rowProps: {},
    cellProps: {},
    conditional: new Map(),
  };
  const ctx: CascadeContext = {
    styles: {
      ...parseStyles(undefined, createDiagnosticSink()),
      defaults: { paraProps: {}, runProps: layer === 'defaults' ? decoration : {} },
      chainOf: (id) => (id === 'Decorated' ? [style] : []),
    },
    theme: EMPTY_THEME,
    settings: DEFAULT_SETTINGS,
    numbering: EMPTY_NUMBERING,
  };
  const editor = createTextEditor({
    sections: [
      {
        id: 's',
        props: DEFAULT_SECTION_PROPS,
        blocks: [
          {
            kind: 'paragraph',
            id: 'p',
            props: layer === 'paragraph' ? { styleId: 'Decorated' } : {},
            runs: [
              {
                kind: 'run',
                id: 'r',
                props:
                  layer === 'direct' ? decoration : layer === 'character' ? { styleId: 'Decorated' } : {},
                content: [{ kind: 'text', text: '前后' }],
              },
            ],
          },
        ],
      },
    ],
  });
  const state = createEditingController(editor);
  const propsOf = (text: string) => {
    const paragraph = [...walkParagraphs(editor.body)][0];
    const run = paragraph?.runs.find((r) =>
      r.content.some((c) => c.kind === 'text' && c.text.includes(text)),
    );
    if (!paragraph || !run) throw new Error(`缺少测试文字 ${text}`);
    return resolveRunProps(ctx, paragraph.props, run.props);
  };
  return { editor, state, propsOf };
}

describe('剪贴板 HTML', () => {
  it.each(['defaults', 'paragraph', 'character', 'direct'] as const)(
    '可信粘贴显式盖掉目标 %s 的高亮、底纹和着重号，撤销重做与周围文字不受影响',
    (layer) => {
      const { editor, state, propsOf } = decoratedTarget(layer);
      const at = { nodeId: 'r', contentIndex: 0, offset: 1 };
      state.select({ start: at, end: at });
      state.insertParagraphs(
        parse(
          `<p class=MsoNormal>普通<span style="background:lime;mso-highlight:green">高亮</span><span style="background:#123456">底纹</span><span style="text-emphasis:filled dot">着重</span></p>`,
        ),
      );
      const plain = propsOf('普通');
      expect(plain.highlight).toBe('none');
      expect(shadingFill(plain.shading)).toBeUndefined();
      expect(plain.emphasis).toBe('none');
      expect(highlightFill(propsOf('高亮').highlight)).toBe('#00ff00');
      expect(shadingFill(propsOf('高亮').shading)).toBeUndefined();
      expect(propsOf('高亮').emphasis).toBe('none');
      expect(propsOf('底纹').highlight).toBe('none');
      expect(shadingFill(propsOf('底纹').shading)).toBe('#123456');
      expect(propsOf('底纹').emphasis).toBe('none');
      expect(propsOf('着重').emphasis).toBe('dot');
      expect(propsOf('着重').highlight).toBe('none');
      expect(shadingFill(propsOf('着重').shading)).toBeUndefined();
      expect(propsOf('前').highlight).toBe('yellow');
      expect(propsOf('后').emphasis).toBe('dot');
      const pasted = editor.body;
      editor.undo();
      expect([...walkParagraphs(editor.body)][0]?.runs[0]?.content).toEqual([{ kind: 'text', text: '前后' }]);
      expect(propsOf('前后').highlight).toBe('yellow');
      editor.redo();
      expect(editor.body).toEqual(pasted);
      expect(shadingFill(propsOf('普通').shading)).toBeUndefined();
    },
  );

  it('普通网页粘贴仍继承目标装饰，自己的普通文字复制后明确关闭装饰', () => {
    const { editor, state, propsOf } = decoratedTarget('paragraph');
    const at = { nodeId: 'r', contentIndex: 0, offset: 1 };
    state.select({ start: at, end: at });
    state.insertParagraphs(parse('<p>网页</p>'));
    expect(propsOf('网页').highlight).toBe('yellow');
    expect(shadingFill(propsOf('网页').shading)).toBe('#d9d9d9');
    expect(propsOf('网页').emphasis).toBe('dot');
    const paragraph = [...walkParagraphs(editor.body)][0];
    const range = paragraph && rangeOfNode(paragraph);
    if (!range) throw new Error('缺少测试段落');
    state.select(range);
    state.insertParagraphs(
      parse(
        fragmentToHtml({
          paragraphs: [{ justification: 'left', runs: [{ text: '自有普通', props: base }] }],
        }),
      ),
    );
    expect(propsOf('自有普通').highlight).toBe('none');
    expect(shadingFill(propsOf('自有普通').shading)).toBeUndefined();
    expect(propsOf('自有普通').emphasis).toBe('none');
  });

  it.each(['background:transparent', 'background:none', 'mso-highlight:none'])(
    '源内部显式关闭 %s 与着重号，不会继承外层或重新露出目标样式',
    (background) => {
      const { state, propsOf } = decoratedTarget('paragraph');
      const at = { nodeId: 'r', contentIndex: 0, offset: 1 };
      state.select({ start: at, end: at });
      state.insertParagraphs(
        parse(
          `<style>.Off {${background};text-emphasis:none;}</style><p class=MsoNormal><span style="background:lime;mso-highlight:green;text-emphasis:filled dot">开启<span class=Off>关闭</span></span></p>`,
        ),
      );
      expect(propsOf('开启').highlight).toBe('green');
      expect(propsOf('开启').emphasis).toBe('dot');
      expect(propsOf('关闭').highlight).toBe('none');
      expect(shadingFill(propsOf('关闭').shading)).toBeUndefined();
      expect(propsOf('关闭').emphasis).toBe('none');
    },
  );

  it('mso-highlight:none 只关闭高亮，源里的独立底纹仍可继承、覆盖或显式清除', () => {
    const runs =
      parse(
        `<p class=MsoNormal><span style="background:#123456"><span style="mso-highlight:none">继承</span><span style="mso-highlight:none;background:blue">覆盖</span><span style="mso-highlight:none;background:transparent">清除</span></span></p>`,
      )[0]?.runs ?? [];
    expect(runs.map((r) => [r.text, r.patch.highlight, r.patch.shading])).toEqual([
      ['继承', 'none', { pattern: 'clear', color: 'auto', fill: '123456' }],
      ['覆盖', 'none', { pattern: 'clear', color: 'auto', fill: '0000FF' }],
      ['清除', 'none', noShading],
    ]);
  });

  it('Word 制表位按数量还原，丢弃显示用空格，类样式和段首段末同样生效', () => {
    const paragraphs = parse(`<style>.Tabs {mso-tab-count:2;font-weight:bold;}</style>
      <p class=MsoNormal><span style="mso-tab-count:1">&nbsp;   </span>甲<span class=Tabs><b>&nbsp; &nbsp;</b></span>乙<span style="mso-tab-count:1"></span></p>`);
    expect(paragraphs[0]?.runs.map((r) => [r.text, r.patch.bold])).toEqual([
      ['\t甲', false],
      ['\t\t', true],
      ['乙\t', false],
    ]);
  });

  it.each(['0', '-1', '1.5', 'NaN', 'Infinity', '99999999999999999999', '10001'])(
    '无效或过大的 Word 制表位数量 %s 保留可见文字且不抛错',
    (count) => {
      expect(
        parse(`<p class=MsoNormal>甲<span style="mso-tab-count:${count}">  </span>乙</p>`)[0]
          ?.runs.map((r) => r.text)
          .join(''),
      ).toBe('甲 乙');
      expect(
        parse(`<p class=MsoNormal>甲<span style="mso-tab-count:${count}"> 乙 </span>丙</p>`)[0]
          ?.runs.map((r) => r.text)
          .join(''),
      ).toBe('甲 乙 丙');
    },
  );

  it('制表位标记不吞正文，不应用块容器，隐藏和编号占位仍跳过', () => {
    const paragraphs = parse(
      `<p class=MsoNormal style="mso-tab-count:2">甲<span style="mso-tab-count:1">正文</span><span style="mso-tab-count:2;display:none"> </span><span style="mso-tab-count:2;mso-list:Ignore"> </span>乙</p>`,
    );
    expect(paragraphs[0]?.runs.map((r) => r.text).join('')).toBe('甲正文乙');
  });

  it.each(['\t', '\t\t', '\t正文\t', '甲\t\n\t乙'])('自己复制的制表位片段 %j 往返保留数量与段落', (text) => {
    const html = fragmentToHtml({ paragraphs: [{ justification: 'left', runs: [{ text, props: base }] }] });
    expect(html).toContain('mso-tab-count:1');
    expect(parse(html)).toHaveLength(1);
    expect(
      parse(html)[0]
        ?.runs.map((r) => r.text)
        .join(''),
    ).toBe(text);
  });

  it('Word 显式连续空格保留首尾和纯空格段，嵌套格式与引号值同样生效', () => {
    const paragraphs = parse(`<style>.Spaces {mso-spacerun:'yes';}</style>
      <p class=MsoNormal><span class=Spaces>  <b>甲  乙</b>  </span> \n</p>
      <p class=MsoNormal><span style='mso-spacerun:"yes"'> &nbsp; </span></p>`);
    expect(paragraphs.map((p) => p.runs.map((r) => r.text).join(''))).toEqual(['  甲  乙  ', '   ']);
    expect(paragraphs[0]?.runs.find((r) => r.text.includes('甲'))?.patch.bold).toBe(true);
  });

  it('显式空格后的源码空白仅删源码部分，段末普通 br 仍去占位，内部换行保留', () => {
    const paragraphs =
      parse(`<p class=MsoNormal>甲<span style="mso-spacerun:yes">  </span><b>&nbsp;</b><br><i>&nbsp;</i> </p>
      <p class=MsoNormal>乙<span style="mso-spacerun:yes">  </span><br>丙</p>
      <p class=MsoNormal>丁<span style="mso-spacerun:yes">  </span><br style="mso-special-character:line-break"></p>`);
    expect(paragraphs.map((p) => p.runs.map((r) => r.text).join(''))).toEqual(['甲  ', '乙  \n丙', '丁  \n']);
  });

  it('带制表位标记的非文字子元素仍按原有换行和块边界读取', () => {
    const paragraphs = parse(
      '<p class=MsoNormal>甲<span style="mso-tab-count:1"><br style="mso-special-character:line-break"></span>乙</p>',
    );
    expect(paragraphs[0]?.runs.map((r) => r.text).join('')).toBe('甲\n乙');
  });

  it('mso-spacerun:no 恢复空白折叠，未标记普通网页空白策略不变', () => {
    expect(
      parse(
        '<p class=MsoNormal><span style="mso-spacerun:yes"> 甲 <span style="mso-spacerun:no"> 乙  丙 </span> 丁 </span></p>',
      )[0]
        ?.runs.map((r) => r.text)
        .join(''),
    ).toBe(' 甲 乙 丙  丁 ');
    expect(
      parse('<p> 甲   乙 </p>')[0]
        ?.runs.map((r) => r.text)
        .join(''),
    ).toBe('甲 乙');
  });

  it('Word 的段落与字符类样式继承字号、字体分桶和对齐，内联样式覆盖类样式', () => {
    const html = `<html><head><style><!--
      body { font-size:12pt; font-family:"Times New Roman"; }
      p.MsoNormal, li.MsoNormal, div.MsoNormal {
        font-size:16pt; mso-fareast-font-family:仿宋; text-align:center;
      }
      span.Strong { font-weight:bold; color:red; }
    --></style></head><body><p class='MsoNormal'>普通<span class="Strong">强调</span><span
      class="Strong" style="font-weight:normal;color:blue;font-size:18pt">覆盖</span></p></body></html>`;
    const paragraphs = parse(html);
    expect(isTrustedHtml(html)).toBe(true);
    expect(paragraphs[0]?.justification).toBe('center');
    const runs = paragraphs[0]?.runs ?? [];
    expect(runs.map((r) => [r.text, r.patch.size, r.patch.bold, r.patch.color])).toEqual([
      ['普通', 320, false, undefined],
      ['强调', 320, true, 'FF0000'],
      ['覆盖', 360, false, '0000FF'],
    ]);
    expect(runs.every((r) => r.patch.fonts?.eastAsia === '仿宋')).toBe(true);
    expect(runs[0]?.patch.fonts?.ascii).toBe('Times New Roman');
  });

  it('规则按选择器优先级与源码顺序合并，不受 class 属性里名字顺序影响', () => {
    const paragraphs = parse(`<style>
      #target { font-size:20pt; }
      p.MsoNormal { font-size:16pt; color:red; }
      .MsoNormal { font-size:10pt; color:blue; }
      .A.B { font-weight:bold; }
      .A { font-weight:normal; }
      .B { font-style:italic; }
    </style><p class="MsoNormal B A" id="target">甲</p>
    <style>p.MsoNormal {color:green}</style><p class="MsoNormal A B">乙</p>`);
    expect(paragraphs.map((p) => p.runs[0]?.patch)).toMatchObject([
      { size: 400, color: '008000', bold: true, italic: true },
      { size: 320, color: '008000', bold: true, italic: true },
    ]);
  });

  it('important 高于普通内联声明，同等级内联与重复声明仍按 CSS 顺序覆盖', () => {
    const runs = parse(`<style>
      p.MsoNormal {font-size:18pt !important; color:red !important; font-weight:bold !important;}
      #target {font-size:24pt; color:green !important;}
    </style><p class="MsoNormal" id="target"
      style="font-size:12pt; color:blue !important; font-weight:normal">甲</p>
    <p class="MsoNormal" style="font-size:20pt !important;font-size:10pt;
      font-size:22pt !important">乙</p>`).map((p) => p.runs[0]?.patch);
    expect(runs).toMatchObject([
      { size: 360, color: '0000FF', bold: true },
      { size: 440, color: 'FF0000', bold: true },
    ]);
  });

  it('原文解析保留字体名内的分号和注释标记，规则中的引号、括号不截断后续声明', () => {
    const runs =
      parse(`<style>
      /* } ; .MsoNormal {font-size:99pt;} */
      p.MsoNormal {font-family:'A;/*原样*/B'; font-size:16pt;
        background:url("data:image/svg+xml;utf8,<svg>{}</svg>"); color: red;}
      span.Em {font-style: /* 注释 */ italic;}
    </style><p class="MsoNormal"><span class="Em">中文</span></p>`)[0]?.runs ?? [];
    expect(runs[0]?.patch).toMatchObject({
      size: 320,
      color: 'FF0000',
      italic: true,
      fonts: { ascii: 'A;/*原样*/B', eastAsia: 'A;/*原样*/B' },
    });
  });

  it('已支持的长短属性共享优先级，类中的简写不盖住内联长属性', () => {
    const runs =
      parse(`<style>
      span.Mark {background:yellow;text-decoration:underline;text-emphasis:filled dot;}
      .Page {page-break-before:always !important;}
    </style><p class=MsoNormal><span class=Mark style="background-color:blue;
      text-decoration-line:none;text-emphasis-style:none">正文</span><br class=Page
      style="break-before:auto">下一页</p>`)[0]?.runs ?? [];
    expect(runs[0]).toMatchObject({
      text: '正文',
      patch: {
        shading: { pattern: 'clear', color: 'auto', fill: '0000FF' },
        underline: 'none',
        emphasis: 'none',
      },
    });
    expect(runs.map((r) => r.text).join('')).toBe('正文\f下一页');
  });

  it('只应用支持的简单选择器，整块 at-rule 与外部样式跳过，不泄漏其声明或文字', () => {
    const paragraphs = parse(`<style>
      @import url("https://example.invalid/word.css");
      @font-face {font-family:Other;src:url(other.woff2);}
      @page WordSection1 {size:595.3pt 841.9pt;}
      @media print {p.MsoNormal {font-size:99pt;color:red;}}
      div p.MsoNormal {font-weight:bold;}
      .123 {font-style:italic;}
      p.MsoNormal {font-size:16pt;}
    </style><link rel="stylesheet" href="https://example.invalid/other.css">
    <div><p class=MsoNormal>正文</p></div>`);
    expect(paragraphs).toHaveLength(1);
    expect(paragraphs[0]?.runs[0]).toMatchObject({ text: '正文', patch: { size: 320, bold: false } });
    expect(paragraphs[0]?.runs[0]?.patch).not.toHaveProperty('color');
  });

  it('类样式中的隐藏文字和编号占位不粘入正文，分页符仍走已有转换', () => {
    const paragraphs = parse(`<style>
      .Hidden {display:none;} .Number {mso-list:Ignore;}
      .AlsoHidden {mso-hide:all;} .Page {page-break-before:always;}
    </style><p class=MsoNormal><span class=Number>1.</span>甲<span class=Hidden>隐藏</span><span
    class=AlsoHidden>隐藏</span><br class=Page>乙</p>`);
    expect(paragraphs.map((p) => p.runs.map((r) => r.text).join(''))).toEqual(['甲\f乙']);
  });

  it('非 CSS 与仅打印的 style 不应用，正文根节点的类样式同样继承或隐藏', () => {
    const html = `<style type="text/plain">.MsoNormal {font-size:99pt;}</style>
      <style media="print">.MsoNormal {font-size:90pt;}</style>
      <style media="screen and (max-width: 10px)">.MsoNormal {color:red;}</style>
      <style media="screen, print">body.MsoNormal {font-size:16pt;}</style>
      <body class=MsoNormal><p>正文</p></body>`;
    expect(parse(html)[0]?.runs[0]?.patch).toMatchObject({ size: 320 });
    expect(parse(html)[0]?.runs[0]?.patch).not.toHaveProperty('color');
    expect(parse('<style>body.MsoNormal {display:none}</style><body class=MsoNormal>隐藏</body>')).toEqual(
      [],
    );
  });

  it('网页类样式不改变合并格式策略，宿主页面的同名类也不会影响粘贴', () => {
    const hostStyle = document.head.appendChild(document.createElement('style'));
    hostStyle.textContent = '.FromHost {font-weight:bold;font-size:99pt;}';
    try {
      const paragraphs = parse(`<style>.Web {font-weight:bold;font-family:Arial;font-size:24px;}</style>
        <p class=Web><span class=FromHost>普通</span><b>加粗</b></p>`);
      expect(paragraphs).toEqual([
        {
          runs: [
            { text: '普通', patch: {} },
            { text: '加粗', patch: { bold: true } },
          ],
        },
      ]);
    } finally {
      hostStyle.remove();
    }
  });

  it('普通网页类名中间出现 Mso 不改为保留源格式', () => {
    const html = '<p class="Other-MsoNormal" style="font-size:99pt;color:red">普通</p>';
    expect(isTrustedHtml(html)).toBe(false);
    expect(parse(html)).toEqual([{ runs: [{ text: '普通', patch: {} }] }]);
  });

  it.each(['class=MsoNormal', "class='MsoNormal'", 'class="Other MsoNormal"'])(
    'Word 标识 %s 允许读取内嵌类样式',
    (attribute) => {
      const html = `<style>.MsoNormal {font-size:16pt;}</style><p ${attribute}>正文</p>`;
      expect(isTrustedHtml(html)).toBe(true);
      expect(parse(html)[0]?.runs[0]?.patch.size).toBe(320);
    },
  );

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
    const common = {
      strike: false,
      size: 320,
      fonts,
      highlight: 'none',
      shading: noShading,
      emphasis: 'none',
    };
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
    const off = {
      italic: false,
      underline: 'none',
      strike: false,
      vertAlign: 'baseline',
      highlight: 'none',
      shading: noShading,
      emphasis: 'none',
    } as const;
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

  it('高亮与字符底纹：复制写 background + mso-highlight，读回分得清哪个是高亮；Word 的 CSS 色名与块背景', () => {
    const html = fragmentToHtml({
      paragraphs: [
        {
          justification: 'left',
          runs: [
            { text: '亮', props: { ...base, highlight: 'darkYellow' } },
            // 中文版 Word「字符底纹」按钮的写法：白底上 15% 的黑网点
            { text: '纹', props: { ...base, shading: { pattern: 'pct15', color: 'auto', fill: 'FFFFFF' } } },
            { text: '素', props: base },
          ],
        },
      ],
    });
    expect(html).toContain('background:#808000;mso-highlight:darkYellow');
    expect(html).toContain('background:#d9d9d9');
    const runs = parse(html)[0]?.runs ?? [];
    expect(runs.map((r) => [r.text, r.patch.highlight, r.patch.shading])).toEqual([
      ['亮', 'darkYellow', noShading],
      ['纹', 'none', { pattern: 'clear', color: 'auto', fill: 'D9D9D9' }],
      ['素', 'none', noShading],
    ]);

    // Word 自己写 CSS 色名；段落（块）上的背景是段落底纹，不能铺到 run 上
    const word = `<html xmlns:o="urn:schemas-microsoft-com:office:office"><body>
<p class=MsoNormal style='background:silver'><span style='background:lime;mso-highlight:lime'>绿</span>普通</p>
</body></html>`;
    const wordRuns = parse(word)[0]?.runs ?? [];
    expect(wordRuns.map((r) => [r.text, r.patch.highlight, r.patch.shading])).toEqual([
      ['绿', 'green', noShading],
      ['普通', 'none', noShading],
    ]);
  });

  it.each([
    'background:rgb(18, 52, 86)',
    'background-color:rgb(18, 52, 86)',
    'background:RGB(18, 52, 86)',
    'background:rgb(18,\t52,\n86) no-repeat',
    'background:none repeat scroll rgb(18, 52, 86)',
    'background:url("red blue.png") center / cover no-repeat rgb(18, 52, 86)',
    'background:linear-gradient(red, blue) rgb(18, 52, 86)',
  ])('可信来源的字符底纹完整读取函数颜色 %s', (declaration) => {
    const runs = parse(`<p class=MsoNormal><span style='${declaration}'>底纹</span></p>`)[0]?.runs;
    expect(runs?.[0]?.patch.shading).toEqual({ pattern: 'clear', color: 'auto', fill: '123456' });
    expect(runs?.[0]?.patch.highlight).toBe('none');
  });

  it('类样式中的 RGB 底纹遵循优先级，mso 高亮仍优先于背景', () => {
    const runs = parse(`<style>.Rgb {background:rgb(18, 52, 86) !important;}</style>
      <p class=MsoNormal><span class=Rgb style="background:red">底纹</span><span class=Rgb style="mso-highlight:rgb(0, 255, 0)">高亮</span></p>`)[0]
      ?.runs;
    expect(runs?.map((r) => [r.text, r.patch.highlight, r.patch.shading])).toEqual([
      ['底纹', 'none', { pattern: 'clear', color: 'auto', fill: '123456' }],
      ['高亮', 'green', noShading],
    ]);
  });

  it.each([
    'url("image red blue.png")',
    'url(image\\ red\\ blue.png)',
    'linear-gradient( red , blue )',
    'image-set(url("red.png") 1x, url("blue.png") 2x)',
    '"image red blue"',
  ])('背景函数或字符串里的颜色词不被误当作字符底纹 %s', (background) => {
    const runs = parse(`<p class=MsoNormal><span style='background:${background}'>普通</span></p>`)[0]?.runs;
    expect(runs?.[0]?.patch.shading).toEqual(noShading);
  });

  it('着重号：复制写 text-emphasis（中文的点在字下），读回成 w:em；不在可信来源里的不带', () => {
    const html = fragmentToHtml({
      paragraphs: [
        {
          justification: 'left',
          runs: [
            { text: '重', props: { ...base, emphasis: 'dot' } },
            { text: '圈', props: { ...base, emphasis: 'circle' } },
            { text: '素', props: base },
          ],
        },
      ],
    });
    expect(html).toContain('text-emphasis:filled dot;text-emphasis-position:under right');
    expect(html).toContain('text-emphasis:open circle;text-emphasis-position:over right');
    const runs = parse(html)[0]?.runs ?? [];
    expect(runs.map((r) => [r.text, r.patch.emphasis])).toEqual([
      ['重', 'dot'],
      ['圈', 'circle'],
      ['素', 'none'],
    ]);
    const web = parse('<p><span style="text-emphasis:filled sesame">点</span></p>')[0]?.runs ?? [];
    expect(web[0]?.patch).not.toHaveProperty('emphasis');
  });

  it('网页 HTML 只带开关且只加不减，字体字号颜色对齐一律不带；块拆段、段内 <br> 是软换行、空白折叠', () => {
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
      { runs: [{ text: '甲\n乙', patch: {} }] },
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

  it('分页符按 Word 的写法出入：自己复制的读回 U+000C，Word 写在段首的也留住，普通 <br> 仍是软换行', () => {
    const html = fragmentToHtml({
      paragraphs: [{ justification: 'left', runs: [{ text: '甲\f乙', props: base }] }],
    });
    expect(html).toContain('page-break-before:always');
    expect(parse(html)[0]?.runs.map((r) => r.text)).toEqual(['甲\f乙']);
    const word =
      "<html xmlns:o='urn:schemas-microsoft-com:office:office'><body><p class=MsoNormal>前</p>" +
      "<p class=MsoNormal><span><br clear=all style='mso-special-character:line-break;page-break-before:always'></span>后<br>尾</p></body></html>";
    expect(parse(word).map((p) => p.runs.map((r) => r.text).join(''))).toEqual(['前', '\f后\n尾']);
  });

  it.each(['\n', '\n\n', '\n正文', '正文\n', '正文\n\n', '\n正文\n'])(
    '自己复制的软换行片段 %j 往返保留原段落与全部换行',
    (text) => {
      const paragraphs = parse(
        fragmentToHtml({
          paragraphs: [{ justification: 'left', runs: [{ text, props: base }] }],
        }),
      );
      expect(paragraphs).toHaveLength(1);
      expect(paragraphs[0]?.runs.map((r) => r.text).join('')).toBe(text);
    },
  );

  it('跨格式 run 的段末换行与后续空段分别保留', () => {
    const paragraphs = parse(
      fragmentToHtml({
        paragraphs: [
          {
            justification: 'center',
            runs: [
              { text: '正文', props: base },
              { text: '\n\n', props: { ...base, bold: true } },
            ],
          },
          { justification: 'left', runs: [] },
        ],
      }),
    );
    expect(paragraphs).toHaveLength(2);
    expect(paragraphs[0]?.runs.map((r) => [r.text, r.patch.bold])).toEqual([
      ['正文', false],
      ['\n\n', true],
    ]);
    expect(paragraphs[1]?.runs).toEqual([]);
  });

  it('Word 显式软换行在段首、段末及纯换行段保留，类样式同样生效', () => {
    const paragraphs = parse(`<style>.Soft {mso-special-character:line-break;}</style>
      <p class=MsoNormal><br class=Soft>正文<br class=Soft></p>
      <p class=MsoNormal><span><br style="mso-special-character:line-break"></span></p>`);
    expect(paragraphs.map((p) => p.runs.map((r) => r.text).join(''))).toEqual(['\n正文\n', '\n']);
  });

  it('网页段首与连续换行仍在同一段，段末占位只去掉最后一个普通 br', () => {
    const paragraphs = parse(
      '<p><br>正文<br><br><b>&nbsp;</b><i>&nbsp;</i></p><div><br></div><p><br><br></p><p>普通<b> </b><br></p>',
    );
    expect(paragraphs.map((p) => p.runs.map((r) => r.text).join(''))).toEqual(['\n正文\n', '', '\n', '普通']);
  });
});
