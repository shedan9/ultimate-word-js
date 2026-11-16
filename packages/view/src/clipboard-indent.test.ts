// @vitest-environment jsdom
import { createDiagnosticSink } from '@uw/core';
import type { Indent, ParaProps, Style, TextHistoryOptions } from '@uw/model';
import {
  createTextEditor,
  DEFAULT_SECTION_PROPS,
  DEFAULT_SETTINGS,
  EMPTY_NUMBERING,
  EMPTY_THEME,
  fragmentOfRange,
  paragraphText,
  parseStyles,
  rangeOfNode,
  resolveBody,
  walkParagraphs,
} from '@uw/model';
import { describe, expect, it } from 'vitest';
import { fragmentToHtml, htmlToParagraphs } from './clipboard.ts';
import { ZERO_INDENT } from './clipboard-indent.ts';
import { createEditingController } from './editing.ts';

const parse = (html: string) => htmlToParagraphs(html, new DOMParser());
const word = (style: string, text = '正文') => `<p class=MsoNormal style="${style}">${text}</p>`;

function target(props: ParaProps = {}, styled = false, options: TextHistoryOptions = {}) {
  const style: Style = {
    id: 'Target',
    type: 'paragraph',
    name: 'Target',
    basedOn: undefined,
    next: undefined,
    isDefault: false,
    paraProps: { indent: { left: 960, right: 480, leftChars: 300, firstLineChars: 200, hanging: 240 } },
    runProps: {},
    tableProps: {},
    rowProps: {},
    cellProps: {},
    conditional: new Map(),
  };
  const ctx = {
    styles: {
      ...parseStyles(undefined, createDiagnosticSink()),
      chainOf: (id: string | undefined) => (id === 'Target' ? [style] : []),
    },
    theme: EMPTY_THEME,
    settings: DEFAULT_SETTINGS,
    numbering: EMPTY_NUMBERING,
  };
  const editor = createTextEditor(
    {
      sections: [
        {
          id: 's',
          props: DEFAULT_SECTION_PROPS,
          blocks: [
            {
              kind: 'paragraph',
              id: 'p',
              props: { ...(styled ? { styleId: 'Target' } : {}), ...props },
              runs: [{ kind: 'run', id: 'r', props: {}, content: [{ kind: 'text', text: '前后' }] }],
            },
          ],
        },
      ],
    },
    options,
  );
  const state = createEditingController(editor);
  const at = { nodeId: 'r', contentIndex: 0, offset: 1 };
  state.select({ start: at, end: at });
  const resolved = () => [...walkParagraphs(resolveBody(ctx, editor.body))];
  return { editor, state, resolved };
}

describe('剪贴板段落缩进', () => {
  it.each([
    ['margin-left:24pt;margin-right:12px;text-indent:18pt', { left: 480, right: 180, firstLine: 360 }],
    ['margin-left:-.5in;margin-right:2.54cm;text-indent:-12pt', { left: -720, right: 1440, hanging: 240 }],
    ['margin:0 1pc 0 25.4mm;text-indent:0', { left: 1440, right: 240 }],
    ['margin:.5pt 1pt 2pt', { left: 20, right: 20 }],
    ['margin:1pt 2pt 3pt 4pt', { left: 80, right: 40 }],
    ['margin:3pt', { left: 60, right: 60 }],
    ['text-indent:32pt;mso-char-indent-count:2', { firstLine: 640, firstLineChars: 200 }],
    ['text-indent:-32pt;mso-char-indent-count:-2.25', { hanging: 640, hangingChars: 225 }],
    ['text-indent:18pt;mso-char-indent-count:0', { firstLine: 360 }],
    [
      'margin-left:24pt;mso-para-margin-left:2gd;mso-para-margin-right:1.5gd',
      { leftChars: 200, rightChars: 150 },
    ],
    ['text-indent:-12pt;mso-char-indent-count:2', { firstLineChars: 200 }],
  ] as const)('读取 Word 缩进 %s', (style, expected) => {
    expect(parse(word(style))[0]?.indent).toEqual({ ...ZERO_INDENT, ...expected });
  });

  it.each([
    '12',
    'auto',
    '10%',
    '2em',
    'calc(10pt + 2pt)',
    '1e999pt',
    'Infinitypt',
    '999999999999999999999pt',
  ])('非法或未支持的长度 %s 不产生错误坐标，首行沿用源继承', (invalid) => {
    const parsed = parse(`<div class=MsoNormal style="margin-left:40pt;text-indent:12pt">
        <p style="margin-left:${invalid};text-indent:${invalid}">正文</p></div>`);
    expect(parsed[0]?.indent).toEqual({ ...ZERO_INDENT, firstLine: 240 });
  });

  it('类与内联的 margin 简写 / 长属性按 important 与源码顺序逐边级联', () => {
    const parsed = parse(`<style>
      p.MsoNormal { margin:0 12pt 0 24pt !important; text-indent:32pt; }
      p.Reset { margin-left:48pt !important; }
      </style><p class="MsoNormal Reset" style="margin:0;margin-right:18pt!important">甲</p>
      <p class=MsoNormal style="margin:0 0 0 10pt!important;margin-left:20pt!important">乙</p>
      <p class=MsoNormal style="margin-left:20pt!important;margin:0!important">丙</p>`);
    expect(parsed.map((p) => p.indent)).toEqual([
      { ...ZERO_INDENT, left: 960, right: 360, firstLine: 640 },
      { ...ZERO_INDENT, left: 400, firstLine: 640 },
      { ...ZERO_INDENT, firstLine: 640 },
    ]);
  });

  it('左右缩进不继承容器；首行继承源父级，行内伪缩进不改段落，零显式取消', () => {
    const parsed =
      parse(`<div class=MsoNormal style="margin-left:90pt;text-indent:32pt;mso-char-indent-count:2">
      <p>甲<span style="margin-left:99pt;text-indent:-99pt">乙</span></p>
      <p style="text-indent:0">丙</p><p style="mso-char-indent-count:0">丁</p></div>`);
    expect(parsed.map((p) => p.indent)).toEqual([
      { ...ZERO_INDENT, firstLine: 640, firstLineChars: 200 },
      ZERO_INDENT,
      ZERO_INDENT,
    ]);
  });

  it('网页缩进不导入，Word 没写缩进时显式归零，纯文本合并目标格式', () => {
    expect(parse('<p style="margin-left:24pt;text-indent:32pt">网页</p>')[0]?.indent).toBeUndefined();
    expect(parse(word(''))[0]?.indent).toEqual(ZERO_INDENT);
    const { state, resolved } = target({ indent: { left: 400, firstLineChars: 200 } });
    state.insertParagraphs(parse('<p style="margin-left:99pt">甲</p><p>乙</p>'));
    expect(resolved().map((p) => p.props.indent)).toEqual([
      { ...ZERO_INDENT, left: 400, firstLineChars: 200 },
      { ...ZERO_INDENT, left: 400, firstLineChars: 200 },
    ]);
  });

  it.each([false, true])(
    '多段粘贴清除目标样式缩进，末段恢复原始字段（有样式 %s），单次撤销重做',
    (styled) => {
      const original = { left: 500, rightChars: 150, hangingChars: 100 };
      const { editor, state, resolved } = target({ indent: original, justification: 'right' }, styled);
      const before = editor.body;
      state.insertParagraphs(
        parse(
          `${word('margin-left:12pt;text-indent:24pt', '甲')}${word('', '')}${word('text-indent:-99pt', '乙')}`,
        ),
      );
      expect(resolved().map((p) => paragraphText(p))).toEqual(['前甲', '', '乙后']);
      expect(
        resolved()
          .slice(0, 2)
          .map((p) => p.props.indent),
      ).toEqual([{ ...ZERO_INDENT, left: 240, firstLine: 480 }, ZERO_INDENT]);
      const raw = [...walkParagraphs(editor.body)];
      expect(raw[2]?.props.indent).toEqual(original);
      expect(raw[2]?.props.justification).toBe('right');
      const pasted = editor.body;
      editor.undo();
      expect(editor.body).toBe(before);
      expect(editor.canUndo).toBe(false);
      editor.redo();
      expect(editor.body).toEqual(pasted);
    },
  );

  it('只复制段内文字不带段落标记，粘贴保持目标缩进', () => {
    const { state, resolved } = target({ indent: { firstLineChars: 200 } });
    state.insertParagraphs(parse(word('text-indent:-24pt;margin-left:48pt')));
    expect(resolved()[0]?.props.indent).toEqual({ ...ZERO_INDENT, firstLineChars: 200 });
  });

  it('模型复制 → HTML → 粘贴保留全部八字段，包括字符单位盖住的兼容值', () => {
    const indent: Required<Indent> = {
      left: 400,
      right: 120,
      firstLine: 640,
      hanging: 300,
      leftChars: 250,
      rightChars: 100,
      firstLineChars: 200,
      hangingChars: 150,
    };
    const { editor, resolved } = target({ indent });
    const body = resolveBody(
      {
        styles: parseStyles(undefined, createDiagnosticSink()),
        theme: EMPTY_THEME,
        settings: DEFAULT_SETTINGS,
        numbering: EMPTY_NUMBERING,
      },
      editor.body,
    );
    const p = [...walkParagraphs(body)][0];
    const range = p && rangeOfNode(p);
    if (!range) throw new Error('缺少测试段落');
    const html = fragmentToHtml(fragmentOfRange(body, range));
    expect(html).toContain('data-uw-indent=');
    expect(html).toContain('margin-left:2.5em');
    expect(html).toContain('text-indent:-1.5em');
    expect(html).toContain('mso-char-indent-count:-1.5');
    const paragraphs = parse(html);
    expect(paragraphs[0]?.indent).toEqual(indent);
    const dest = target({}, true);
    dest.state.insertParagraphs([...paragraphs, { runs: [] }]);
    expect(dest.resolved()[0]?.props.indent).toEqual(indent);
    expect(resolved()[0]?.props.indent).toEqual(indent);
  });

  it.each(['bad', '[]', '{}', '{"left":"20"}', 'null'])('损坏的内部缩进标记 %s 回退到 CSS', (attribute) => {
    const html = `<div data-uw-fragment=1><p data-uw-indent='${attribute}' style="margin-left:12pt">正文</p></div>`;
    expect(parse(html)[0]?.indent).toEqual({ ...ZERO_INDENT, left: 240 });
  });

  it('预提交失败时，缩进、正文、选区与历史整批保留', () => {
    const { editor, state } = target({ indent: { firstLineChars: 200 } }, false, {
      validate: () => {
        throw new Error('拒绝提交');
      },
    });
    const original = editor.body;
    const selection = state.selection;
    expect(() => state.insertParagraphs(parse(`${word('margin-left:12pt')}${word('')}`))).toThrow('拒绝提交');
    expect(editor.body).toBe(original);
    expect(state.selection).toEqual(selection);
    expect(editor.canUndo).toBe(false);
  });
});
