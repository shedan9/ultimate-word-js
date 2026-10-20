import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import { contentControlSpans, rangeOfContentControl, textOfContentControl } from './content-control.ts';
import { parseBody } from './parse-body.ts';
import type { DocRange } from './position.ts';
import { createTextEditor } from './text-transaction.ts';

const run = '<w:r><w:rPr><w:b/></w:rPr><w:t>原值</w:t></w:r>';
const control = (tag: string, content: string, pr = '') =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/>${pr}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;

function editorOf(scope: string, lock = '', type = 'richText') {
  const inner = control('inner', run, `<w:text/>${lock}`);
  const content =
    scope === 'inline'
      ? inner
      : scope === 'block'
        ? `<w:p>${inner}</w:p>`
        : control('middle', `<w:p>${inner}</w:p>`);
  const outer = control('outer', content, `<w:${type}/>`);
  return createTextEditor(
    parseBody(
      parseXml(
        `<w:document><w:body>${scope === 'inline' ? `<w:p>${outer}</w:p>` : outer}</w:body></w:document>`,
      ),
      createDiagnosticSink(),
    ),
  );
}

describe.each(['inline', 'block', 'nestedBlock'])('覆盖以子控件开头的外层：%s', (scope) => {
  it.each(['新值', '原值', '', '甲\n乙'])('填写 %j 移除子控件并保留目标、位置与撤销重做', (value) => {
    const editor = editorOf(scope);
    const before = editor.body;
    const outer = contentControlSpans(before).find((s) => s.control.tag === 'outer');
    if (outer === undefined) throw new Error('样本缺少外层');
    const start = rangeOfContentControl(outer)?.start;
    let filled: DocRange | undefined;
    expect(
      editor.tx((t) => {
        filled = t.fillContentControl(outer.control.id, value);
      }),
    ).toBeDefined();
    const spans = contentControlSpans(editor.body);
    expect(spans.map((s) => s.control.tag)).toEqual(['outer']);
    if (spans[0] === undefined) throw new Error('填写后缺少目标');
    expect(textOfContentControl(spans[0])).toBe(value);
    expect(filled?.start).toEqual(start);
    expect(spans[0]?.runs[0]?.props.bold).toBe(true);
    const after = editor.body;
    editor.undo();
    expect(editor.body).toBe(before);
    editor.redo();
    expect(editor.body).toBe(after);
  });

  it.each(['sdtLocked', 'sdtContentLocked', 'contentLocked'])(
    '子控件 %s 仍阻止覆盖，捕获错误也回滚',
    (lock) => {
      const editor = editorOf(scope, `<w:lock w:val="${lock}"/>`);
      const before = editor.body;
      const outer = contentControlSpans(before).find((s) => s.control.tag === 'outer');
      if (outer === undefined) throw new Error('样本缺少外层');
      const range = rangeOfContentControl(outer);
      if (range === undefined) throw new Error('样本缺少范围');
      editor.tx((t) => void t.setRunProps(range, { italic: true }));
      editor.undo();
      expect(() =>
        editor.tx((t) => {
          t.setRunProps(range, { underline: 'single' });
          try {
            t.fillContentControl(outer.control.id, '原值');
          } catch {}
        }),
      ).toThrow(/全部回滚/);
      expect(editor.body).toBe(before);
      expect(editor.canUndo).toBe(false);
      expect(editor.canRedo).toBe(true);
    },
  );
});

it('移除子控件后仍遵守目标自身的单行规则', () => {
  const editor = editorOf('inline', '', 'text');
  const outer = contentControlSpans(editor.body).find((s) => s.control.tag === 'outer');
  if (outer === undefined) throw new Error('样本缺少外层');
  editor.tx((t) => void t.fillContentControl(outer.control.id, '甲\n乙'));
  expect(contentControlSpans(editor.body).map((s) => [s.control.tag, textOfContentControl(s)])).toEqual([
    ['outer', '甲 乙'],
  ]);
});

describe.each(['dropDownList', 'comboBox'])('%s 覆盖首个子控件', (type) => {
  it.each(['string', 'index'])('%s 入口同步保存值与绑定，重复提交保留重做', (method) => {
    const pr = `<w:dataBinding w:xpath="/outer"/><w:${type} w:lastValue="a"><w:listItem w:displayText="原值" w:value="a"/><w:listItem w:displayText="原值" w:value="b"/></w:${type}>`;
    const editor = createTextEditor(
      parseBody(
        parseXml(
          `<w:document><w:body><w:p>${control('outer', control('inner', run), pr)}</w:p></w:body></w:document>`,
        ),
        createDiagnosticSink(),
      ),
    );
    const before = editor.body;
    const id = contentControlSpans(before).find((s) => s.control.tag === 'outer')?.control.id;
    if (id === undefined) throw new Error('样本缺少外层');
    const fill = (value: number) =>
      editor.tx(
        (t) =>
          void (method === 'index'
            ? t.selectContentControlItem(id, value)
            : t.fillContentControl(id, value === 0 ? 'a' : 'b')),
      );
    fill(1);
    expect(
      contentControlSpans(editor.body).map((s) => [s.control.tag, s.control.lastValue, s.control.dataBound]),
    ).toEqual([['outer', 'b', undefined]]);
    fill(0);
    editor.undo();
    expect(fill(1)).toBeUndefined();
    expect(editor.canRedo).toBe(true);
    editor.undo();
    expect(editor.body).toBe(before);
  });
});

it('只覆盖目标的后代，保留祖先与相邻控件', () => {
  const outer = control('outer', control('inner', run));
  const sibling = control('sibling', '<w:r><w:t>旁边</w:t></w:r>');
  const editor = createTextEditor(
    parseBody(
      parseXml(
        `<w:document><w:body>${control('ancestor', `<w:p>${outer}${sibling}</w:p>`)}</w:body></w:document>`,
      ),
      createDiagnosticSink(),
    ),
  );
  const id = contentControlSpans(editor.body).find((s) => s.control.tag === 'outer')?.control.id;
  if (id === undefined) throw new Error('样本缺少外层');
  editor.tx((t) => void t.fillContentControl(id, '新值'));
  expect(contentControlSpans(editor.body).map((s) => [s.control.tag, textOfContentControl(s)])).toEqual([
    ['ancestor', '新值旁边'],
    ['outer', '新值'],
    ['sibling', '旁边'],
  ]);
});
