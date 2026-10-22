import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import { contentControlSpans, rangeOfContentControl, textOfContentControl } from './content-control.ts';
import { parseBody } from './parse-body.ts';
import { createTextEditor } from './text-transaction.ts';

const control = (tag: string, content: string, pr = '') =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/>${pr}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const text = '<w:p><w:r><w:t>原值</w:t></w:r></w:p>';

function editorOf(content: string, pr = '') {
  return createTextEditor(
    parseBody(
      parseXml(`<w:document><w:body>${control('outer', content, pr)}</w:body></w:document>`),
      createDiagnosticSink(),
    ),
  );
}

describe.each([
  ['首段空白', `<w:p/>${text}`],
  ['末段空白', `${text}<w:p/>`],
  ['首末空白', `<w:p/>${text}<w:p/>`],
  ['全为空段', '<w:p/><w:p/><w:p/>'],
])('块级控件范围包含%s', (_label, content) => {
  it.each(['新值', '', '甲\n乙'])('填写 %j 覆盖完整内容，撤销重做恢复全部段落', (value) => {
    const editor = editorOf(content);
    const before = editor.body;
    const span = contentControlSpans(before)[0];
    if (span === undefined) throw new Error('样本缺少控件');
    editor.tx((t) => void t.fillContentControl(span.control.id, value));
    const after = editor.body;
    const filled = contentControlSpans(after)[0];
    if (filled === undefined) throw new Error('填写后缺少控件');
    expect(filled.paragraphs).toHaveLength(1);
    expect(textOfContentControl(filled)).toBe(value);
    editor.undo();
    expect(editor.body).toBe(before);
    editor.redo();
    expect(editor.body).toBe(after);
  });
});

it.each(['head', 'tail'])('范围保留 %s 空段的段落锚点', (side) => {
  const editor = editorOf(side === 'head' ? `<w:p/>${text}` : `${text}<w:p/>`);
  const span = contentControlSpans(editor.body)[0];
  if (span === undefined) throw new Error('样本缺少控件');
  const empty = side === 'head' ? span.paragraphs[0] : span.paragraphs.at(-1);
  expect(rangeOfContentControl(span)?.[side === 'head' ? 'start' : 'end']).toEqual({
    nodeId: empty?.id,
    contentIndex: 0,
    offset: 0,
  });
});

it.each(['head', 'tail'])('覆盖移除 %s 空段中的子外壳，锁定外壳时整次回滚', (side) => {
  for (const locked of [false, true]) {
    const inner = control('inner', '<w:p/>', locked ? '<w:lock w:val="sdtLocked"/>' : '');
    const editor = editorOf(side === 'head' ? inner + text : text + inner);
    const before = editor.body;
    const outer = contentControlSpans(before).find((s) => s.control.tag === 'outer');
    if (outer === undefined) throw new Error('样本缺少外层');
    const fill = () => editor.tx((t) => void t.fillContentControl(outer.control.id, '覆盖'));
    if (locked) {
      expect(fill).toThrow(/锁定了外壳/);
      expect(editor.body).toBe(before);
      expect(editor.canUndo).toBe(false);
    } else {
      fill();
      expect(contentControlSpans(editor.body).map((s) => [s.control.tag, textOfContentControl(s)])).toEqual([
        ['outer', '覆盖'],
      ]);
    }
  }
});

describe.each(['dropDownList', 'comboBox'])('%s 空段边界', (type) => {
  it.each(['string', 'index'])('%s 入口覆盖空段，同步绑定与保存值，重复提交保留重做', (method) => {
    const editor = editorOf(
      `<w:p/>${text}<w:p/>`,
      `<w:dataBinding w:xpath="/outer"/><w:${type} w:lastValue="old"><w:listItem w:displayText="原值" w:value="new"/></w:${type}>`,
    );
    const before = editor.body;
    const span = contentControlSpans(before)[0];
    if (span === undefined) throw new Error('样本缺少控件');
    const fill = () =>
      editor.tx(
        (t) =>
          void (method === 'index'
            ? t.selectContentControlItem(span.control.id, 0)
            : t.fillContentControl(span.control.id, 'new')),
      );
    fill();
    const filled = contentControlSpans(editor.body)[0];
    expect(filled?.paragraphs).toHaveLength(1);
    expect(filled?.control.lastValue).toBe('new');
    expect(filled?.control.dataBound).toBeUndefined();
    expect(filled && textOfContentControl(filled)).toBe('原值');
    const range = filled && rangeOfContentControl(filled);
    if (range === undefined) throw new Error('填写后缺少范围');
    editor.tx((t) => void t.setRunProps(range, { bold: true }));
    editor.undo();
    const after = editor.body;
    expect(fill()).toBeUndefined();
    expect(editor.body).toBe(after);
    expect(editor.canRedo).toBe(true);
    editor.undo();
    expect(editor.body).toBe(before);
  });
});

describe.each(['head', 'tail'])('%s 边界空段的后代内容保护', (side) => {
  it.each(['contentLocked', 'sdtContentLocked', 'dropDownList', 'checkbox'])(
    '%s 拒绝覆盖，捕获异常仍回滚格式、绑定与历史',
    (restriction) => {
      const pr = restriction.endsWith('Locked')
        ? `<w:lock w:val="${restriction}"/>`
        : restriction === 'checkbox'
          ? '<w14:checkbox/>'
          : '<w:dropDownList/>';
      const inner = control('inner', '<w:p/>', pr);
      const editor = editorOf(
        side === 'head' ? inner + text : text + inner,
        '<w:dataBinding w:xpath="/outer"/>',
      );
      const before = editor.body;
      const outer = contentControlSpans(before).find((s) => s.control.tag === 'outer');
      const range = outer && rangeOfContentControl(outer);
      if (outer === undefined || range === undefined) throw new Error('样本缺少外层');
      editor.tx((t) => void t.setRunProps(range, { bold: true }));
      editor.undo();
      expect(() =>
        editor.tx((t) => {
          t.setRunProps(range, { italic: true });
          try {
            t.fillContentControl(outer.control.id, '覆盖');
          } catch {}
        }),
      ).toThrow(/全部回滚/);
      expect(editor.body).toBe(before);
      expect(editor.canUndo).toBe(false);
      expect(editor.canRedo).toBe(true);
    },
  );
});
