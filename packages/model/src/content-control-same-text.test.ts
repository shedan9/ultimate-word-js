import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import { contentControlSpans, rangeOfContentControl, textOfContentControl } from './content-control.ts';
import { parseBody } from './parse-body.ts';
import { createTextEditor } from './text-transaction.ts';

const formatted = '<w:r><w:rPr><w:b/></w:rPr><w:t>张</w:t></w:r><w:r><w:t>三</w:t></w:r>';
function editorOf(type: string, content = formatted, pr = '', block = false) {
  const sdt = `<w:sdt><w:sdtPr><w:tag w:val="field"/>${type}${pr}</w:sdtPr><w:sdtContent>${block ? `<w:p>${content}</w:p>` : content}</w:sdtContent></w:sdt>`;
  return createTextEditor(
    parseBody(
      parseXml(`<w:document><w:body>${block ? sdt : `<w:p>${sdt}</w:p>`}</w:body></w:document>`),
      createDiagnosticSink(),
    ),
  );
}
function spanOf(editor: ReturnType<typeof editorOf>) {
  const span = contentControlSpans(editor.body).find((s) => s.control.tag === 'field');
  if (span === undefined) throw new Error('样本缺少控件');
  return span;
}

describe.each([false, true])('重复填写文字保留内容：块级 %s', (block) => {
  it.each(['<w:text/>', '<w:richText/>', '<w:date/>'])(
    '%s 当前值不拆 run、不丢格式、不记历史，也不清掉重做栈',
    (type) => {
      const editor = editorOf(type, formatted, '', block);
      const before = editor.body;
      const span = spanOf(editor);
      const range = rangeOfContentControl(span);
      if (range === undefined) throw new Error('样本缺少范围');
      editor.tx((t) => void t.insertText(range.end, '甲'));
      editor.undo();
      expect(editor.body).toBe(before);
      expect(
        editor.tx((t) => {
          expect(t.fillContentControl(span.control.id, '张三')).toEqual(range);
        }),
      ).toBeUndefined();
      expect(editor.body).toBe(before);
      expect(editor.canUndo).toBe(false);
      expect(editor.canRedo).toBe(true);
      editor.redo();
      expect(textOfContentControl(spanOf(editor))).toBe('张三甲');
    },
  );

  it.each(['<w:text/>', '<w:richText/>'])('%s 仅解除绑定，与文字格式和位置无关，失败整次回滚', (type) => {
    const editor = editorOf(type, formatted, '<w:dataBinding w:xpath="/field"/>', block);
    const before = editor.body;
    const { id } = spanOf(editor).control;
    expect(() =>
      editor.tx((t) => {
        t.fillContentControl(id, '张三');
        t.fillContentControl('missing', '失败');
      }),
    ).toThrow(/没有内容控件/);
    expect(editor.body).toBe(before);
    expect(editor.canUndo).toBe(false);
    expect(editor.tx((t) => void t.fillContentControl(id, '张三'))).toBeDefined();
    const after = editor.body;
    expect(after.sections).toEqual(before.sections);
    expect(spanOf(editor).control.dataBound).toBeUndefined();
    expect(editor.tx((t) => void t.fillContentControl(id, '张三'))).toBeUndefined();
    editor.undo();
    expect(editor.body).toBe(before);
    editor.redo();
    expect(editor.body).toBe(after);
  });
});

it.each(['', '<w:dataBinding w:xpath="/field"/>'])(
  '日期字符串相同仍清 fullDate 与绑定，保留局部格式；再次填写无修改：%s',
  (binding) => {
    const editor = editorOf(
      '<w:date w:fullDate="2026-10-19T00:00:00Z"><w:dateFormat w:val="yyyy年M月d日"/></w:date>',
      formatted,
      binding,
    );
    const before = editor.body;
    const { id } = spanOf(editor).control;
    editor.tx((t) => void t.fillContentControl(id, '张三'));
    const after = editor.body;
    expect(after.sections).toEqual(before.sections);
    expect(spanOf(editor).control.date).toEqual({ format: 'yyyy年M月d日' });
    expect(spanOf(editor).control.dataBound).toBeUndefined();
    expect(editor.tx((t) => void t.fillContentControl(id, '张三'))).toBeUndefined();
    editor.undo();
    expect(editor.body).toBe(before);
    editor.redo();
    expect(editor.body).toBe(after);
  },
);

it('同事务先拆 run 改格式再填当前文字，保留刚设置的局部格式', () => {
  const editor = editorOf('<w:richText/>', '<w:r><w:t>张三</w:t></w:r>');
  const before = editor.body;
  const span = spanOf(editor);
  const start = rangeOfContentControl(span)?.start;
  if (start === undefined) throw new Error('样本缺少范围');
  editor.tx((t) => {
    t.setRunProps({ start, end: { ...start, offset: 1 } }, { bold: true });
    t.fillContentControl(span.control.id, '张三');
  });
  expect(spanOf(editor).runs.map((r) => [r.props.bold, r.content])).toEqual([
    [true, [{ kind: 'text', text: '张' }]],
    [undefined, [{ kind: 'text', text: '三' }]],
  ]);
  editor.undo();
  expect(editor.body).toBe(before);
});

it('软换行与制表位按规范化文字比较，空块级不凭空建立 run', () => {
  const content = '<w:r><w:rPr><w:b/></w:rPr><w:t>甲</w:t><w:br/><w:tab/><w:t>乙</w:t></w:r>';
  for (const [type, inner, value, block] of [
    ['<w:text w:multiLine="1"/>', content, '甲\r\n\t乙', false],
    ['<w:richText/>', content, '甲\r\t乙', true],
    ['<w:text/>', '', '', true],
    ['<w:text/>', '<w:r><w:t>甲 乙</w:t></w:r>', '甲\n乙', false],
  ] as const) {
    const editor = editorOf(type, inner, '', block);
    const before = editor.body;
    const span = spanOf(editor);
    expect(editor.tx((t) => void t.fillContentControl(span.control.id, value))).toBeUndefined();
    expect(editor.body).toBe(before);
  }
});

it('占位文字相同仍清占位标记并使用控件格式', () => {
  const editor = editorOf('<w:text/>', formatted, '<w:showingPlcHdr/><w:rPr><w:i/></w:rPr>');
  expect(editor.tx((t) => void t.fillContentControl(spanOf(editor).control.id, '张三'))).toBeDefined();
  expect(spanOf(editor).control.showingPlaceholder).toBe(false);
  expect(spanOf(editor).runs.every((r) => r.props.italic && r.props.bold === undefined)).toBe(true);
});

it('内容锁与下拉框祖先仍拒绝填写当前值', () => {
  const editor = editorOf('<w:text/>', formatted, '<w:lock w:val="contentLocked"/>');
  const before = editor.body;
  expect(() => editor.tx((t) => void t.fillContentControl(spanOf(editor).control.id, '张三'))).toThrow(
    /锁定/,
  );
  expect(editor.body).toBe(before);
  const inner = `<w:sdt><w:sdtPr><w:text/></w:sdtPr><w:sdtContent>${formatted}</w:sdtContent></w:sdt>`;
  const nested = editorOf('<w:dropDownList/>', inner);
  const id = contentControlSpans(nested.body).find((s) => s.control.type === 'text')?.control.id;
  if (id === undefined) throw new Error('样本缺少子控件');
  expect(() => nested.tx((t) => void t.fillContentControl(id, '张三'))).toThrow(/下拉框只能/);
});

it('外层相同文字仍执行原有覆盖逻辑，域与对象不因纯文字摘要相同被跳过', () => {
  const inner = `<w:sdt><w:sdtPr><w:text/></w:sdtPr><w:sdtContent>${formatted}</w:sdtContent></w:sdt>`;
  const nested = editorOf('<w:richText/>', inner);
  expect(nested.tx((t) => void t.fillContentControl(spanOf(nested).control.id, '张三'))).toBeDefined();
  for (const content of [
    '<w:r><w:t>张三</w:t><w:drawing/></w:r>',
    '<w:r><w:fldChar w:fldCharType="begin"/><w:instrText>PAGE</w:instrText><w:fldChar w:fldCharType="separate"/><w:t>张三</w:t><w:fldChar w:fldCharType="end"/></w:r>',
  ]) {
    const editor = editorOf('<w:richText/>', content);
    const before = editor.body;
    expect(() => editor.tx((t) => void t.fillContentControl(spanOf(editor).control.id, '张三'))).toThrow();
    expect(editor.body).toBe(before);
  }
});
