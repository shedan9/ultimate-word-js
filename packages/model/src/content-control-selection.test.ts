import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import { contentControlSpans, textOfContentControl } from './content-control.ts';
import { parseBody } from './parse-body.ts';
import { createTextEditor } from './text-transaction.ts';

const list = (type = 'dropDownList', lock = '') =>
  `<w:p><w:sdt><w:sdtPr><w:tag w:val="level"/>${lock}<w:${type} w:lastValue="a">` +
  '<w:listItem w:displayText="相同" w:value="a"/><w:listItem w:displayText="相同" w:value="b"/>' +
  '<w:listItem w:displayText="b" w:value="c"/></w:' +
  type +
  '></w:sdtPr><w:sdtContent><w:r><w:t>相同</w:t></w:r></w:sdtContent></w:sdt></w:p>';
const editorOf = (xml = list()) =>
  createTextEditor(
    parseBody(parseXml(`<w:document><w:body>${xml}</w:body></w:document>`), createDiagnosticSink()),
  );

function text(editor: ReturnType<typeof editorOf>) {
  const span = contentControlSpans(editor.body)[0];
  if (span === undefined) throw new Error('样本缺少控件');
  return textOfContentControl(span);
}

describe('按序号选择列表项', () => {
  it.each(['dropDownList', 'comboBox'])('%s 精确选择重复显示文字并区分保存值冲突', (type) => {
    const editor = editorOf(list(type));
    const before = editor.body;
    const id = contentControlSpans(before)[0]?.control.id ?? '';
    editor.tx((t) => void t.selectContentControlItem(id, 1));
    expect(editor.body.contentControls?.[id]?.lastValue).toBe('b');
    expect(text(editor)).toBe('相同');
    // 第二项的保存值恰好是第三项的显示文字，按字符串填写无法指定第二项。
    editor.undo();
    expect(editor.body).toBe(before);
    editor.redo();
    expect(editor.body.contentControls?.[id]?.lastValue).toBe('b');
    editor.tx((t) => void t.selectContentControlItem(id, 2));
    expect(editor.body.contentControls?.[id]?.lastValue).toBe('c');
    expect(text(editor)).toBe('b');
  });

  it.each([-1, 3, 0.5, NaN, Infinity, undefined, null, '1'])(
    '非法序号 %s 回滚之前的修改，即使调用方捕获错误',
    (index) => {
      const editor = editorOf();
      const before = editor.body;
      const id = contentControlSpans(before)[0]?.control.id ?? '';
      expect(() =>
        editor.tx((t) => {
          t.selectContentControlItem(id, 1);
          try {
            t.selectContentControlItem(id, index as number);
          } catch {}
        }),
      ).toThrow(/全部回滚/);
      expect(editor.body).toBe(before);
      expect(editor.canUndo).toBe(false);
    },
  );

  it('未改变选择不记历史，自身或父级内容锁仍拒绝', () => {
    const editor = editorOf();
    const id = contentControlSpans(editor.body)[0]?.control.id ?? '';
    expect(editor.tx((t) => void t.selectContentControlItem(id, 0))).toBeUndefined();
    for (const xml of [
      list('dropDownList', '<w:lock w:val="contentLocked"/>'),
      `<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent>${list()}</w:sdtContent></w:sdt>`,
    ]) {
      const locked = editorOf(xml);
      const before = locked.body;
      const target = contentControlSpans(before).find((s) => s.control.tag === 'level');
      expect(() => locked.tx((t) => void t.selectContentControlItem(target?.control.id ?? '', 1))).toThrow(
        /锁定/,
      );
      expect(locked.body).toBe(before);
    }
  });

  it('非列表与空列表拒绝，不扩大对外层下拉框的填写权限', () => {
    for (const xml of [
      list('comboBox').replace(/<w:listItem[^>]+\/>/g, ''),
      list().replaceAll('dropDownList', 'text'),
    ]) {
      const editor = editorOf(xml);
      const id = contentControlSpans(editor.body)[0]?.control.id ?? '';
      expect(() => editor.tx((t) => void t.selectContentControlItem(id, 0))).toThrow();
      expect(editor.canUndo).toBe(false);
    }
    const nested = editorOf(
      `<w:sdt><w:sdtPr><w:dropDownList/></w:sdtPr><w:sdtContent>${list()}</w:sdtContent></w:sdt>`,
    );
    const id = contentControlSpans(nested.body).find((s) => s.control.tag === 'level')?.control.id ?? '';
    expect(() => nested.tx((t) => void t.selectContentControlItem(id, 1))).toThrow(/下拉框只能/);
  });
});
