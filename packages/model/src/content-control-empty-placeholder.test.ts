import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import { contentControlSpans, textOfContentControl } from './content-control.ts';
import { parseBody } from './parse-body.ts';
import type { DocRange } from './position.ts';
import { createTextEditor } from './text-transaction.ts';

function editorOf(type: string, block: boolean, props = '<w:rPr><w:b/></w:rPr>', placeholder = true) {
  const run =
    '<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/><w:i/><w:color w:val="888888"/>' +
    '<w:lang w:eastAsia="zh-CN"/></w:rPr><w:t>请输入</w:t></w:r>';
  const control = `<w:sdt><w:sdtPr><w:tag w:val="field"/>${props}${placeholder ? '<w:showingPlcHdr/>' : ''}<w:dataBinding w:xpath="/field"/>${type}</w:sdtPr><w:sdtContent>${block ? `<w:p>${run}</w:p>` : run}</w:sdtContent></w:sdt>`;
  return createTextEditor(
    parseBody(
      parseXml(`<w:document><w:body>${block ? control : `<w:p>${control}</w:p>`}</w:body></w:document>`),
      createDiagnosticSink(),
    ),
  );
}

function spanOf(editor: ReturnType<typeof editorOf>) {
  const span = contentControlSpans(editor.body)[0];
  if (span === undefined) throw new Error('样本缺少控件');
  return span;
}

describe.each([false, true])('空值填写重置占位格式：块级 %s', (block) => {
  it.each(['<w:text/>', '<w:richText/>', '<w:date/>', '<w:comboBox/>'])(
    '%s 空文字槽使用控件格式，后续输入与历史一致',
    (type) => {
      const editor = editorOf(type, block);
      const before = editor.body;
      const id = spanOf(editor).control.id;
      let filled: DocRange | undefined;
      editor.tx((t) => {
        filled = t.fillContentControl(id, '');
      });
      const empty = editor.body;
      const span = spanOf(editor);
      expect(span.control.showingPlaceholder).toBe(false);
      expect(span.control.dataBound).toBeUndefined();
      expect(span.runs).toHaveLength(1);
      expect(span.runs[0]?.props).toEqual({ bold: true, langEastAsia: 'zh-CN' });
      expect(textOfContentControl(span)).toBe('');
      expect(filled?.start).toEqual(filled?.end);
      if (filled === undefined) throw new Error('样本缺少填写范围');
      const end = filled.end;
      editor.tx((t) => void t.insertText(end, '张三'));
      expect(spanOf(editor).runs[0]?.props).toEqual({ bold: true, langEastAsia: 'zh-CN' });
      editor.undo();
      expect(editor.body).toBe(empty);
      expect(editor.tx((t) => void t.fillContentControl(id, ''))).toBeUndefined();
      expect(editor.canRedo).toBe(true);
      editor.undo();
      expect(editor.body).toBe(before);
      editor.redo();
      expect(editor.body).toBe(empty);
      editor.redo();
      expect(textOfContentControl(spanOf(editor))).toBe('张三');
    },
  );

  it.each(['fill', 'select'])('列表选中空显示值也重置占位格式：%s', (method) => {
    const editor = editorOf(
      '<w:dropDownList><w:listItem w:displayText="" w:value="empty"/></w:dropDownList>',
      block,
    );
    editor.tx((t) => {
      const id = spanOf(editor).control.id;
      if (method === 'fill') t.fillContentControl(id, 'empty');
      else t.selectContentControlItem(id, 0);
    });
    const span = spanOf(editor);
    expect(textOfContentControl(span)).toBe('');
    expect(span.control.lastValue).toBe('empty');
    expect(span.runs[0]?.props).toEqual({ bold: true, langEastAsia: 'zh-CN' });
  });

  it('没有控件格式也清占位样式，错误时格式与元数据整次回滚', () => {
    const editor = editorOf('<w:text/>', block, '');
    const before = editor.body;
    const id = spanOf(editor).control.id;
    expect(() =>
      editor.tx((t) => {
        t.fillContentControl(id, '');
        t.fillContentControl('missing', '失败');
      }),
    ).toThrow(/没有内容控件/);
    expect(editor.body).toBe(before);
    expect(editor.canUndo).toBe(false);
    editor.tx((t) => {
      const range = t.fillContentControl(id, '');
      t.insertText(range.end, '张三');
    });
    expect(spanOf(editor).runs[0]?.props).toEqual({ langEastAsia: 'zh-CN' });
    editor.undo();
    expect(editor.body).toBe(before);
  });

  it('非占位控件填空保留原有格式', () => {
    const editor = editorOf('<w:richText/>', block, '<w:rPr><w:b/></w:rPr>', false);
    const before = spanOf(editor).runs[0]?.props;
    editor.tx((t) => void t.fillContentControl(spanOf(editor).control.id, ''));
    expect(spanOf(editor).runs[0]?.props).toEqual(before);
  });
});
