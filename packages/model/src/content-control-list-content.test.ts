import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import { contentControlSpans, rangeOfContentControl, textOfContentControl } from './content-control.ts';
import { parseBody } from './parse-body.ts';
import type { TextTransaction } from './text-transaction.ts';
import { createTextEditor } from './text-transaction.ts';

const formatted = '<w:r><w:rPr><w:b/></w:rPr><w:t>相</w:t></w:r><w:r><w:t>同</w:t></w:r>';
const nested = `${formatted}<w:sdt><w:sdtPr><w:text/></w:sdtPr><w:sdtContent><w:r><w:t/></w:r></w:sdtContent></w:sdt>`;

describe.each(['dropDownList', 'comboBox'])('%s 填当前值校验完整内容', (type) => {
  function editorOf(content: string, block = false, pr = '', display = '相同') {
    const sdt = `<w:sdt><w:sdtPr><w:tag w:val="field"/>${pr}<w:${type} w:lastValue="a"><w:listItem w:displayText="${display}" w:value="a"/><w:listItem w:displayText="${display}" w:value="b"/></w:${type}></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
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

  describe.each(['string', 'index'])('%s 填写入口', (method) => {
    function fill(t: TextTransaction, id: string, second = false) {
      return method === 'index'
        ? t.selectContentControlItem(id, second ? 1 : 0)
        : t.fillContentControl(id, second ? 'b' : 'a');
    }

    it.each([false, true])('普通多 run 内容保留格式、位置与重做：块级 %s', (block) => {
      const editor = editorOf(block ? `<w:p>${formatted}</w:p>` : formatted, block);
      const before = editor.body;
      const span = spanOf(editor);
      editor.tx((t) => void fill(t, span.control.id, true));
      expect(spanOf(editor).control.lastValue).toBe('b');
      expect(editor.body.sections).toEqual(before.sections);
      editor.undo();
      expect(editor.body).toBe(before);
      expect(
        editor.tx((t) => {
          expect(fill(t, span.control.id)).toEqual(rangeOfContentControl(span));
        }),
      ).toBeUndefined();
      expect(editor.body).toBe(before);
      expect(editor.canUndo).toBe(false);
      expect(editor.canRedo).toBe(true);
      editor.redo();
      expect(spanOf(editor).control.lastValue).toBe('b');
    });

    it.each([
      '<w:r><w:t>相同</w:t><w:drawing/></w:r>',
      '<w:r><w:fldChar w:fldCharType="begin"/><w:instrText>PAGE</w:instrText><w:fldChar w:fldCharType="separate"/><w:t>相同</w:t><w:fldChar w:fldCharType="end"/></w:r>',
      '<w:del w:id="1" w:author="甲"><w:r><w:delText>旧</w:delText></w:r></w:del><w:r><w:t>相同</w:t></w:r>',
    ])('对象、域或删除修订不能被文字摘要掩盖：%s', (content) => {
      const editor = editorOf(content, false, '<w:dataBinding w:xpath="/field"/>');
      const before = editor.body;
      const { id } = spanOf(editor).control;
      expect(textOfContentControl(spanOf(editor))).toBe('相同');
      expect(() => editor.tx((t) => void fill(t, id, true))).toThrow();
      expect(editor.body).toBe(before);
      expect(editor.canUndo).toBe(false);
    });

    it('覆盖外层相同文字仍移除嵌套控件，同撤销与重做', () => {
      const editor = editorOf(nested);
      const before = editor.body;
      const { id } = spanOf(editor).control;
      expect(contentControlSpans(before)).toHaveLength(2);
      expect(editor.tx((t) => void fill(t, id))).toBeDefined();
      expect(textOfContentControl(spanOf(editor))).toBe('相同');
      expect(contentControlSpans(editor.body)).toHaveLength(1);
      const after = editor.body;
      expect(spanOf(editor).runs.every((r) => r.contentControl === id)).toBe(true);
      editor.undo();
      expect(editor.body).toBe(before);
      editor.redo();
      expect(editor.body).toBe(after);
    });

    it.each([false, true])('摘要中的换行不能掩盖分页符或多段结构：块级 %s', (block) => {
      const content = block
        ? '<w:p><w:r><w:t>相</w:t></w:r></w:p><w:p><w:r><w:t>同</w:t></w:r></w:p>'
        : '<w:r><w:t>相</w:t><w:br w:type="page"/><w:t>同</w:t></w:r>';
      const editor = editorOf(content, block, '', '相&#10;同');
      const before = editor.body;
      const { id } = spanOf(editor).control;
      expect(textOfContentControl(spanOf(editor))).toBe('相\n同');
      expect(editor.tx((t) => void fill(t, id))).toBeDefined();
      const after = spanOf(editor);
      expect(after.paragraphs).toHaveLength(1);
      expect(after.runs.flatMap((r) => r.content).filter((c) => c.kind === 'break')).toEqual([
        { kind: 'break', breakType: 'line' },
      ]);
      expect(textOfContentControl(after)).toBe('相\n同');
      editor.undo();
      expect(editor.body).toBe(before);
    });

    it('捕获填写错误仍回滚同事务之前的格式修改与绑定状态', () => {
      const editor = editorOf(
        '<w:r><w:t>相同</w:t><w:drawing/></w:r>',
        false,
        '<w:dataBinding w:xpath="/field"/>',
      );
      const before = editor.body;
      const span = spanOf(editor);
      const range = rangeOfContentControl(span);
      if (range === undefined) throw new Error('样本缺少范围');
      expect(() =>
        editor.tx((t) => {
          t.setRunProps(range, { italic: true });
          try {
            fill(t, span.control.id);
          } catch {}
        }),
      ).toThrow(/全部回滚/);
      expect(editor.body).toBe(before);
      expect(editor.canUndo).toBe(false);
    });
  });
});
