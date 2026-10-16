import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import { contentControlSpans, rangeOfContentControl, textOfContentControl } from './content-control.ts';
import type { Body } from './nodes.ts';
import { parseBody } from './parse-body.ts';
import { createTextEditor } from './text-transaction.ts';

const control = (type: string, content = '<w:r><w:t>占位文字</w:t></w:r>', tag = 'field') =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:showingPlcHdr/><w:dataBinding w:xpath="/${tag}"/>${type}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const bodyOf = (xml: string) =>
  parseBody(parseXml(`<w:document><w:body>${xml}</w:body></w:document>`), createDiagnosticSink());
function spanOf(body: Body, tag = 'field') {
  const span = contentControlSpans(body).find((s) => s.control.tag === tag);
  if (span === undefined) throw new Error(`样本缺少控件 ${tag}`);
  return span;
}
function rangeOf(body: Body) {
  const range = rangeOfContentControl(spanOf(body));
  if (range === undefined) throw new Error('样本控件没有范围');
  return range;
}

describe('文字模板控件同步占位状态与绑定', () => {
  it.each(['<w:text/>', '<w:richText/>', ''])('%s 普通替换清状态，与文字同撤销重做', (type) => {
    const editor = createTextEditor(bodyOf(`<w:p>${control(type)}</w:p>`));
    const before = editor.body;
    editor.tx((t) => void t.replaceText(rangeOf(before), '张三'));
    const after = editor.body;
    expect(textOfContentControl(spanOf(after))).toBe('张三');
    expect(spanOf(after).control.showingPlaceholder).toBe(false);
    expect(spanOf(after).control.dataBound).toBeUndefined();
    editor.undo();
    expect(editor.body).toBe(before);
    editor.redo();
    expect(editor.body).toBe(after);
  });

  it.each(['<w:text/>', '<w:richText/>'])('%s 删除至空仍清状态，格式与空操作保持原状态', (type) => {
    const editor = createTextEditor(bodyOf(`<w:p>${control(type)}</w:p>`));
    const original = spanOf(editor.body).control;
    editor.tx((t) => void t.setRunProps(rangeOf(editor.body), { bold: true }));
    expect(spanOf(editor.body).control).toBe(original);
    const before = editor.body;
    expect(editor.tx((t) => void t.insertText(rangeOf(before).start, ''))).toBeUndefined();
    expect(editor.body).toBe(before);
    editor.tx((t) => void t.deleteRange(rangeOf(before)));
    expect(textOfContentControl(spanOf(editor.body))).toBe('');
    expect(spanOf(editor.body).control.showingPlaceholder).toBe(false);
    expect(spanOf(editor.body).control.dataBound).toBeUndefined();
  });

  it.each(['<w:text/>', '<w:richText/>'])('%s 显式填写相同文字也解除绑定，非占位同样处理', (type) => {
    const xml = `<w:p>${control(type)}</w:p>`.replace('<w:showingPlcHdr/>', '');
    const editor = createTextEditor(bodyOf(xml));
    const before = editor.body;
    editor.tx((t) => void t.fillContentControl(spanOf(before).control.id, '占位文字'));
    expect(textOfContentControl(spanOf(editor.body))).toBe('占位文字');
    expect(spanOf(editor.body).control.dataBound).toBeUndefined();
    editor.undo();
    expect(editor.body).toBe(before);
  });

  it('块级富文本与内层纯文本同同步，未改的邻居保留状态', () => {
    const inner = control('<w:text w:multiLine="1"/>');
    const editor = createTextEditor(
      bodyOf(`${control('', `<w:p>${inner}</w:p>`, 'outer')}<w:p>${control('', undefined, 'other')}</w:p>`),
    );
    const other = spanOf(editor.body, 'other').control;
    editor.tx((t) => {
      const at = t.insertText(rangeOf(editor.body).end, '甲');
      t.insertText(t.splitParagraph(at), '乙');
    });
    for (const tag of ['field', 'outer']) {
      expect(textOfContentControl(spanOf(editor.body, tag))).toBe('占位文字甲\n乙');
      expect(spanOf(editor.body, tag).control.showingPlaceholder).toBe(false);
      expect(spanOf(editor.body, tag).control.dataBound).toBeUndefined();
    }
    expect(spanOf(editor.body, 'other').control).toBe(other);
  });

  it('空块级纯文本首次输入同步，制表位也算显示变化', () => {
    for (const tab of [false, true]) {
      const editor = createTextEditor(bodyOf(control('<w:text/>', '<w:p/>')));
      editor.tx((t) => {
        const at = rangeOf(editor.body).start;
        if (tab) t.insertInline(at, 'tab');
        else t.insertText(at, '甲');
      });
      expect(spanOf(editor.body).control.showingPlaceholder).toBe(false);
      expect(spanOf(editor.body).control.dataBound).toBeUndefined();
      expect(textOfContentControl(spanOf(editor.body))).toBe(tab ? '\t' : '甲');
    }
  });

  it('连续输入合并撤销，校验可见同步状态且拒绝时全部回滚', () => {
    const body = bodyOf(`<w:p>${control('<w:text/>')}</w:p>`);
    const editor = createTextEditor(body, { now: () => 1 });
    const before = editor.body;
    let at = rangeOf(body).end;
    for (const text of ['甲', '乙'])
      editor.tx(
        (t) => {
          at = t.insertText(at, text);
        },
        { origin: 'input' },
      );
    expect(spanOf(editor.body).control.dataBound).toBeUndefined();
    editor.undo();
    expect(editor.body).toBe(before);
    expect(editor.canUndo).toBe(false);
    editor.redo();
    expect(textOfContentControl(spanOf(editor.body))).toBe('占位文字甲乙');
    const rejected = createTextEditor(body, {
      validate(draft) {
        expect(spanOf(draft).control.showingPlaceholder).toBe(false);
        expect(spanOf(draft).control.dataBound).toBeUndefined();
        throw new Error('拒绝提交');
      },
    });
    const rejectedBefore = rejected.body;
    expect(() => rejected.tx((t) => void t.insertText(rangeOf(body).end, '甲'))).toThrow('拒绝提交');
    expect(rejected.body).toBe(rejectedBefore);
    expect(rejected.canUndo).toBe(false);
  });

  it('事务最终文字未变时保留状态，不误报已经填写', () => {
    const editor = createTextEditor(bodyOf(`<w:p>${control('<w:text/>')}</w:p>`));
    const before = spanOf(editor.body).control;
    editor.tx((t) => {
      const start = rangeOf(editor.body).end;
      const end = t.insertText(start, '甲');
      t.deleteRange({ start, end });
    });
    expect(spanOf(editor.body).control).toBe(before);
  });
});
