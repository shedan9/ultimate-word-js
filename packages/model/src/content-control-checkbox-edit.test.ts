import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import { contentControlSpans, rangeOfContentControl, textOfContentControl } from './content-control.ts';
import type { Body } from './nodes.ts';
import { walkParagraphs } from './nodes.ts';
import { parseBody } from './parse-body.ts';
import { createTextEditor, rangeEditable } from './text-transaction.ts';

const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
const control = (content: string, pr = '<w14:checkbox/>', tag = 'check') =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/>${pr}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const editorOf = (xml = `<w:p>${run('前')}${control(run('☐'))}${run('后')}</w:p>`) =>
  createTextEditor(
    parseBody(parseXml(`<w:document><w:body>${xml}</w:body></w:document>`), createDiagnosticSink()),
  );

function span(body: Body, tag = 'check') {
  const s = contentControlSpans(body).find((s) => s.control.tag === tag);
  if (s === undefined) throw new Error(`样本缺少控件：${tag}`);
  return s;
}

function range(body: Body, tag = 'check') {
  const r = rangeOfContentControl(span(body, tag));
  if (r === undefined) throw new Error('样本缺少控件范围');
  return r;
}

describe('复选框只通过勾选接口修改内容', () => {
  it.each(['insert', 'delete', 'replace', 'tab', 'lineBreak', 'pageBreak', 'split'] as const)(
    '拒绝普通 %s 命令，保留显示和状态',
    (kind) => {
      const editor = editorOf();
      const before = editor.body;
      const r = range(before);
      expect(rangeEditable(before, r)).toBe(false);
      expect(() =>
        editor.tx((t) => {
          if (kind === 'insert') t.insertText(r.end, '改');
          else if (kind === 'delete') t.deleteRange(r);
          else if (kind === 'replace') t.replaceText(r, '☒');
          else if (kind === 'split') t.splitParagraph(r.start);
          else t.insertInline(r.start, kind);
        }),
      ).toThrow(/复选框只能/);
      expect(editor.body).toBe(before);
      expect(editor.canUndo).toBe(false);
    },
  );

  it('专用切换后同事务输入仍拒绝，捕获异常也回滚全部修改并保留重做历史', () => {
    const editor = editorOf();
    const before = editor.body;
    const id = span(before).control.id;
    editor.tx((t) => void t.setContentControlChecked(id, true));
    expect(textOfContentControl(span(editor.body))).toBe('☒');
    expect(() => editor.tx((t) => void t.deleteRange(range(editor.body)))).toThrow(/复选框只能/);
    editor.undo();
    expect(editor.body).toBe(before);
    expect(() =>
      editor.tx((t) => {
        const selected = t.setContentControlChecked(id, true);
        try {
          t.insertText(selected.end, '改');
        } catch {}
      }),
    ).toThrow(/全部回滚/);
    expect(editor.body).toBe(before);
    expect(editor.canRedo).toBe(true);
    editor.redo();
    expect(span(editor.body).control.checkbox?.checked).toBe(true);
    editor.tx((t) => void t.setContentControlChecked(id, false));
    expect(textOfContentControl(span(editor.body))).toBe('☐');
  });

  it('空操作、格式与批注允许，边界外的文字仍可编辑', () => {
    const editor = editorOf();
    const before = editor.body;
    const r = range(before);
    expect(editor.tx((t) => void t.insertText(r.start, ''))).toBeUndefined();
    expect(editor.tx((t) => void t.deleteRange({ start: r.start, end: r.start }))).toBeUndefined();
    editor.tx((t) => {
      t.setRunProps(r, { bold: true });
      t.addComment(r, { text: '说明', author: '作者' });
    });
    expect(span(editor.body).control).toBe(span(before).control);
    const p = [...walkParagraphs(editor.body)][0];
    const first = p?.runs[0];
    if (first === undefined) throw new Error('样本缺少前文');
    editor.tx(
      (t) =>
        void t.deleteRange({
          start: { nodeId: first.id, contentIndex: 0, offset: 0 },
          end: range(editor.body).start,
        }),
    );
    expect(textOfContentControl(span(editor.body))).toBe('☐');
    editor.tx((t) => void t.setContentControlChecked(span(editor.body).control.id, true));
    expect(textOfContentControl(span(editor.body))).toBe('☒');
  });

  it.each(['<w:p/>', `<w:p>${run('☐')}</w:p>`])('块级复选框保护空段及段落标记：%s', (content) => {
    const editor = editorOf(`<w:p>${run('前')}</w:p>${control(content)}`);
    const before = editor.body;
    expect(() => editor.tx((t) => void t.insertText(range(before).end, '改'))).toThrow(/复选框只能/);
    const first = [...walkParagraphs(before)][0];
    if (first === undefined) throw new Error('样本缺少前段');
    expect(() => editor.tx((t) => void t.joinParagraph(first.id))).toThrow(/复选框只能/);
    expect(editor.body).toBe(before);
  });

  it.each(['', '☐'])('覆盖外层不能移除复选框后代，即使内容为空：%s', (text) => {
    const editor = editorOf(control(`<w:p>${run('前')}${control(run(text))}</w:p>`, '', 'outer'));
    const before = editor.body;
    expect(() =>
      editor.tx((t) => void t.fillContentControl(span(before, 'outer').control.id, '覆盖')),
    ).toThrow(/复选框只能/);
    expect(editor.body).toBe(before);
  });

  it.each(['text', 'date', 'dropDownList', 'checkbox'] as const)(
    '填写内层 %s 不能绕过复选框父控件',
    (type) => {
      const pr =
        type === 'checkbox'
          ? '<w14:checkbox/>'
          : type === 'date'
            ? '<w:date><w:dateFormat w:val="yyyy-MM-dd"/></w:date>'
            : type === 'dropDownList'
              ? '<w:dropDownList><w:listItem w:displayText="新" w:value="1"/></w:dropDownList>'
              : '<w:text/>';
      const inner = control(run(type === 'checkbox' ? '☐' : '原'), pr, 'inner');
      const editor = editorOf(control(`<w:p>${inner}</w:p>`));
      const before = editor.body;
      const id = span(before, 'inner').control.id;
      expect(() =>
        editor.tx((t) => {
          if (type === 'checkbox') t.setContentControlChecked(id, true);
          else if (type === 'date') t.setContentControlDate(id, '2026-10-16');
          else if (type === 'dropDownList') t.selectContentControlItem(id, 0);
          else t.fillContentControl(id, '新');
        }),
      ).toThrow(/复选框只能/);
      expect(editor.body).toBe(before);
    },
  );

  it.each(['rows', 'columns'] as const)('%s 结构删除不能移除复选框内容', (kind) => {
    const editor = editorOf(
      `<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p>${control(run('☐'))}</w:p></w:tc></w:tr></w:tbl>`,
    );
    const before = editor.body;
    expect(() =>
      editor.tx((t) => {
        if (kind === 'rows') t.deleteRows(range(before));
        else t.deleteColumns(range(before));
      }),
    ).toThrow(/复选框只能/);
    expect(editor.body).toBe(before);
  });

  it('更新域结果不能移除复选框', () => {
    const editor = editorOf(
      '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        control(run('☐')) +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
    );
    const before = editor.body;
    const start = [...walkParagraphs(before)][0]?.runs[0]?.id ?? '';
    expect(() =>
      editor.tx(
        (t) =>
          void t.replaceFieldResult(start, [
            { props: {}, runs: [{ content: [{ kind: 'text', text: '新' }] }] },
          ]),
      ),
    ).toThrow(/复选框只能/);
    expect(editor.body).toBe(before);
  });
});
