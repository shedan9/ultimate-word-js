import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import { contentControlSpans, rangeOfContentControl } from './content-control.ts';
import { parseBody } from './parse-body.ts';
import type { TextTransaction } from './text-transaction.ts';
import { createTextEditor } from './text-transaction.ts';

function editorOf(type: 'date' | 'checkbox', block: boolean, revision = '', bound = true) {
  const pr =
    type === 'date'
      ? '<w:date w:fullDate="2026-10-13T00:00:00Z"><w:dateFormat w:val="yyyy"/></w:date>'
      : '<w14:checkbox/>';
  const run = `<w:r><w:t>${type === 'date' ? '2026' : '☐'}</w:t></w:r>`;
  const pPr = revision ? `<w:pPr><w:rPr><w:${revision} w:id="7" w:author="甲"/></w:rPr></w:pPr>` : '';
  const control = (content: string) =>
    `<w:sdt><w:sdtPr><w:tag w:val="field"/>${bound ? '<w:dataBinding w:xpath="/field"/>' : ''}${pr}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
  const xml = block ? control(`<w:p>${pPr}${run}</w:p>`) : `<w:p>${pPr}${control(run)}</w:p>`;
  return createTextEditor(
    parseBody(parseXml(`<w:document><w:body>${xml}</w:body></w:document>`), createDiagnosticSink()),
  );
}

function fill(t: TextTransaction, type: 'date' | 'checkbox', id: string, changed: boolean) {
  return type === 'date'
    ? t.setContentControlDate(id, changed ? '2026-10-14' : '2026-10-13')
    : t.setContentControlChecked(id, changed);
}

describe.each(['date', 'checkbox'] as const)('%s 填写检查段落标记修订', (type) => {
  it.each([false, true])('普通内容仍可填写并撤销重做：块级 %s', (block) => {
    const editor = editorOf(type, block);
    const before = editor.body;
    const span = contentControlSpans(before)[0];
    if (span === undefined) throw new Error('样本缺少控件');
    editor.tx((t) => void fill(t, type, span.control.id, true));
    const after = editor.body;
    expect(contentControlSpans(after)[0]?.control.dataBound).toBeUndefined();
    editor.undo();
    expect(editor.body).toBe(before);
    editor.redo();
    expect(editor.body).toBe(after);
  });

  for (const block of [false, true])
    for (const revision of ['ins', 'del'])
      it.each([false, true])(`段落标记 ${revision} 拒绝同值与改值：块级 ${block}，改值 %s`, (changed) => {
        // 同值提交也可能解除绑定；年份格式还覆盖只改日期属性、不改显示文字的路径。
        const editor = editorOf(type, block, revision);
        const before = editor.body;
        const span = contentControlSpans(before)[0];
        if (span === undefined) throw new Error('样本缺少控件');
        expect(span.paragraphs[0]?.markRevision).toBeDefined();
        expect(span.runs.every((r) => r.revision === undefined)).toBe(true);
        expect(() => editor.tx((t) => void fill(t, type, span.control.id, changed))).toThrow(/无修订/);
        expect(editor.body).toBe(before);
        expect(editor.canUndo).toBe(false);
      });

  it('捕获错误仍回滚已有修改并保留重做历史；无绑定的同值也校验', () => {
    const editor = editorOf(type, true, 'del', false);
    const before = editor.body;
    const span = contentControlSpans(before)[0];
    const range = span === undefined ? undefined : rangeOfContentControl(span);
    if (span === undefined || range === undefined) throw new Error('样本缺少控件范围');
    editor.tx((t) => void t.setRunProps(range, { bold: true }));
    editor.undo();
    expect(() =>
      editor.tx((t) => {
        t.setRunProps(range, { italic: true });
        try {
          fill(t, type, span.control.id, false);
        } catch {}
      }),
    ).toThrow(/全部回滚/);
    expect(editor.body).toBe(before);
    expect(editor.canUndo).toBe(false);
    expect(editor.canRedo).toBe(true);
    editor.redo();
    expect(contentControlSpans(editor.body)[0]?.runs[0]?.props).toEqual({ bold: true });
  });
});
