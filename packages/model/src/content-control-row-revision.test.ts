import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import { contentControlSpans, rangeOfContentControl } from './content-control.ts';
import { parseBody } from './parse-body.ts';
import type { TextTransaction } from './text-transaction.ts';
import { createTextEditor } from './text-transaction.ts';

function editorOf(
  type: 'date' | 'checkbox',
  block: boolean,
  revision: string,
  nested = false,
  empty = false,
) {
  const pr =
    type === 'date'
      ? '<w:date w:fullDate="2026-10-13T00:00:00Z"><w:dateFormat w:val="yyyy"/></w:date>'
      : '<w14:checkbox/>';
  const control = (content: string) =>
    `<w:sdt><w:sdtPr><w:tag w:val="field"/><w:dataBinding w:xpath="/field"/>${pr}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
  const run = empty ? '' : `<w:r><w:t>${type === 'date' ? '2026' : '☐'}</w:t></w:r>`;
  const content = block ? control(`<w:p>${run}</w:p>`) : `<w:p>${control(run)}</w:p>`;
  const table = (blocks: string, mark = '') =>
    `<w:tbl><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr>${mark ? `<w:trPr><w:${mark} w:id="7" w:author="甲"/></w:trPr>` : ''}<w:tc>${blocks}</w:tc></w:tr></w:tbl>`;
  // 内层表格的行没有修订，必须沿祖先找；邻表修订不能误伤目标。
  const xml = table(nested ? `${table(content)}<w:p/>` : content, revision) + table('<w:p/>', 'del');
  return createTextEditor(
    parseBody(parseXml(`<w:document><w:body>${xml}<w:p/></w:body></w:document>`), createDiagnosticSink()),
  );
}

function fill(t: TextTransaction, type: 'date' | 'checkbox', id: string, changed: boolean) {
  return type === 'date'
    ? t.setContentControlDate(id, changed ? '2026-10-14' : '2026-10-13')
    : t.setContentControlChecked(id, changed);
}

describe.each(['date', 'checkbox'] as const)('%s 填写检查表格行修订', (type) => {
  for (const block of [false, true])
    for (const nested of [false, true]) {
      it(`邻表修订不影响普通行填写与历史：块级 ${block}，嵌套 ${nested}`, () => {
        const editor = editorOf(type, block, '', nested);
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

      for (const revision of ['ins', 'del'])
        it.each([false, true])(
          `${revision} 行拒绝同值解除绑定与改值：块级 ${block}，嵌套 ${nested}，改值 %s`,
          (changed) => {
            const editor = editorOf(type, block, revision, nested);
            const before = editor.body;
            const span = contentControlSpans(before)[0];
            if (span === undefined) throw new Error('样本缺少控件');
            expect(span.paragraphs[0]?.markRevision).toBeUndefined();
            expect(span.runs.every((r) => r.revision === undefined)).toBe(true);
            expect(() => editor.tx((t) => void fill(t, type, span.control.id, changed))).toThrow(/无修订/);
            expect(editor.body).toBe(before);
            expect(editor.canUndo).toBe(false);
          },
        );
    }

  it('捕获行修订错误仍回滚整个事务并保留重做', () => {
    const editor = editorOf(type, true, 'del', true);
    const before = editor.body;
    const span = contentControlSpans(before)[0];
    const range = span === undefined ? undefined : rangeOfContentControl(span);
    if (span === undefined || range === undefined) throw new Error('样本缺少范围');
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
    expect(editor.canRedo).toBe(true);
    editor.redo();
    expect(contentControlSpans(editor.body)[0]?.runs[0]?.props).toEqual({ bold: true });
  });
});

it.each(['ins', 'del'])('空块级日期也不能在祖先 %s 行修订中创建文字槽', (revision) => {
  const editor = editorOf('date', true, revision, true, true);
  const before = editor.body;
  const span = contentControlSpans(before)[0];
  if (span === undefined) throw new Error('样本缺少控件');
  expect(span.runs).toEqual([]);
  expect(() => editor.tx((t) => void t.setContentControlDate(span.control.id, '2026-10-14'))).toThrow(
    /无修订/,
  );
  expect(editor.body).toBe(before);
  expect(editor.canUndo).toBe(false);
});
