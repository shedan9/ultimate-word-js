import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import { contentControlSpans, rangeOfContentControl, textOfContentControl } from './content-control.ts';
import type { Body } from './nodes.ts';
import { walkParagraphs } from './nodes.ts';
import { parseBody } from './parse-body.ts';
import { createTextEditor } from './text-transaction.ts';

const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
const control = (content: string, multiLine = '', tag = 'text') =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:text${multiLine}/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const parse = (xml: string) =>
  parseBody(parseXml(`<w:document><w:body>${xml}</w:body></w:document>`), createDiagnosticSink());

function span(body: Body, tag = 'text') {
  const s = contentControlSpans(body).find((s) => s.control.tag === tag);
  if (s === undefined) throw new Error(`样本缺少控件：${tag}`);
  return s;
}

function range(body: Body, tag = 'text') {
  const r = rangeOfContentControl(span(body, tag));
  if (r === undefined) throw new Error('样本缺少控件范围');
  return r;
}

describe('单行纯文本控件的普通编辑', () => {
  it.each(['', ' w:multiLine="0"', ' w:multiLine="false"', ' w:multiLine="off"'])(
    '关闭 multiLine 时拒绝换行：%s',
    (multiLine) => {
      const editor = createTextEditor(parse(`<w:p>${control(run('原文'), multiLine)}</w:p>`));
      const before = editor.body;
      for (const kind of ['lineBreak', 'pageBreak', 'split', 'replace'] as const) {
        expect(() =>
          editor.tx((t) => {
            const r = range(before);
            if (kind === 'split') t.splitParagraph(r.end);
            else if (kind === 'replace') t.replaceText(r, '甲\r\n乙');
            else t.insertInline(r.end, kind);
          }),
        ).toThrow(/单行纯文本控件不能换行/);
        expect(editor.body).toBe(before);
        expect(editor.canUndo).toBe(false);
      }
    },
  );

  it('错误即使被捕获也回滚前面的文字、格式与选区相关变更，保留重做历史', () => {
    const editor = createTextEditor(parse(`<w:p>${control(run('原文'))}</w:p>`));
    editor.tx((t) => void t.insertText(range(editor.body).end, '改'));
    editor.undo();
    const before = editor.body;
    expect(() =>
      editor.tx((t) => {
        const r = range(before);
        t.setRunProps(r, { bold: true });
        try {
          t.replaceText(r, '甲\n乙');
        } catch {}
      }),
    ).toThrow(/全部回滚/);
    expect(editor.body).toBe(before);
    expect(editor.canRedo).toBe(true);
    editor.redo();
    expect(textOfContentControl(span(editor.body))).toBe('原文改');
  });

  it.each(['inline', 'block', 'empty'] as const)('嵌套多行控件不能绕过 %s 单行父控件', (scope) => {
    const inner = control(run('原文'), ' w:multiLine="1"', 'inner');
    const content = scope === 'empty' ? '<w:p/>' : scope === 'block' ? `<w:p>${inner}</w:p>` : inner;
    const xml = control(content);
    const editor = createTextEditor(parse(scope === 'inline' ? `<w:p>${xml}</w:p>` : xml));
    const before = editor.body;
    expect(() => editor.tx((t) => void t.insertInline(range(before).end, 'lineBreak'))).toThrow(
      /单行纯文本控件不能换行/,
    );
    expect(() => editor.tx((t) => void t.splitParagraph(range(before).end))).toThrow(
      /单行纯文本控件不能换行/,
    );
    if (scope !== 'empty')
      expect(() =>
        editor.tx((t) => void t.fillContentControl(span(before, 'inner').control.id, '甲\n乙')),
      ).toThrow(/单行纯文本控件不能换行/);
    expect(editor.body).toBe(before);
    expect(editor.canUndo).toBe(false);
  });

  it('多行父控件不能取消单行子控件的约束，父控件自己的文字仍可换行', () => {
    const editor = createTextEditor(
      parse(control(`<w:p>${run('前')}${control(run('原文'))}</w:p>`, ' w:multiLine="1"', 'outer')),
    );
    expect(() => editor.tx((t) => void t.insertInline(range(editor.body).end, 'lineBreak'))).toThrow(
      /单行纯文本控件不能换行/,
    );
    editor.tx((t) => void t.insertInline(range(editor.body, 'outer').start, 'lineBreak'));
    expect(textOfContentControl(span(editor.body, 'outer'))).toBe('\n前原文');
  });

  it.each(['1', 'true', 'on'])('multiLine=%s 允许软换行和拆段，并随撤销重做', (value) => {
    const editor = createTextEditor(parse(control(`<w:p>${run('原文')}</w:p>`, ` w:multiLine="${value}"`)));
    const before = editor.body;
    editor.tx((t) => {
      const at = t.insertInline(range(before).end, 'lineBreak');
      t.insertText(t.splitParagraph(at), '后');
    });
    expect(textOfContentControl(span(editor.body))).toBe('原文\n\n后');
    expect([...walkParagraphs(editor.body)]).toHaveLength(2);
    editor.undo();
    expect(editor.body).toBe(before);
    editor.redo();
    expect(textOfContentControl(span(editor.body))).toBe('原文\n\n后');
  });

  it('单行控件仍允许普通文字、制表位、格式与填写，控件外仍可拆段', () => {
    const editor = createTextEditor(parse(`<w:p>${run('前')}${control(run('原文'))}${run('后')}</w:p>`));
    editor.tx((t) => {
      t.insertText(range(editor.body).end, '改');
      t.insertInline(range(editor.body).start, 'tab');
    });
    expect(textOfContentControl(span(editor.body))).toBe('\t原文改');
    editor.tx((t) => void t.setRunProps(range(editor.body), { bold: true }));
    editor.tx((t) => void t.fillContentControl(span(editor.body).control.id, '甲\r\n乙\r丙\n丁'));
    expect(textOfContentControl(span(editor.body))).toBe('甲 乙 丙 丁');
    const p = [...walkParagraphs(editor.body)][0];
    const last = p?.runs.at(-1);
    if (last === undefined) throw new Error('样本缺少外部文字');
    editor.tx((t) => void t.splitParagraph({ nodeId: last.id, contentIndex: 0, offset: 1 }));
    expect([...walkParagraphs(editor.body)]).toHaveLength(2);
  });

  it('导入时已有的多行内容仍可删除换行，不因无关编辑被拒绝', () => {
    const editor = createTextEditor(
      parse(control(`<w:p><w:r><w:t>甲</w:t><w:br/><w:t>乙</w:t></w:r></w:p>`)),
    );
    const r = range(editor.body);
    editor.tx((t) => void t.insertText(r.end, '改'));
    editor.tx(
      (t) =>
        void t.deleteRange({
          start: { ...r.start, contentIndex: 1, offset: 0 },
          end: { ...r.start, contentIndex: 1, offset: 1 },
        }),
    );
    expect(textOfContentControl(span(editor.body))).toBe('甲乙改');
  });

  it.each(['paragraphs', 'break'] as const)('直接装配域结果也不能写入 %s 换行', (kind) => {
    const field =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      run('旧结果') +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const editor = createTextEditor(parse(control(`<w:p>${field}</w:p>`)));
    const before = editor.body;
    const draft = { props: {}, runs: [{ content: [{ kind: 'text' as const, text: '新结果' }] }] };
    expect(() =>
      editor.tx(
        (t) =>
          void t.replaceFieldResult(
            range(before).start.nodeId,
            kind === 'paragraphs'
              ? [draft, draft]
              : [{ props: {}, runs: [{ content: [{ kind: 'break', breakType: 'line' }] }] }],
          ),
      ),
    ).toThrow(/单行纯文本控件不能换行/);
    expect(editor.body).toBe(before);
    editor.tx((t) => void t.replaceFieldResult(range(before).start.nodeId, [draft]));
    expect(textOfContentControl(span(editor.body))).toBe('新结果');
  });
});
