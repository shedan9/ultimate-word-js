import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import { contentControlSpans, rangeOfContentControl } from './content-control.ts';
import type { Body } from './nodes.ts';
import { walkParagraphs } from './nodes.ts';
import { parseBody } from './parse-body.ts';
import type { DocRange } from './position.ts';
import { createTextEditor, rangeEditable } from './text-transaction.ts';

function parse(xml: string): Body {
  return parseBody(parseXml(`<w:document><w:body>${xml}</w:body></w:document>`), createDiagnosticSink());
}

function control(content: string, lock = 'contentLocked', tag = 'locked'): string {
  return `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:lock w:val="${lock}"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
}

const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
const paragraph = (text: string) => `<w:p>${run(text)}</w:p>`;

function required<T>(value: T | undefined): T {
  if (value === undefined) throw new Error('样本缺少节点');
  return value;
}

function range(body: Body, tag = 'locked'): DocRange {
  const span = contentControlSpans(body).find((s) => s.control.tag === tag);
  const found = span === undefined ? undefined : rangeOfContentControl(span);
  if (found === undefined) throw new Error(`样本缺少控件范围：${tag}`);
  return found;
}

describe('普通编辑遵守内容控件的内容锁', () => {
  it.each(['contentLocked', 'sdtContentLocked'])('%s 拒绝插字、换行、制表位、删除、替换与拆段', (lock) => {
    for (const kind of ['insert', 'tab', 'lineBreak', 'pageBreak', 'delete', 'replace', 'split'] as const) {
      const editor = createTextEditor(parse(`<w:p>${control(run('锁定文字'), lock)}</w:p>`));
      const before = editor.body;
      const r = range(before);
      expect(() =>
        editor.tx((t) => {
          if (kind === 'insert') t.insertText(r.start, '改');
          else if (kind === 'delete') t.deleteRange(r);
          else if (kind === 'replace') t.replaceText(r, '替换');
          else if (kind === 'split') t.splitParagraph(r.start);
          else t.insertInline(r.start, kind);
        }),
      ).toThrow(/锁定/);
      expect(editor.body).toBe(before);
      expect(editor.canUndo).toBe(false);
    }
  });

  it.each(['<w:tab/>', '<w:br/>', '<w:sym w:font="Wingdings" w:char="F0FE"/>'])(
    '锁中的非文字片段也不能删除：%s',
    (fragment) => {
      const editor = createTextEditor(parse(`<w:p>${control(`<w:r>${fragment}</w:r>`)}</w:p>`));
      const before = editor.body;
      expect(() => editor.tx((t) => void t.deleteRange(range(before)))).toThrow(/锁定/);
      expect(editor.body).toBe(before);
    },
  );

  it('内层行内控件继承外层块级内容锁，不能借填外层覆盖锁住的后代', () => {
    const inner = control(run('内层'), 'sdtLocked', 'inner');
    const editor = createTextEditor(parse(control(`<w:p>${inner}</w:p>`)));
    expect(() => editor.tx((t) => void t.insertText(range(editor.body, 'inner').end, '改'))).toThrow(/锁定/);
    const nested = createTextEditor(
      parse(control(`<w:p>${run('外层')}${control(run(''), 'contentLocked', 'inner')}</w:p>`, 'sdtLocked')),
    );
    const before = nested.body;
    const outer = contentControlSpans(before).find((s) => s.control.tag === 'locked');
    expect(() => nested.tx((t) => void t.fillContentControl(required(outer).control.id, '覆盖'))).toThrow(
      /锁定/,
    );
    expect(nested.body).toBe(before);
  });

  it('空块级控件的首次输入与拆段不能绕过锁，空操作仍不记历史', () => {
    const editor = createTextEditor(parse(control('<w:p/>')));
    const before = editor.body;
    const r = range(before);
    expect(editor.tx((t) => void t.insertText(r.start, ''))).toBeUndefined();
    expect(editor.tx((t) => void t.deleteRange(r))).toBeUndefined();
    expect(() => editor.tx((t) => void t.insertText(r.start, '改'))).toThrow(/锁定/);
    expect(() => editor.tx((t) => void t.splitParagraph(r.start))).toThrow(/锁定/);
    expect(editor.body).toBe(before);
    expect(editor.canUndo).toBe(false);
  });

  it('跨 run 删除覆盖中间锁定内容时整次回滚，即使调用方捕获命令错误', () => {
    const editor = createTextEditor(parse(`<w:p>${run('前')}${control(run('中'))}${run('后')}</w:p>`));
    const before = editor.body;
    const p = required([...walkParagraphs(before)][0]);
    const start = { nodeId: required(p.runs[0]).id, contentIndex: 0, offset: 0 };
    const end = { nodeId: required(p.runs[2]).id, contentIndex: 0, offset: 1 };
    expect(() =>
      editor.tx((t) => {
        t.insertText(end, '先改');
        try {
          t.deleteRange({ start, end });
        } catch {}
      }),
    ).toThrow(/全部回滚/);
    expect(editor.body).toBe(before);
    expect(editor.canUndo).toBe(false);
  });

  it('删除在锁定 run 的零偏移处结束时不误拦外面的文字，锁旁编辑可撤销', () => {
    const editor = createTextEditor(parse(`<w:p>${run('外面')}${control(run('锁中'))}</w:p>`));
    const before = editor.body;
    const p = required([...walkParagraphs(before)][0]);
    editor.tx((t) => {
      const start = { nodeId: required(p.runs[0]).id, contentIndex: 0, offset: 0 };
      t.deleteRange({ start, end: range(before).start });
      t.insertText(start, '新');
    });
    expect(required(required([...walkParagraphs(editor.body)][0]).runs[1]).content).toEqual([
      { kind: 'text', text: '锁中' },
    ]);
    editor.undo();
    expect(editor.body).toBe(before);
  });

  it.each(['前', '锁中'])('块级控件与相邻段落合并会改变锁内段落标记，必须拒绝：%s', (first) => {
    const editor = createTextEditor(
      parse(`${paragraph('前')}${control(paragraph('锁中'))}${paragraph('后')}`),
    );
    const before = editor.body;
    const p = required(
      [...walkParagraphs(before)].find((p) =>
        p.runs[0]?.content.some((c) => c.kind === 'text' && c.text === first),
      ),
    );
    expect(() => editor.tx((t) => void t.joinParagraph(p.id))).toThrow(/锁定/);
    expect(editor.body).toBe(before);
  });

  it('跨段删除不能移除中间锁定的空段落', () => {
    const editor = createTextEditor(parse(`${paragraph('前')}${control('<w:p/>')}${paragraph('后')}`));
    const before = editor.body;
    const ps = [...walkParagraphs(before)];
    expect(() =>
      editor.tx(
        (t) =>
          void t.deleteRange({
            start: { nodeId: required(required(ps[0]).runs[0]).id, contentIndex: 0, offset: 0 },
            end: { nodeId: required(required(ps[2]).runs[0]).id, contentIndex: 0, offset: 1 },
          }),
      ),
    ).toThrow(/锁定/);
    expect(editor.body).toBe(before);
  });

  it('锁定范围不交给全部替换；外壳锁的内容仍可改，格式与批注不当成改字', () => {
    const editor = createTextEditor(
      parse(`<w:p>${control(run('锁中'))}${control(run('外壳'), 'sdtLocked', 'shell')}</w:p>`),
    );
    expect(rangeEditable(editor.body, range(editor.body))).toBe(false);
    expect(rangeEditable(editor.body, range(editor.body, 'shell'))).toBe(true);
    const r = range(editor.body);
    editor.tx((t) => {
      t.addComment(r, { text: '批注', author: '作者' });
      t.setRunProps({ start: r.start, end: { ...r.start, offset: 1 } }, { bold: true });
      t.replaceText(range(editor.body, 'shell'), '可改');
    });
    expect(editor.body.comments).toBeDefined();
    expect(() => editor.tx((t) => void t.insertText(range(editor.body).end, '改'))).toThrow(/锁定/);
  });

  it.each(['rows', 'columns'])('删%s不能移除行内锁住的内容，即使操作位置在未锁单元格', (kind) => {
    const cell = (content: string) => `<w:tc><w:p>${content}</w:p></w:tc>`;
    const editor = createTextEditor(
      parse(
        `<w:tbl><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid><w:tr>${cell(run('外面'))}${cell(control(run('锁中')))}</w:tr></w:tbl>`,
      ),
    );
    const before = editor.body;
    const p = required([...walkParagraphs(before)][0]);
    const r = {
      start: { nodeId: required(p.runs[0]).id, contentIndex: 0, offset: 0 },
      end: range(before).end,
    };
    expect(() =>
      editor.tx((t) => {
        if (kind === 'rows') t.deleteRows(r);
        else t.deleteColumns(r);
      }),
    ).toThrow(/锁定/);
    expect(editor.body).toBe(before);
  });

  it.each(['insertRow', 'insertColumn', 'deleteRows', 'deleteColumns', 'mergeCells', 'splitCell'] as const)(
    '包着整张表的内容锁也保护表结构：%s',
    (kind) => {
      const cells =
        kind === 'splitCell'
          ? `<w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>${paragraph('锁中')}</w:tc>`
          : `<w:tc>${paragraph('甲')}</w:tc><w:tc>${paragraph('乙')}</w:tc>`;
      const editor = createTextEditor(
        parse(
          control(
            `<w:tbl><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid><w:tr>${cells}</w:tr></w:tbl>`,
          ),
        ),
      );
      const before = editor.body;
      const r = range(before);
      expect(() =>
        editor.tx((t) => {
          if (kind === 'insertRow') t.insertRow(r.start, 'below');
          else if (kind === 'insertColumn') t.insertColumn(r.start, 'right');
          else if (kind === 'splitCell') t.splitCell(r.start);
          else t[kind](r);
        }),
      ).toThrow(/锁定/);
      expect(editor.body).toBe(before);
    },
  );

  it('合并单元格不能丢掉锁定的空段落', () => {
    const editor = createTextEditor(
      parse(
        `<w:tbl><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc>${paragraph('正文')}</w:tc><w:tc>${control('<w:p/>')}</w:tc></w:tr></w:tbl>`,
      ),
    );
    const before = editor.body;
    const p = required([...walkParagraphs(before)][0]);
    expect(() =>
      editor.tx(
        (t) =>
          void t.mergeCells({
            start: { nodeId: required(p.runs[0]).id, contentIndex: 0, offset: 0 },
            end: range(before).end,
          }),
      ),
    ).toThrow(/锁定/);
    expect(editor.body).toBe(before);
  });

  it('更新域结果也不能覆盖内容锁', () => {
    const editor = createTextEditor(
      parse(
        `<w:p>${control(
          '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
            run('旧目录') +
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
        )}</w:p>`,
      ),
    );
    const before = editor.body;
    expect(() =>
      editor.tx(
        (t) =>
          void t.replaceFieldResult(range(before).start.nodeId, [
            { props: {}, runs: [{ content: [{ kind: 'text', text: '新目录' }] }] },
          ]),
      ),
    ).toThrow(/锁定/);
    expect(editor.body).toBe(before);
  });
});
