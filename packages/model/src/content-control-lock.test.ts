import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import { contentControlSpans, rangeOfContentControl, textOfContentControl } from './content-control.ts';
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

const dropdown = (content: string, extra = '') =>
  `<w:sdt><w:sdtPr><w:tag w:val="locked"/>${extra}<w:dropDownList w:lastValue="2"><w:listItem w:displayText="紧急" w:value="1"/><w:listItem w:displayText="一般" w:value="2"/></w:dropDownList></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;

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

describe('下拉框只通过填写接口选择选项', () => {
  it.each(['insert', 'tab', 'lineBreak', 'pageBreak', 'delete', 'replace', 'split'] as const)(
    '普通 %s 命令拒绝改写下拉框，即使替换成有效选项',
    (kind) => {
      const editor = createTextEditor(parse(`<w:p>${dropdown(run('一般'))}</w:p>`));
      const before = editor.body;
      const r = range(before);
      expect(rangeEditable(before, r)).toBe(false);
      expect(() =>
        editor.tx((t) => {
          if (kind === 'insert') t.insertText(r.end, '改');
          else if (kind === 'delete') t.deleteRange(r);
          else if (kind === 'replace') t.replaceText(r, '紧急');
          else if (kind === 'split') t.splitParagraph(r.start);
          else t.insertInline(r.start, kind);
        }),
      ).toThrow(/下拉框只能/);
      expect(editor.body).toBe(before);
      expect(editor.canUndo).toBe(false);
    },
  );

  it('填写只临时放行目标，同事务再输入即失败并回滚填写', () => {
    const editor = createTextEditor(parse(`<w:p>${dropdown(run('一般'))}</w:p>`));
    const before = editor.body;
    const id = required(contentControlSpans(before)[0]).control.id;
    expect(() =>
      editor.tx((t) => {
        const filled = t.fillContentControl(id, '1');
        try {
          t.insertText(filled.end, '改');
        } catch {}
      }),
    ).toThrow(/全部回滚/);
    expect(editor.body).toBe(before);
    editor.tx((t) => void t.fillContentControl(id, '1'));
    expect(required(contentControlSpans(editor.body)[0]).control.lastValue).toBe('1');
    editor.undo();
    expect(editor.body).toBe(before);
    editor.redo();
    expect(() => editor.tx((t) => void t.insertText(range(editor.body).end, '改'))).toThrow(/下拉框只能/);
    expect(required(contentControlSpans(editor.body)[0]).control.lastValue).toBe('1');
  });

  it('占位文字可填写，空操作、格式与批注不误拦，也不改变元数据', () => {
    const editor = createTextEditor(parse(`<w:p>${dropdown(run('选择'), '<w:showingPlcHdr/>')}</w:p>`));
    const before = editor.body;
    const r = range(before);
    expect(editor.tx((t) => void t.insertText(r.start, ''))).toBeUndefined();
    expect(editor.tx((t) => void t.deleteRange({ start: r.start, end: r.start }))).toBeUndefined();
    editor.tx((t) => {
      const formatted = t.setRunProps({ start: r.start, end: { ...r.start, offset: 1 } }, { bold: true });
      t.addComment(formatted, { text: '说明', author: '作者' });
    });
    expect(required(contentControlSpans(editor.body)[0]).control).toBe(
      required(contentControlSpans(before)[0]).control,
    );
    editor.tx(
      (t) => void t.fillContentControl(required(contentControlSpans(editor.body)[0]).control.id, '1'),
    );
    expect(textOfContentControl(required(contentControlSpans(editor.body)[0]))).toBe('紧急');
  });

  it('空块级下拉框首次输入拒绝，填写可创建文字槽', () => {
    const editor = createTextEditor(parse(dropdown('<w:p/>')));
    const before = editor.body;
    expect(() => editor.tx((t) => void t.insertText(range(before).start, '改'))).toThrow(/下拉框只能/);
    expect(editor.body).toBe(before);
    editor.tx((t) => void t.fillContentControl(required(contentControlSpans(before)[0]).control.id, '1'));
    expect(textOfContentControl(required(contentControlSpans(editor.body)[0]))).toBe('紧急');
  });

  it('填内层不能改外层下拉框，填外层不能覆盖下拉框后代，包括空控件', () => {
    const inner = control(run('一般'), '', 'inner');
    const editor = createTextEditor(parse(dropdown(`<w:p>${inner}</w:p>`)));
    const before = editor.body;
    const id = required(contentControlSpans(before).find((s) => s.control.tag === 'inner')).control.id;
    expect(() => editor.tx((t) => void t.fillContentControl(id, '改'))).toThrow(/下拉框只能/);
    expect(() => editor.tx((t) => void t.insertText(range(before, 'inner').end, '改'))).toThrow(/下拉框只能/);
    for (const text of ['一般', '']) {
      const outer = createTextEditor(
        parse(control(`<w:p>${run('前')}${dropdown(run(text))}</w:p>`, '', 'outer')),
      );
      const body = outer.body;
      const outerId = required(contentControlSpans(body).find((s) => s.control.tag === 'outer')).control.id;
      expect(() => outer.tx((t) => void t.fillContentControl(outerId, '覆盖'))).toThrow(/下拉框只能/);
      expect(outer.body).toBe(body);
    }
  });

  it('下拉框边界外的文字仍可删改，失败保留原有重做历史', () => {
    const editor = createTextEditor(parse(`<w:p>${run('前')}${dropdown(run('一般'))}</w:p>`));
    const p = required([...walkParagraphs(editor.body)][0]);
    const at = { nodeId: required(p.runs[0]).id, contentIndex: 0, offset: 0 };
    editor.tx((t) => void t.deleteRange({ start: at, end: range(editor.body).start }));
    editor.undo();
    const before = editor.body;
    expect(() => editor.tx((t) => void t.deleteRange(range(before)))).toThrow(/下拉框只能/);
    expect(editor.body).toBe(before);
    expect(editor.canRedo).toBe(true);
    editor.redo();
    expect(textOfContentControl(required(contentControlSpans(editor.body)[0]))).toBe('一般');
  });

  it.each(['<w:tab/>', '<w:br/>', '<w:sym w:font="Wingdings" w:char="F0FE"/>'])(
    '非文字片段删除也不能改写下拉框：%s',
    (fragment) => {
      const editor = createTextEditor(parse(`<w:p>${dropdown(`<w:r>${fragment}</w:r>`)}</w:p>`));
      const before = editor.body;
      expect(() => editor.tx((t) => void t.deleteRange(range(before)))).toThrow(/下拉框只能/);
      expect(editor.body).toBe(before);
    },
  );

  it.each(['checkbox', 'date'] as const)('填写嵌套的 %s 也不能绕过外层下拉框', (type) => {
    const pr =
      type === 'checkbox'
        ? '<w14:checkbox/>'
        : '<w:date w:fullDate="2026-10-13T00:00:00Z"><w:dateFormat w:val="yyyy年M月"/></w:date>';
    const inner = `<w:sdt><w:sdtPr><w:tag w:val="inner"/>${pr}</w:sdtPr><w:sdtContent>${run(type === 'checkbox' ? '☐' : '2026年10月')}</w:sdtContent></w:sdt>`;
    const editor = createTextEditor(parse(dropdown(`<w:p>${inner}</w:p>`)));
    const before = editor.body;
    const id = required(contentControlSpans(before).find((s) => s.control.tag === 'inner')).control.id;
    expect(() =>
      editor.tx((t) => {
        if (type === 'checkbox') t.setContentControlChecked(id, true);
        else t.setContentControlDate(id, '2026-10-14');
      }),
    ).toThrow(/下拉框只能/);
    expect(editor.body).toBe(before);
  });

  it('更新域结果不能绕过下拉框限制', () => {
    const editor = createTextEditor(
      parse(
        `<w:p>${dropdown(
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
    ).toThrow(/下拉框只能/);
    expect(editor.body).toBe(before);
  });

  it.each(['rows', 'columns', 'join'] as const)('%s 结构命令不能移除下拉框内容或段落标记', (kind) => {
    const body =
      kind === 'join'
        ? parse(`${paragraph('前')}${dropdown(paragraph('一般'))}`)
        : parse(
            `<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p>${dropdown(run('一般'))}</w:p></w:tc></w:tr></w:tbl>`,
          );
    const editor = createTextEditor(body);
    const before = editor.body;
    expect(() =>
      editor.tx((t) => {
        if (kind === 'join') t.joinParagraph(required([...walkParagraphs(before)][0]).id);
        else if (kind === 'rows') t.deleteRows(range(before));
        else t.deleteColumns(range(before));
      }),
    ).toThrow(/下拉框只能/);
    expect(editor.body).toBe(before);
  });
});

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

describe('内容控件外壳锁', () => {
  it('内部删光与重新填写仍保留外壳，格式拆 run、拆段和撤销重做不误拦', () => {
    const editor = createTextEditor(parse(`<w:p>${control(run('原文'), 'sdtLocked')}</w:p>`));
    const before = editor.body;
    const id = required(contentControlSpans(before)[0]).control.id;
    editor.tx((t) => void t.deleteRange(range(before)));
    expect(textOfContentControl(required(contentControlSpans(editor.body)[0]))).toBe('');
    editor.tx((t) => void t.fillContentControl(id, '新文字'));
    editor.tx(
      (t) =>
        void t.setRunProps(
          { start: range(editor.body).start, end: { ...range(editor.body).start, offset: 1 } },
          { bold: true },
        ),
    );
    editor.tx((t) => void t.splitParagraph(range(editor.body).end));
    expect(textOfContentControl(required(contentControlSpans(editor.body)[0]))).toBe('新文字\n');
    editor.undo();
    editor.redo();
    expect(required(contentControlSpans(editor.body)[0]).control).toMatchObject({ id, lock: 'sdtLocked' });
  });

  it.each(['', '后代'])('填外层不能清理掉内层外壳，即使内层原本没有字：%s', (text) => {
    const editor = createTextEditor(
      parse(control(`<w:p>${run('前')}${control(run(text), 'sdtLocked', 'inner')}</w:p>`, '', 'outer')),
    );
    const before = editor.body;
    const id = required(contentControlSpans(before).find((s) => s.control.tag === 'outer')).control.id;
    expect(() => editor.tx((t) => void t.fillContentControl(id, '覆盖'))).toThrow(/锁定了外壳/);
    expect(editor.body).toBe(before);
    expect(editor.canUndo).toBe(false);
  });

  it('外壳锁沿嵌套成员识别，不能通过删除无锁后代间接删掉锁定祖先', () => {
    const inner = control(control(run('后代'), '', 'child'), 'sdtLocked', 'inner');
    const editor = createTextEditor(parse(control(`<w:p>${run('前')}${inner}</w:p>`, '', 'outer')));
    const before = editor.body;
    const id = required(contentControlSpans(before).find((s) => s.control.tag === 'outer')).control.id;
    expect(() => editor.tx((t) => void t.fillContentControl(id, '覆盖'))).toThrow(/锁定了外壳/);
    expect(editor.body).toBe(before);
  });

  it.each(['join', 'delete'] as const)('跨块边界的 %s 不能移除块级外壳，包含空段也保护', (kind) => {
    for (const content of [
      '<w:p/>',
      `<w:p>${run('里面')}</w:p>`,
      `<w:p>${control(run('后代'), '', 'child')}</w:p>`,
    ]) {
      const editor = createTextEditor(
        parse(`${paragraph('前')}${control(content, 'sdtLocked')}${paragraph('后')}`),
      );
      const before = editor.body;
      const ps = [...walkParagraphs(before)];
      expect(() =>
        editor.tx((t) => {
          if (kind === 'join') t.joinParagraph(required(ps[0]).id);
          else
            t.deleteRange({
              start: { nodeId: required(required(ps[0]).runs[0]).id, contentIndex: 0, offset: 0 },
              end: { nodeId: required(required(ps[2]).runs[0]).id, contentIndex: 0, offset: 1 },
            });
        }),
      ).toThrow(/锁定了外壳/);
      expect(editor.body).toBe(before);
    }
  });

  it('同一块级外壳内合段允许；移除部分成员不等于删除整个控件', () => {
    const editor = createTextEditor(parse(control(`${paragraph('甲')}${paragraph('乙')}`, 'sdtLocked')));
    const p = required([...walkParagraphs(editor.body)][0]);
    editor.tx((t) => void t.joinParagraph(p.id));
    expect(textOfContentControl(required(contentControlSpans(editor.body)[0]))).toBe('甲乙');
  });

  it('锁住父外壳不会锁住无锁子外壳，覆盖父内容仍可移除子控件', () => {
    const editor = createTextEditor(
      parse(control(`<w:p>${run('前')}${control(run('后代'), '', 'child')}</w:p>`, 'sdtLocked')),
    );
    const id = required(contentControlSpans(editor.body).find((s) => s.control.tag === 'locked')).control.id;
    editor.tx((t) => void t.fillContentControl(id, '覆盖'));
    expect(contentControlSpans(editor.body).map((s) => [s.control.id, textOfContentControl(s)])).toEqual([
      [id, '覆盖'],
    ]);
  });

  it.each(['rows', 'columns', 'merge'] as const)('表格的 %s 不能丢掉单元格里的锁定空外壳', (kind) => {
    const editor = createTextEditor(
      parse(
        `<w:tbl><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc>${paragraph('正文')}</w:tc><w:tc>${control('<w:p/>', 'sdtLocked')}</w:tc></w:tr></w:tbl>${paragraph('后')}`,
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
        else if (kind === 'columns') t.deleteColumns(r);
        else t.mergeCells(r);
      }),
    ).toThrow(/锁定了外壳/);
    expect(editor.body).toBe(before);
  });

  it('整表外壳允许插行删行，最后一行删除才拒绝', () => {
    const editor = createTextEditor(
      parse(
        control(
          `<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc>${paragraph('甲')}</w:tc></w:tr></w:tbl>`,
          'sdtLocked',
        ),
      ),
    );
    let inserted = range(editor.body).start;
    editor.tx((t) => {
      inserted = t.insertRow(range(editor.body).start, 'below');
    });
    editor.tx((t) => void t.deleteRows({ start: inserted, end: inserted }));
    const before = editor.body;
    expect(() => editor.tx((t) => void t.deleteRows(range(before)))).toThrow(/锁定了外壳/);
    expect(editor.body).toBe(before);
    expect(contentControlSpans(before)).toHaveLength(1);
  });

  it('更新域结果不能把结果区里的锁定外壳一并替换', () => {
    const editor = createTextEditor(
      parse(
        `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${control(run('旧目录'), 'sdtLocked')}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`,
      ),
    );
    const before = editor.body;
    const first = required(required([...walkParagraphs(before)][0]).runs[0]);
    expect(() =>
      editor.tx(
        (t) =>
          void t.replaceFieldResult(first.id, [
            { props: {}, runs: [{ content: [{ kind: 'text', text: '新目录' }] }] },
          ]),
      ),
    ).toThrow(/锁定了外壳/);
    expect(editor.body).toBe(before);
  });

  it('捕获外壳删除失败仍回滚全部命令，保留旧撤销与重做栈', () => {
    const editor = createTextEditor(parse(`${paragraph('前')}${control(paragraph('壳内'), 'sdtLocked')}`));
    const p = required([...walkParagraphs(editor.body)][0]);
    const at = { nodeId: required(p.runs[0]).id, contentIndex: 0, offset: 1 };
    editor.tx((t) => void t.insertText(at, '旧改'));
    editor.undo();
    const before = editor.body;
    expect(() =>
      editor.tx((t) => {
        t.insertText(at, '先改');
        try {
          t.joinParagraph(p.id);
        } catch {}
      }),
    ).toThrow(/全部回滚/);
    expect(editor.body).toBe(before);
    expect(editor.canUndo).toBe(false);
    expect(editor.canRedo).toBe(true);
    editor.redo();
    expect(required([...walkParagraphs(editor.body)][0]).runs[0]?.content).toEqual([
      { kind: 'text', text: '前旧改' },
    ]);
  });
});
