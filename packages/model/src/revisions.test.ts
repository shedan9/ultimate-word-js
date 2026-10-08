/**
 * 修订的显示（revisions.ts）：三种视图怎么折进级联结果、作者配色、修订列表、编辑护栏。
 *
 * 判据都是「级联结果」而不是画面：最终 / 原始状态下看不见的那一半必须折成 `hidden` ——
 * 查找、复制、排版、命中测试都认这一个字段，折对了它们就都对。
 */
import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import type { CascadeContext } from './cascade.ts';
import type { Body, ResolvedBody, ResolvedRun } from './nodes.ts';
import { walkParagraphs } from './nodes.ts';
import { parseNumbering } from './numbering.ts';
import { parseBody } from './parse-body.ts';
import { createResolveCache, resolveBody } from './resolve-body.ts';
import type { RevisionView } from './revisions.ts';
import { listRevisions, REVISION_AUTHOR_COLORS, revisionAuthors, revisionDisplay } from './revisions.ts';
import { findText } from './search.ts';
import { DEFAULT_SETTINGS } from './settings.ts';
import { parseStyles } from './styles.ts';
import { createTextEditor, rangeEditable } from './text-transaction.ts';
import { EMPTY_THEME } from './theme.ts';

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const DOC = `
  <w:p>
    <w:r><w:t>原文</w:t></w:r>
    <w:del w:id="1" w:author="张三" w:date="2026-10-01T08:00:00Z"><w:r><w:delText>删掉</w:delText></w:r></w:del>
    <w:ins w:id="2" w:author="李四" w:date="2026-10-02T08:00:00Z"><w:r><w:t>新增</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>加粗</w:t></w:r></w:ins>
    <w:r><w:t>结尾</w:t></w:r>
  </w:p>
  <w:p><w:ins w:id="3" w:author="张三"><w:r><w:t>第二段</w:t></w:r></w:ins></w:p>`;

function setup() {
  const sink = createDiagnosticSink();
  const ctx: CascadeContext = {
    styles: parseStyles(parseXml(`<w:styles ${W_NS}/>`), sink),
    theme: EMPTY_THEME,
    settings: DEFAULT_SETTINGS,
    numbering: parseNumbering(undefined, sink),
  };
  const body: Body = parseBody(
    parseXml(`<w:document ${W_NS}><w:body>${DOC}<w:sectPr/></w:body></w:document>`),
    sink,
  );
  const authors = revisionAuthors([body]);
  const resolve = (view: RevisionView, b: Body = body): ResolvedBody =>
    resolveBody(ctx, b, { revisions: revisionDisplay(view, authors) });
  return { ctx, body, authors, resolve };
}

function runs(body: ResolvedBody): ResolvedRun[] {
  return [...walkParagraphs(body)].flatMap((p) => p.runs);
}

/** 这一版里看得见的字 */
function visible(body: ResolvedBody): string {
  return runs(body)
    .filter((r) => !r.props.hidden)
    .map((r) => r.content.map((c) => (c.kind === 'text' ? c.text : '')).join(''))
    .join('');
}

describe('三种视图折进级联结果', () => {
  it('最终状态：被删的字折成 hidden，插入的字与普通文字无异', () => {
    const { resolve } = setup();
    const r = resolve('final');
    expect(visible(r)).toBe('原文新增加粗结尾第二段');
    expect(runs(r).some((x) => x.props.revision !== undefined)).toBe(false);
    // 查找认 hidden：删掉的字搜不到，删除处两侧的字连得上（与 Word 最终状态下查找一致）
    expect(findText(r, '删掉')).toEqual([]);
    expect(findText(r, '原文新增')).toHaveLength(1);
  });

  it('原始状态：插入的字折成 hidden，被删的字照常显示', () => {
    const { resolve } = setup();
    expect(visible(resolve('original'))).toBe('原文删掉结尾');
  });

  it('所有标记：两样都显示，带作者色；作者按第一次出现的顺序配色', () => {
    const { resolve, authors } = setup();
    expect(authors).toEqual(['张三', '李四']);
    const r = resolve('markup');
    expect(visible(r)).toBe('原文删掉新增加粗结尾第二段');
    const marks = runs(r).map((x) => x.props.revision);
    const [red, blue] = REVISION_AUTHOR_COLORS;
    expect(marks).toEqual([
      undefined,
      { kind: 'delete', color: red },
      { kind: 'insert', color: blue },
      { kind: 'insert', color: blue },
      undefined,
      { kind: 'insert', color: red },
    ]);
    // 修订不改格式：加粗那一截照样加粗
    expect(runs(r)[3]?.props.bold).toBe(true);
  });

  it('换视图让级联备忘整份作废，同一份冻结的树也重新折', () => {
    const { ctx, body, authors } = setup();
    const frozen = createTextEditor(body).body;
    const cache = createResolveCache();
    const final = resolveBody(ctx, frozen, { cache, revisions: revisionDisplay('final', authors) });
    const markup = resolveBody(ctx, frozen, { cache, revisions: revisionDisplay('markup', authors) });
    expect(visible(final)).not.toBe(visible(markup));
    expect(visible(markup)).toBe('原文删掉新增加粗结尾第二段');
  });
});

describe('修订列表', () => {
  it('同一段里相邻、作者 / 时间 / 种类相同的 run 合成一处，范围从首 run 头到末 run 尾', () => {
    const { body } = setup();
    const list = listRevisions(body);
    expect(list.map(({ kind, author, date, text }) => ({ kind, author, date, text }))).toEqual([
      { kind: 'delete', author: '张三', date: '2026-10-01T08:00:00Z', text: '删掉' },
      { kind: 'insert', author: '李四', date: '2026-10-02T08:00:00Z', text: '新增加粗' },
      { kind: 'insert', author: '张三', date: undefined, text: '第二段' },
    ]);
    const ins = list[1];
    const p = [...walkParagraphs(body)][0];
    expect(ins?.range.start.nodeId).toBe(p?.runs[2]?.id);
    expect(ins?.range.end).toEqual({ nodeId: p?.runs[3]?.id, contentIndex: 0, offset: 2 });
  });
});

describe('编辑护栏', () => {
  it('往被删除的修订里插字被拒绝，整个事务回滚', () => {
    const { body } = setup();
    const editor = createTextEditor(body);
    const del = [...walkParagraphs(editor.body)][0]?.runs[1];
    expect(del?.revision?.kind).toBe('delete');
    const at = { nodeId: del?.id ?? '', contentIndex: 0, offset: 1 };
    const before = editor.body;
    expect(() => editor.tx((t) => void t.insertText(at, '字'))).toThrow('被删除的修订');
    expect(() => editor.tx((t) => void t.insertInline(at, 'tab'))).toThrow('被删除的修订');
    expect(editor.body).toBe(before);
    expect(editor.canUndo).toBe(false);
  });

  it('插入的修订里照常能打字；删掉一段被删除的修订等于接受了它', () => {
    const { body } = setup();
    const editor = createTextEditor(body);
    const [orig, del, ins] = [...walkParagraphs(editor.body)][0]?.runs ?? [];
    editor.tx((t) => void t.insertText({ nodeId: ins?.id ?? '', contentIndex: 0, offset: 2 }, '！'));
    const after = [...walkParagraphs(editor.body)][0]?.runs ?? [];
    expect(after[2]?.content).toEqual([{ kind: 'text', text: '新增！' }]);
    expect(after[2]?.revision?.kind).toBe('insert');
    // 从「原文」末尾删到「新增」开头：中间那段被删的修订一起清空
    editor.tx(
      (t) =>
        void t.deleteRange({
          start: { nodeId: orig?.id ?? '', contentIndex: 0, offset: 2 },
          end: { nodeId: ins?.id ?? '', contentIndex: 0, offset: 0 },
        }),
    );
    const cleared = [...walkParagraphs(editor.body)][0]?.runs.find((r) => r.id === del?.id);
    expect(cleared?.content.map((c) => (c.kind === 'text' ? c.text : '')).join('')).toBe('');
  });

  it('批量替换的预筛：首字落在被删除的修订里不可替换', () => {
    const { body } = setup();
    const del = [...walkParagraphs(body)][0]?.runs[1];
    const range = {
      start: { nodeId: del?.id ?? '', contentIndex: 0, offset: 0 },
      end: { nodeId: del?.id ?? '', contentIndex: 0, offset: 2 },
    };
    expect(rangeEditable(body, range)).toBe(false);
  });
});
