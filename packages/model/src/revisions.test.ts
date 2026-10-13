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
import { buildRunOrder, compareDocPositions, rangeOfNode } from './order.ts';
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

function setup(xml = DOC) {
  const sink = createDiagnosticSink();
  const ctx: CascadeContext = {
    styles: parseStyles(parseXml(`<w:styles ${W_NS}/>`), sink),
    theme: EMPTY_THEME,
    settings: DEFAULT_SETTINGS,
    numbering: parseNumbering(undefined, sink),
  };
  const body: Body = parseBody(
    parseXml(`<w:document ${W_NS}><w:body>${xml}<w:sectPr/></w:body></w:document>`),
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
  const mark = (kind: 'ins' | 'del', author: string) => `<w:${kind} w:id="1" w:author="${author}"/>`;
  const p = (text: string, revision = '') =>
    `<w:p>${revision ? `<w:pPr><w:rPr>${revision}</w:rPr></w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`;
  const cell = (blocks: string) => `<w:tc>${blocks}</w:tc>`;
  const row = (cells: string, revision = '') =>
    `<w:tr>${revision ? `<w:trPr>${revision}</w:trPr>` : ''}${cells}</w:tr>`;
  const table = (rows: string) => `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>${rows}</w:tbl>`;

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
    expect(list.every((span) => span.target === 'text' && span.rowId === undefined)).toBe(true);
  });

  it('文字与段落标记分别列出：同一作者删掉整段也不能吞掉段落标记，范围折叠在本段末尾', () => {
    const { body } = setup(
      `<w:p><w:pPr><w:rPr>${mark('del', '张三')}</w:rPr></w:pPr>` +
        '<w:del w:id="2" w:author="张三"><w:r><w:delText>删段</w:delText></w:r></w:del></w:p>' +
        p('下段'),
    );
    const list = listRevisions(body);
    const first = [...walkParagraphs(body)][0];
    expect(list.map(({ target, kind, text }) => ({ target, kind, text }))).toEqual([
      { target: 'text', kind: 'delete', text: '删段' },
      { target: 'paragraph-mark', kind: 'delete', text: '\n' },
    ]);
    expect(list[1]?.paragraphId).toBe(first?.id);
    expect(list[1]?.range).toEqual({ start: list[0]?.range.end, end: list[0]?.range.end });
    expect(list[1]?.rowId).toBeUndefined();
  });

  it('空段落标记也可定位，移动修订的作者 / 时间 / move 保留', () => {
    const { body } = setup(
      '<w:p><w:pPr><w:rPr><w:moveTo w:id="1" w:author="张三" w:date="2026-10-01T08:00:00Z"/></w:rPr></w:pPr></w:p>',
    );
    const first = [...walkParagraphs(body)][0];
    expect(listRevisions(body)).toEqual([
      {
        target: 'paragraph-mark',
        kind: 'insert',
        move: true,
        author: '张三',
        date: '2026-10-01T08:00:00Z',
        paragraphId: first?.id,
        text: '\n',
        range: rangeOfNode(first as NonNullable<typeof first>),
      },
    ]);
  });

  it('整行修订单列在文字前，摘要保留已删文字并以制表位分格、换行分段；范围含首末空段', () => {
    const { body } = setup(
      table(
        row(
          cell(`<w:p/>${p('甲')}`) +
            cell(
              '<w:p><w:del w:id="2" w:author="李四"><w:r><w:delText>乙</w:delText></w:r></w:del></w:p><w:p/>',
            ),
          mark('del', '张三'),
        ),
      ),
    );
    const tbl = body.sections[0]?.blocks[0];
    if (tbl?.kind !== 'table' || tbl.rows[0] === undefined) throw new Error('缺少表格行');
    const list = listRevisions(body);
    const paragraphs = [...walkParagraphs(body)];
    expect(list.map(({ target, author, text }) => ({ target, author, text }))).toEqual([
      { target: 'row', author: '张三', text: '\n甲\t乙\n' },
      { target: 'text', author: '李四', text: '乙' },
    ]);
    expect(list[0]).toMatchObject({
      rowId: tbl.rows[0].id,
      paragraphId: paragraphs[0]?.id,
      range: rangeOfNode(tbl.rows[0]),
    });
    const order = buildRunOrder(body);
    for (const span of list) expect(compareDocPositions(order, span.range.start, span.range.end)).toBe(-1);
  });

  it('嵌套表格的修订行递归列出，外行的范围与摘要覆盖内表，空格的行也不丢', () => {
    const inner = table(row(cell(p('内甲')) + cell(p('内乙')), mark('ins', '李四')));
    const { body } = setup(
      table(row(cell(inner + p('外尾')), mark('del', '张三')) + row(cell('<w:p/>'), mark('ins', '王五'))),
    );
    const list = listRevisions(body);
    expect(list.map(({ target, text, author }) => ({ target, text, author }))).toEqual([
      { target: 'row', text: '内甲\t内乙\n外尾', author: '张三' },
      { target: 'row', text: '内甲\t内乙', author: '李四' },
      { target: 'row', text: '', author: '王五' },
    ]);
    expect(list[0]?.range.start).toEqual(list[1]?.range.start);
    expect(list[0]?.rowId).not.toBe(list[1]?.rowId);
    expect(list[2]?.range.start).toEqual(list[2]?.range.end);
    expect(revisionAuthors([body])).toEqual(['张三', '李四', '王五']);
  });

  it('作者按行首、文字、段尾标记的顺序去重，正文后接其他部件；纯结构修订也有作者色', () => {
    const { body, resolve } = setup(
      p('首段', mark('del', '标记作者')) +
        table(row(cell(p('格', mark('ins', '格标记作者'))), mark('del', '行作者'))) +
        '<w:p><w:ins w:id="2" w:author="文字作者"><w:r><w:t>字</w:t></w:r></w:ins></w:p>' +
        p('末段', mark('ins', '行作者')),
    );
    const other = setup(p('页眉', mark('ins', '页眉作者'))).body;
    const authors = revisionAuthors([body, other]);
    expect(authors).toEqual(['标记作者', '行作者', '格标记作者', '文字作者', '页眉作者']);
    expect(revisionDisplay('markup', authors).colors.行作者).toBe(REVISION_AUTHOR_COLORS[1]);
    expect(
      runs(resolve('markup')).find((r) => r.revision?.author === '文字作者')?.props.revision?.color,
    ).toBe(REVISION_AUTHOR_COLORS[3]);
  });

  it('异常行没有任何段落时不返回无效范围，作者仍收集；缺作者的结构修订记为空串', () => {
    const { body } = setup(table(row('', '<w:del w:id="1"/>')));
    expect(listRevisions(body)).toEqual([]);
    expect(revisionAuthors([body])).toEqual(['']);
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

describe('段落标记与整行的修订（结构在级联那一步折）', () => {
  const NUMBERING = `<w:numbering ${W_NS}>
    <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>
    <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
  const del = (id: number) => `<w:del w:id="${id}" w:author="张三" w:date="2026-10-01T08:00:00Z"/>`;
  const ins = (id: number) => `<w:ins w:id="${id}" w:author="李四"/>`;

  function resolveXml(xml: string, view: RevisionView): { body: Body; resolved: ResolvedBody } {
    const sink = createDiagnosticSink();
    const ctx: CascadeContext = {
      styles: parseStyles(parseXml(`<w:styles ${W_NS}/>`), sink),
      theme: EMPTY_THEME,
      settings: DEFAULT_SETTINGS,
      numbering: parseNumbering(parseXml(NUMBERING), sink),
    };
    const body = parseBody(
      parseXml(`<w:document ${W_NS}><w:body>${xml}<w:sectPr/></w:body></w:document>`),
      sink,
    );
    return {
      body,
      resolved: resolveBody(ctx, body, { revisions: revisionDisplay(view, revisionAuthors([body])) }),
    };
  }

  /** 每一块：段落写「id:对齐:看得见的字」，表格写行数 */
  function shape(r: ResolvedBody): string[] {
    return (r.sections[0]?.blocks ?? []).map((b) =>
      b.kind === 'table'
        ? `table:${b.rows.length}`
        : `${b.id}:${b.props.justification}:${b.props.numbering.label?.text ?? ''}${visible({ sections: [{ id: '', props: r.sections[0]?.props as never, blocks: [b] }] })}`,
    );
  }

  it('最终状态下被删的段落标记：本段并进下一段，用下一段的段落属性与 id；所有标记 / 原始状态下仍是两段', () => {
    const xml =
      `<w:p><w:pPr><w:jc w:val="right"/><w:rPr>${del(1)}</w:rPr></w:pPr><w:r><w:t>甲</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>乙</w:t></w:r></w:p>`;
    const { body, resolved } = resolveXml(xml, 'final');
    const [p1, p2] = [...walkParagraphs(body)];
    expect(p1?.markRevision).toEqual({ kind: 'delete', author: '张三', date: '2026-10-01T08:00:00Z' });
    expect(shape(resolved)).toEqual([`${p2?.id}:center:甲乙`]);
    expect(shape(resolveXml(xml, 'markup').resolved)).toHaveLength(2);
    expect(shape(resolveXml(xml, 'original').resolved)).toHaveLength(2);
  });

  it('整段删除（字与标记一起删）干干净净地消失，下一段格式不变；原始状态下插入的整段同理', () => {
    const deleted =
      `<w:p><w:pPr><w:jc w:val="right"/><w:rPr>${del(1)}</w:rPr></w:pPr><w:del w:id="2" w:author="张三"><w:r><w:delText>删掉的段</w:delText></w:r></w:del></w:p>` +
      `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>留下</w:t></w:r></w:p>`;
    expect(shape(resolveXml(deleted, 'final').resolved).map((s) => s.split(':').slice(1).join(':'))).toEqual([
      'center:留下',
    ]);
    const inserted =
      `<w:p><w:pPr><w:rPr>${ins(1)}</w:rPr></w:pPr><w:ins w:id="2" w:author="李四"><w:r><w:t>新段</w:t></w:r></w:ins></w:p>` +
      `<w:p><w:r><w:t>原有</w:t></w:r></w:p>`;
    expect(shape(resolveXml(inserted, 'original').resolved).map((s) => s.split(':').at(-1))).toEqual([
      '原有',
    ]);
    expect(shape(resolveXml(inserted, 'final').resolved)).toHaveLength(2);
  });

  it('并掉的段落不计数：被删的第 1 条之后，下一条显示 1.', () => {
    const num = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>';
    const xml =
      `<w:p><w:pPr>${num}<w:rPr>${del(1)}</w:rPr></w:pPr><w:del w:id="2" w:author="张三"><w:r><w:delText>甲</w:delText></w:r></w:del></w:p>` +
      `<w:p><w:pPr>${num}</w:pPr><w:r><w:t>乙</w:t></w:r></w:p>` +
      `<w:p><w:pPr>${num}</w:pPr><w:r><w:t>丙</w:t></w:r></w:p>`;
    expect(shape(resolveXml(xml, 'final').resolved).map((s) => s.split(':').at(-1))).toEqual([
      '1.乙',
      '2.丙',
    ]);
    expect(shape(resolveXml(xml, 'markup').resolved).map((s) => s.split(':').at(-1))).toEqual([
      '1.甲',
      '2.乙',
      '3.丙',
    ]);
  });

  it('下一块是表格或这一列到头了就不并（表格前那个段落标记 Word 删不掉）', () => {
    const xml =
      `<w:p><w:pPr><w:rPr>${del(1)}</w:rPr></w:pPr><w:r><w:t>甲</w:t></w:r></w:p>` +
      '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>格</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
      `<w:p><w:pPr><w:rPr>${del(2)}</w:rPr></w:pPr><w:r><w:t>末</w:t></w:r></w:p>`;
    expect(shape(resolveXml(xml, 'final').resolved).map((s) => s.split(':').at(-1))).toEqual([
      '甲',
      '1',
      '末',
    ]);
  });

  it('整行修订：最终状态下被删的行不进树，原始状态下插入的行不进树；行全没了整张表不出', () => {
    const row = (text: string, rev = '') =>
      `<w:tr>${rev ? `<w:trPr>${rev}</w:trPr>` : ''}<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc></w:tr>`;
    const tbl = (rows: string) =>
      `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>${rows}</w:tbl><w:p/>`;
    const xml = tbl(row('一') + row('二', del(1)) + row('三', ins(2)));
    const rowsOf = (view: RevisionView) => shape(resolveXml(xml, view).resolved)[0];
    expect(rowsOf('final')).toBe('table:2');
    expect(rowsOf('original')).toBe('table:2');
    expect(rowsOf('markup')).toBe('table:3');
    expect(shape(resolveXml(tbl(row('二', del(1))), 'final').resolved)).toHaveLength(1);
  });

  it('删掉合并区的首行：下一行的 continue 改成 restart，内容照常显示', () => {
    const cell = (merge: string, text: string) =>
      `<w:tc><w:tcPr><w:vMerge${merge ? ` w:val="${merge}"` : ''}/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
    const xml =
      `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:trPr>${del(1)}</w:trPr>${cell('restart', '首')}</w:tr>` +
      `<w:tr>${cell('', '')}</w:tr><w:tr>${cell('', '')}</w:tr></w:tbl><w:p/>`;
    const t = resolveXml(xml, 'final').resolved.sections[0]?.blocks[0];
    if (t?.kind !== 'table') throw new Error('表格不见了');
    expect(t.rows.map((r) => r.cells[0]?.vMerge)).toEqual(['restart', 'continue']);
  });

  it('拆段：原段落标记归后段，前段的标记是新打的；合段：留下的是后段的标记', () => {
    const { body } = resolveXml(
      `<w:p><w:pPr><w:rPr>${del(1)}</w:rPr></w:pPr><w:r><w:t>甲乙</w:t></w:r></w:p><w:p><w:r><w:t>丙</w:t></w:r></w:p>`,
      'final',
    );
    const editor = createTextEditor(body);
    const first = [...walkParagraphs(editor.body)][0];
    editor.tx((t) => void t.splitParagraph({ nodeId: first?.runs[0]?.id ?? '', contentIndex: 0, offset: 1 }));
    const [a, b, c] = [...walkParagraphs(editor.body)];
    expect([a?.markRevision?.kind, b?.markRevision?.kind, c?.markRevision]).toEqual([
      undefined,
      'delete',
      undefined,
    ]);
    // 合 a、b：删掉的是 a 的标记，留下 b 的（被删的）；再合 b、c：留下 c 的（没有修订）
    editor.tx((t) => void t.joinParagraph(a?.id ?? ''));
    expect([...walkParagraphs(editor.body)][0]?.markRevision?.kind).toBe('delete');
    editor.tx((t) => void t.joinParagraph(a?.id ?? ''));
    expect([...walkParagraphs(editor.body)][0]?.markRevision).toBeUndefined();
  });
});
