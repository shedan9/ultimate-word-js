import type { Paragraph, Run, Table, TextEditor } from '@uw/model';
import {
  createTextEditor,
  DEFAULT_SECTION_PROPS,
  paragraphText,
  rangeOfNode,
  walkParagraphs,
} from '@uw/model';
import { describe, expect, it } from 'vitest';
import type { ParagraphQuery } from './editing.ts';
import { createEditingController } from './editing.ts';

function setup(text: string | Run[] = '') {
  const editor = createTextEditor({
    sections: [
      {
        id: 's',
        props: DEFAULT_SECTION_PROPS,
        blocks: [
          {
            kind: 'paragraph',
            id: 'p',
            props: {},
            runs:
              typeof text !== 'string'
                ? text
                : text
                  ? [{ kind: 'run', id: 'r', props: {}, content: [{ kind: 'text', text }] }]
                  : [],
          },
        ],
      },
    ],
  });
  const state = createEditingController(editor);
  const texts = () => [...walkParagraphs(editor.body)].map(paragraphText);
  return { editor, state, texts };
}
describe('编辑输入状态', () => {
  it('空段输入、Enter、段首退格合段以及撤销映射', () => {
    const { editor, state, texts } = setup();
    state.insert('甲');
    state.enter();
    state.insert('乙');
    state.move('backward');
    state.delete('backward');
    expect(texts()).toEqual(['甲乙']);
    const undo = editor.undo();
    if (undo) state.apply(undo);
    expect(texts()).toEqual(['甲', '乙']);
    state.insert('前');
    expect(texts()).toEqual(['甲前', '乙']);
  });
  it('组合期不改模型，确认只提交一次，取消保留选区文字', () => {
    const { editor, state, texts } = setup('旧字');
    const p = [...walkParagraphs(editor.body)][0];
    const range = p && rangeOfNode(p);
    if (!range) throw new Error('缺少段落');
    state.select(range);
    state.compositionStart();
    state.insert('错误');
    state.enter();
    expect(texts()).toEqual(['旧字']);
    state.compositionEnd('中文');
    expect(texts()).toEqual(['中文']);
    const undo = editor.undo();
    if (undo) state.apply(undo);
    expect(texts()).toEqual(['旧字']);
    expect(editor.canUndo).toBe(false);
    state.compositionStart();
    state.compositionCancel();
    state.compositionEnd('错误');
    expect(texts()).toEqual(['旧字']);
  });
  it('退格按字素删除 emoji 与组合音标，多行粘贴单次撤销', () => {
    const { editor, state, texts } = setup();
    state.insert('甲\n😀e\u0301');
    expect(texts()).toEqual(['甲', '😀e\u0301']);
    state.delete('backward');
    expect(texts()).toEqual(['甲', '😀']);
    state.delete('backward');
    expect(texts()).toEqual(['甲', '']);
    editor.undo();
    editor.undo();
    editor.undo();
    expect(texts()).toEqual(['']);
  });
});

const pos = (offset: number, nodeId = 'r', contentIndex = 0) => ({ nodeId, contentIndex, offset });
const run = (id: string, text: string): Run => ({
  kind: 'run',
  id,
  props: {},
  content: [{ kind: 'text', text }],
});

describe('全文选区与文档边界', () => {
  it('首末空段也属于全文，替换全文单次撤销恢复段落结构', () => {
    const { state, editor, texts } = setup();
    state.insert('\n甲\n\n乙\n');
    const paragraphs = [...walkParagraphs(editor.body)];
    const first = paragraphs[0];
    const last = paragraphs.at(-1);
    const start = first && rangeOfNode(first)?.start;
    const end = last && rangeOfNode(last)?.end;
    if (!start || !end) throw new Error('缺少测试段落');
    state.selectAll();
    expect(state.selection).toEqual({ start, end });
    expect(state.anchor).toEqual(start);
    expect(state.focus).toEqual(end);
    state.insert('替换');
    expect(texts()).toEqual(['替换']);
    editor.undo();
    expect(texts()).toEqual(['', '甲', '', '乙', '']);
    editor.redo();
    expect(texts()).toEqual(['替换']);
  });

  it('边界扩选保留原锚点，跨过锚点后反向，普通导航折叠到指定边界', () => {
    const { state, editor } = setup('甲乙丙丁');
    state.select({ start: pos(2), end: pos(3) });
    state.moveToDocumentBoundary('start', true);
    expect(state.selection).toEqual({ start: pos(0), end: pos(2) });
    expect(state.anchor).toEqual(pos(2));
    expect(state.focus).toEqual(pos(0));
    state.moveToDocumentBoundary('end', true);
    expect(state.selection).toEqual({ start: pos(2), end: pos(4) });
    state.moveToDocumentBoundary('start');
    expect(state.selection).toEqual({ start: pos(0), end: pos(0) });
    state.moveToDocumentBoundary('end');
    expect(state.selection).toEqual({ start: pos(4), end: pos(4) });
    expect(editor.canUndo).toBe(false);
  });

  it('跨节下钻首末表格，全选范围覆盖不同单元格但不放宽删除限制', () => {
    const paragraph = (id: string, runs: Run[]): Paragraph => ({ kind: 'paragraph', id, props: {}, runs });
    const table = (id: string, blocks: (Paragraph | Table)[]): Table => ({
      kind: 'table',
      id,
      props: {},
      grid: [],
      rows: [
        {
          kind: 'row',
          id: `${id}-row`,
          props: {},
          cells: [
            {
              kind: 'cell',
              id: `${id}-cell`,
              props: {},
              gridSpan: 1,
              vMerge: 'none',
              blocks,
            },
          ],
        },
      ],
    });
    const lastRun = run('last', '乙');
    lastRun.content.push({ kind: 'text', text: '😀' });
    const editor = createTextEditor({
      sections: [
        { id: 'empty-section', props: DEFAULT_SECTION_PROPS, blocks: [] },
        {
          id: 'first-section',
          props: DEFAULT_SECTION_PROPS,
          blocks: [table('first-table', [paragraph('empty', [])])],
        },
        {
          id: 'last-section',
          props: DEFAULT_SECTION_PROPS,
          blocks: [table('outer', [table('inner', [paragraph('last-p', [lastRun])])])],
        },
      ],
    });
    const state = createEditingController(editor);
    state.moveToDocumentBoundary('end');
    expect(state.focus).toEqual(pos(2, 'last', 1));
    state.moveToDocumentBoundary('start');
    expect(state.focus).toEqual(pos(0, 'empty'));
    state.selectAll();
    const selection = { start: pos(0, 'empty'), end: pos(2, 'last', 1) };
    expect(state.selection).toEqual(selection);
    const original = editor.body;
    expect(() => state.insert('替换')).toThrow();
    expect(editor.body).toBe(original);
    expect(state.selection).toEqual(selection);
    expect(editor.canUndo).toBe(false);
  });

  it('空文档与空段安全处理，组合期不改变选区', () => {
    const state = createEditingController(createTextEditor({ sections: [] }));
    state.selectAll();
    state.moveToDocumentBoundary('end', true);
    expect(state.selection).toBeUndefined();
    const empty = setup().state;
    empty.selectAll();
    empty.moveToDocumentBoundary('end');
    expect(empty.selection).toEqual({ start: pos(0, 'p'), end: pos(0, 'p') });
    const composing = setup('甲乙丙').state;
    composing.select({ start: pos(2), end: pos(1) });
    composing.compositionStart();
    composing.selectAll();
    composing.moveToDocumentBoundary('start', true);
    composing.moveToDocumentBoundary('end');
    expect(composing.anchor).toEqual(pos(2));
    expect(composing.focus).toEqual(pos(1));
  });

  it('边界导航打断连续输入合并，编辑后全选和导航重新读取文档边界', () => {
    const { state, editor, texts } = setup('甲');
    state.moveToDocumentBoundary('end');
    state.insert('乙', true);
    state.moveToDocumentBoundary('end');
    state.insert('丙', true);
    editor.undo();
    expect(texts()).toEqual(['甲乙']);
    state.moveToDocumentBoundary('end');
    expect(state.focus).toEqual(pos(2));
    editor.undo();
    expect(texts()).toEqual(['甲']);
    state.selectAll();
    expect(state.selection).toEqual({ start: pos(0), end: pos(1) });
  });
});

describe('按词删除', () => {
  it('双向删除沿用导航边界，包含经过的空白，标点单独删除', () => {
    const { state, texts, editor } = setup('hello  world!  ');
    state.select({ start: pos(3), end: pos(3) });
    state.delete('forward', 'word');
    expect(texts()).toEqual(['hel  world!  ']);
    expect(state.focus).toEqual(pos(3));
    state.delete('forward', 'word');
    expect(texts()).toEqual(['hel!  ']);
    state.delete('forward', 'word');
    expect(texts()).toEqual(['hel  ']);
    state.delete('forward', 'word');
    expect(texts()).toEqual(['hel']);
    state.delete('backward', 'word');
    expect(texts()).toEqual(['']);
    editor.undo();
    expect(texts()).toEqual(['hel']);
    const backward = setup('hello  world  ');
    backward.state.select({ start: pos(14), end: pos(14) });
    backward.state.delete('backward', 'word');
    expect(backward.texts()).toEqual(['hello  ']);
    backward.state.delete('backward', 'word');
    expect(backward.texts()).toEqual(['']);
  });

  it('中文词语跨样式与片段删除，一次撤销恢复格式和链接', () => {
    const left = run('a', '你好世');
    const right = run('b', '界');
    right.props = { bold: true };
    right.hyperlink = { url: 'https://example.test' };
    right.content.push({ kind: 'text', text: '欢迎' });
    const { state, editor, texts } = setup([left, right]);
    state.select({ start: pos(1, 'b'), end: pos(1, 'b') });
    state.delete('backward', 'word');
    expect(texts()).toEqual(['你好欢迎']);
    expect(state.focus).toEqual(pos(2, 'a'));
    editor.undo();
    expect([...walkParagraphs(editor.body)][0]?.runs).toEqual([left, right]);
    expect(editor.canUndo).toBe(false);
    editor.redo();
    expect(texts()).toEqual(['你好欢迎']);
  });

  it('跨样式 emoji 与组合音标保持完整，已有反向选区只删选中部分', () => {
    const { state, texts } = setup([run('a', 'e'), run('b', '\u0301👩'), run('c', '\u200d💻!')]);
    state.delete('forward', 'word');
    expect(texts()).toEqual(['👩\u200d💻!']);
    state.delete('forward', 'word');
    expect(texts()).toEqual(['!']);
    const selected = setup('hello world');
    selected.state.select({ start: pos(4), end: pos(1) });
    selected.state.delete('backward', 'word');
    expect(selected.texts()).toEqual(['ho world']);
    expect(selected.state.focus).toEqual(pos(1));
  });

  it('段落边界只合段，文档边界与组合期间不创建删除事务', () => {
    for (const direction of ['backward', 'forward'] as const) {
      const { state, editor, texts } = setup('one');
      state.move('forward', false, 'word');
      state.enter();
      state.insert('two');
      state.move('backward', false, 'word');
      if (direction === 'forward') state.move('backward');
      state.delete(direction, 'word');
      expect(texts()).toEqual(['onetwo']);
      editor.undo();
      expect(texts()).toEqual(['one', 'two']);
    }
    const { state, editor, texts } = setup('hello');
    state.delete('backward', 'word');
    state.move('forward', false, 'word');
    state.delete('forward', 'word');
    state.compositionStart();
    state.delete('backward', 'word');
    expect(texts()).toEqual(['hello']);
    expect(editor.canUndo).toBe(false);
    const empty = setup();
    empty.state.delete('backward', 'word');
    empty.state.delete('forward', 'word');
    expect(empty.texts()).toEqual(['']);
    expect(empty.editor.canUndo).toBe(false);
  });

  it('跨不可编辑片段的删除原子回滚并保留选区与历史', () => {
    const r = run('r', 'one');
    r.content.push({ kind: 'tab' }, { kind: 'text', text: 'two' });
    const { state, editor } = setup([r]);
    state.select({ start: pos(3), end: pos(3) });
    const body = editor.body;
    const selection = state.selection;
    expect(() => state.delete('forward', 'word')).toThrow();
    expect(editor.body).toBe(body);
    expect(editor.canUndo).toBe(false);
    expect(state.selection).toEqual(selection);
  });
});

describe('按词与跨样式导航', () => {
  it('按词跳过空白，Shift 反向扩选跨过锚点后仍保留方向', () => {
    const { state, editor } = setup('hello  world');
    state.move('forward', false, 'word');
    expect(state.focus).toEqual(pos(5));
    state.move('forward', false, 'word');
    expect(state.focus).toEqual(pos(12));
    state.move('backward', true, 'word');
    expect(state.selection).toEqual({ start: pos(7), end: pos(12) });
    const changes = editor.tx((tx) => {
      tx.insertText(pos(0), '!');
    });
    if (changes) state.apply(changes);
    expect(state.anchor).toEqual(pos(13));
    expect(state.focus).toEqual(pos(8));
    state.move('forward', true, 'word');
    expect(state.selection).toEqual({ start: pos(13), end: pos(13) });
    state.select({ start: pos(6), end: pos(6) });
    state.move('backward', true, 'word');
    state.move('forward', true, 'word');
    state.move('forward', true, 'word');
    expect(state.selection).toEqual({ start: pos(6), end: pos(13) });
    state.move('backward');
    expect(state.focus).toEqual(pos(6));
  });

  it('中文词语跨 run / 片段选中，替换与撤销保留原样式和链接', () => {
    const left = run('a', '你好世');
    const right = run('b', '界');
    right.props = { bold: true };
    right.hyperlink = { url: 'https://example.test' };
    right.content.push({ kind: 'text', text: '欢迎' });
    const { state, editor, texts } = setup([left, right]);
    state.selectWord(pos(0, 'b'));
    expect(state.selection).toEqual({ start: pos(2, 'a'), end: pos(1, 'b') });
    state.insert('地球');
    expect(texts()).toEqual(['你好地球欢迎']);
    editor.undo();
    expect([...walkParagraphs(editor.body)][0]?.runs).toEqual([left, right]);
    state.select({ start: pos(2, 'a'), end: pos(2, 'a') });
    state.move('forward', false, 'word');
    expect(state.focus).toEqual(pos(1, 'b'));
    state.move('forward', false, 'word');
    expect(state.focus).toEqual(pos(2, 'b', 1));
  });

  it('跨样式的组合音标和 ZWJ emoji 按完整字素移动与删除', () => {
    const { state, texts, editor } = setup([run('a', 'e'), run('b', '\u0301👩'), run('c', '\u200d💻!')]);
    state.move('forward');
    expect(state.focus).toEqual(pos(1, 'b'));
    state.move('forward');
    expect(state.focus).toEqual(pos(3, 'c'));
    state.delete('backward');
    expect(texts()).toEqual(['e\u0301!']);
    editor.undo();
    state.selectWord(pos(2, 'c'));
    expect(state.selection).toEqual({ start: pos(1, 'b'), end: pos(3, 'c') });
  });

  it('词中点击、段尾点击与空段落不产生无效位置，组合期忽略选词和移动', () => {
    const { state } = setup('hello 😀!');
    state.selectWord(pos(3));
    expect(state.selection).toEqual({ start: pos(0), end: pos(5) });
    state.selectWord(pos(5), 'before');
    expect(state.selection).toEqual({ start: pos(0), end: pos(5) });
    state.selectWord(pos(9));
    expect(state.selection).toEqual({ start: pos(8), end: pos(9) });
    state.compositionStart();
    state.selectWord(pos(1));
    state.move('backward', true, 'word');
    expect(state.selection).toEqual({ start: pos(8), end: pos(9) });
    const empty = setup().state;
    empty.selectWord(pos(0, 'p'));
    empty.move('forward', false, 'word');
    expect(empty.focus).toEqual(pos(0, 'p'));
  });

  it('制表位阻断选词，空片段不多停一次，段落边界逐段移动', () => {
    const r = run('r', 'one');
    r.content.push({ kind: 'tab' }, { kind: 'text', text: 'two' }, { kind: 'text', text: '' });
    const { state } = setup([r]);
    state.selectWord(pos(1, 'r', 2));
    expect(state.selection).toEqual({ start: pos(0, 'r', 2), end: pos(3, 'r', 2) });
    const second = setup('one').state;
    second.move('forward', false, 'word');
    second.enter();
    second.insert('two');
    second.move('backward', false, 'word');
    second.move('backward', false, 'word');
    expect(second.focus).toEqual(pos(3));
    second.move('backward', false, 'word');
    expect(second.focus).toEqual(pos(0));
  });
});

describe('纯文本剪切事务', () => {
  it('反向跨样式选区只剪选中内容，复制回调只执行一次且独立撤销', () => {
    const { state, editor, texts } = setup([
      run('a', '甲乙'),
      { ...run('b', '丙丁'), props: { bold: true } },
    ]);
    state.select({ start: pos(1, 'b'), end: pos(1, 'a') });
    const original = editor.body;
    let writes = 0;
    state.cut(() => {
      writes++;
    });
    expect(writes).toBe(1);
    expect(texts()).toEqual(['甲丁']);
    expect(state.focus).toEqual(pos(1, 'a'));
    editor.undo();
    expect(editor.body).toEqual(original);
    expect(editor.canUndo).toBe(false);
    editor.redo();
    expect(texts()).toEqual(['甲丁']);
  });

  it('剪贴板写入失败保留模型、反向选区与撤销历史', () => {
    const { state, editor } = setup('hello');
    state.select({ start: pos(4), end: pos(1) });
    const original = editor.body;
    const focus = state.focus;
    const selection = state.selection;
    expect(() =>
      state.cut(() => {
        throw new Error('写入失败');
      }),
    ).toThrow('写入失败');
    expect(editor.body).toBe(original);
    expect(state.selection).toEqual(selection);
    expect(state.focus).toEqual(focus);
    expect(editor.canUndo).toBe(false);
  });

  it('不可编辑范围在写剪贴板前拒绝，折叠选区与组合期不剪切', () => {
    const { state, editor } = setup([
      {
        ...run('r', '甲'),
        content: [{ kind: 'text', text: '甲' }, { kind: 'tab' }, { kind: 'text', text: '乙' }],
      },
    ]);
    let writes = 0;
    const write = () => {
      writes++;
    };
    state.cut(write);
    state.select({ start: pos(0), end: pos(1, 'r', 2) });
    const selection = state.selection;
    expect(() => state.cut(write)).toThrow();
    state.compositionStart();
    state.cut(write);
    expect(writes).toBe(0);
    expect(state.selection).toEqual(selection);
    expect(editor.canUndo).toBe(false);
  });

  it('只选段落接缝时剪切合段，撤销恢复空段', () => {
    const { state, editor, texts } = setup('甲');
    state.move('forward');
    state.enter();
    const paragraphs = [...walkParagraphs(editor.body)];
    const first = paragraphs[0] && rangeOfNode(paragraphs[0]);
    const last = paragraphs[1] && rangeOfNode(paragraphs[1]);
    if (!first || !last) throw new Error('缺少测试段落');
    const start = first.end;
    const end = last.start;
    state.select({ start, end });
    state.cut(() => {});
    expect(texts()).toEqual(['甲']);
    editor.undo();
    expect(texts()).toEqual(['甲', '']);
  });
});

describe('字符格式切换', () => {
  function formatted(text: string | Run[] = '') {
    const ctx = setup(text);
    // 测试树没有样式，直接格式即最终格式；门面传入级联后的格式。
    const state = createEditingController(ctx.editor, (range) => {
      const props: { bold: boolean; italic: boolean; underline: string }[] = [];
      for (const p of walkParagraphs(ctx.editor.body)) {
        const ids = p.runs.map((r) => r.id);
        const from = range.start.nodeId === p.id ? 0 : ids.indexOf(range.start.nodeId);
        const to = range.end.nodeId === p.id ? 0 : ids.indexOf(range.end.nodeId);
        if (!p.runs.length && from === 0)
          props.push({ bold: !!p.props.markRunProps?.bold, italic: false, underline: 'none' });
        for (const r of p.runs.slice(Math.max(0, from), to < 0 ? undefined : to + 1))
          props.push({
            bold: !!r.props.bold,
            italic: !!r.props.italic,
            underline: r.props.underline ?? 'none',
          });
      }
      return props;
    });
    return { ...ctx, state };
  }
  const runs = (editor: ReturnType<typeof setup>['editor']) =>
    [...walkParagraphs(editor.body)].flatMap((p) =>
      p.runs.map((r) => [paragraphText({ ...p, runs: [r] }), !!r.props.bold] as const),
    );

  it('选区切换加粗保留反向选区，全部加粗时再按一次取消，整次一个撤销单元', () => {
    const { editor, state } = formatted('甲乙丙丁');
    state.select({
      start: { nodeId: 'r', contentIndex: 0, offset: 3 },
      end: { nodeId: 'r', contentIndex: 0, offset: 1 },
    });
    state.toggleFormat('bold');
    expect(runs(editor)).toEqual([
      ['甲', false],
      ['乙丙', true],
      ['丁', false],
    ]);
    const bold = [...walkParagraphs(editor.body)][0]?.runs[1];
    expect(state.selection).toEqual({
      start: { nodeId: bold?.id, contentIndex: 0, offset: 0 },
      end: { nodeId: bold?.id, contentIndex: 0, offset: 2 },
    });
    expect(state.focus).toEqual(state.selection?.start);
    state.toggleFormat('bold');
    expect(runs(editor).map(([, b]) => b)).toEqual([false, false, false]);
    editor.undo();
    editor.undo();
    expect(runs(editor)).toEqual([['甲乙丙丁', false]]);
  });

  it('折叠光标暂存格式到下一次输入，再按一次抵消，移动光标丢弃', () => {
    const { editor, state } = formatted('正文');
    state.select({
      start: { nodeId: 'r', contentIndex: 0, offset: 2 },
      end: { nodeId: 'r', contentIndex: 0, offset: 2 },
    });
    state.toggleFormat('bold');
    expect(state.pendingFormat).toEqual({ bold: true });
    expect(editor.canUndo).toBe(false);
    state.insert('加', true);
    state.insert('粗', true);
    expect(runs(editor)).toEqual([
      ['正文', false],
      ['加粗', true],
    ]);
    expect(state.pendingFormat).toBeUndefined();
    state.toggleFormat('italic');
    state.toggleFormat('italic');
    expect(state.pendingFormat).toEqual({ italic: false });
    state.move('backward');
    expect(state.pendingFormat).toBeUndefined();
  });

  it('空段落里切换直接改段落标记，组合输入提交时带上格式', () => {
    const { editor, state } = formatted();
    state.toggleFormat('underline');
    expect(state.pendingFormat).toBeUndefined();
    expect([...walkParagraphs(editor.body)][0]?.props.markRunProps?.underline).toBe('single');
    state.compositionStart();
    state.compositionEnd('下划线');
    expect([...walkParagraphs(editor.body)][0]?.runs[0]?.props.underline).toBe('single');
    const other = formatted('正文');
    other.state.toggleFormat('bold');
    other.state.compositionStart();
    other.state.toggleFormat('italic');
    expect(other.state.pendingFormat).toEqual({ bold: true });
    other.state.compositionEnd('前');
    expect(runs(other.editor)).toEqual([
      ['前', true],
      ['正文', false],
    ]);
  });
});

/**
 * 测试树没有样式：numId > 0 即视为画出了编号，层级取直接格式；
 * 编号格式取树上的定义（新建列表加进去的），查不到按十进制。
 */
function paragraphQuery(editor: TextEditor): ParagraphQuery {
  return (range) => {
    const paragraphs = [...walkParagraphs(editor.body)];
    const at = (id: string) => paragraphs.findIndex((p) => p.id === id || p.runs.some((r) => r.id === id));
    return paragraphs.slice(at(range.start.nodeId), at(range.end.nodeId) + 1).map((p) => {
      const numId = p.props.numbering?.numId ?? 0;
      const level = p.props.numbering?.level ?? 0;
      const defs = editor.body.numbering;
      const abstract = defs?.abstract[defs.instances[numId]?.abstractNumId ?? -1];
      const format = abstract?.levels[level]?.numFmt ?? 'decimal';
      // 编号层级的缩进按 Word 默认的每级 420 模拟，直接格式盖在上面
      const indent = {
        ...{ left: 0, right: 0, firstLine: 0, hanging: 0, leftChars: 0, rightChars: 0, firstLineChars: 0 },
        ...{ hangingChars: 0, ...(numId > 0 ? { left: 420 * (level + 1), hanging: 420 } : {}) },
        ...p.props.indent,
      };
      return {
        id: p.id,
        props: {
          justification: p.props.justification ?? 'both',
          indent,
          numbering: numId > 0 ? { numId, level, label: { format } as never } : { numId, level },
        },
      };
    });
  };
}

describe('段落对齐切换', () => {
  it('折叠光标改所在段，再按一次回到左对齐，选区与方向不变', () => {
    const ctx = setup('甲乙丙');
    const state = createEditingController(ctx.editor, undefined, paragraphQuery(ctx.editor));
    const range = {
      start: { nodeId: 'r', contentIndex: 0, offset: 2 },
      end: { nodeId: 'r', contentIndex: 0, offset: 1 },
    };
    state.select(range);
    const before = { selection: state.selection, focus: state.focus };
    const justification = () => [...walkParagraphs(ctx.editor.body)][0]?.props.justification;
    state.align('center');
    expect(justification()).toBe('center');
    expect({ selection: state.selection, focus: state.focus }).toEqual(before);
    state.align('center');
    expect(justification()).toBe('left');
    state.align('right');
    expect(justification()).toBe('right');
    ctx.editor.undo();
    ctx.editor.undo();
    expect(justification()).toBe('center');
    state.compositionStart();
    state.align('both');
    expect(justification()).toBe('center');
  });
});

describe('列表层级', () => {
  function list(levels: number[], texts: string[] = levels.map((_, i) => `项${i}`)) {
    const editor = createTextEditor({
      sections: [
        {
          id: 's',
          props: DEFAULT_SECTION_PROPS,
          blocks: levels.map((level, i) => ({
            kind: 'paragraph' as const,
            id: `p${i}`,
            props: level < 0 ? {} : { numbering: { numId: 3, level } },
            runs: texts[i]
              ? [
                  {
                    kind: 'run' as const,
                    id: `r${i}`,
                    props: {},
                    content: [{ kind: 'text' as const, text: texts[i] as string }],
                  },
                ]
              : [],
          })),
        },
      ],
    });
    const state = createEditingController(editor, undefined, paragraphQuery(editor));
    const numbering = () => [...walkParagraphs(editor.body)].map((p) => p.props.numbering);
    return { editor, state, numbering };
  }
  const at = (id: string, offset = 0) => ({ nodeId: id, contentIndex: 0, offset });

  it('段首 Tab 降一级、Shift Tab 升一级，段中与非列表段不接管，层级夹在 0–8', () => {
    const { editor, state, numbering } = list([0, 8, -1]);
    state.select({ start: at('r0'), end: at('r0') });
    expect(state.listIndentable()).toBe(true);
    state.indentList('in');
    expect(numbering()[0]).toEqual({ numId: 3, level: 1 });
    state.indentList('out');
    state.indentList('out');
    expect(numbering()[0]).toEqual({ numId: 3, level: 0 });
    state.select({ start: at('r1'), end: at('r1') });
    state.indentList('in');
    expect(numbering()[1]?.level).toBe(8);
    state.select({ start: at('r0', 1), end: at('r0', 1) });
    expect(state.listIndentable()).toBe(false);
    state.select({ start: at('r2'), end: at('r2') });
    expect(state.listIndentable()).toBe(false);
    editor.undo();
    expect(numbering()[0]).toEqual({ numId: 3, level: 1 });
  });

  it('跨段选区里每段各自升降一级，一次撤销；混入非列表段则不接管', () => {
    const { editor, state, numbering } = list([0, 2, -1]);
    state.select({ start: at('r0', 1), end: at('r1', 1) });
    state.indentList('in');
    expect(numbering().map((n) => n?.level)).toEqual([1, 3, undefined]);
    expect(state.selection).toEqual({ start: at('r0', 1), end: at('r1', 1) });
    editor.undo();
    expect(numbering().map((n) => n?.level)).toEqual([0, 2, undefined]);
    state.select({ start: at('r1', 1), end: at('r2', 1) });
    expect(state.listIndentable()).toBe(false);
  });

  it('空列表项 Enter 先升级再结束列表，段首退格去编号而不合段', () => {
    const { state, numbering, editor } = list([1, 0], ['', '正文']);
    const empty = at('p0');
    state.select({ start: empty, end: empty });
    state.enter();
    expect(numbering()[0]).toEqual({ numId: 3, level: 0 });
    state.enter();
    expect(numbering()[0]).toEqual({ numId: 0, level: 0 });
    state.enter();
    expect([...walkParagraphs(editor.body)]).toHaveLength(3);
    state.select({ start: at('r1'), end: at('r1') });
    state.delete('backward');
    expect(numbering()[2]).toEqual({ numId: 0, level: 0 });
    expect([...walkParagraphs(editor.body)].map(paragraphText)).toEqual(['', '', '正文']);
    state.delete('backward');
    expect([...walkParagraphs(editor.body)].map(paragraphText)).toEqual(['', '正文']);
  });

  it('段首退格去编号时把层级缩进写成直接格式，文字留在原处；空项 Enter 结束列表不写', () => {
    const { editor, state } = list([2, 0], ['项', '']);
    const props = (i: number) => [...walkParagraphs(editor.body)][i]?.props;
    state.select({ start: at('r0'), end: at('r0') });
    state.delete('backward');
    expect(props(0)?.indent).toEqual({
      left: 1260,
      leftChars: 0,
      firstLine: 0,
      firstLineChars: 0,
      hanging: 0,
      hangingChars: 0,
    });
    editor.undo();
    expect(props(0)?.indent).toBeUndefined();
    expect(props(0)?.numbering).toEqual({ numId: 3, level: 2 });
    state.select({ start: at('p1'), end: at('p1') });
    state.enter();
    expect(props(1)).toEqual({ numbering: { numId: 0, level: 0 } });
  });

  it('新建列表：普通段套用新定义，再按一次取消；一次撤销连定义回退', () => {
    const { editor, state, numbering } = list([-1, -1, -1]);
    state.select({ start: at('r0', 1), end: at('r1', 1) });
    state.toggleList('decimal');
    expect(numbering()).toEqual([{ numId: 1, level: 0 }, { numId: 1, level: 0 }, undefined]);
    expect(editor.body.numbering?.instances[1]).toBeDefined();
    expect(state.selection).toEqual({ start: at('r0', 1), end: at('r1', 1) });
    state.toggleList('decimal');
    expect(numbering().map((n) => n?.numId)).toEqual([0, 0, undefined]);
    editor.undo();
    editor.undo();
    expect(numbering()).toEqual([undefined, undefined, undefined]);
    expect(editor.body.numbering).toBeUndefined();
  });

  it('紧邻上一段是同类列表时接着数，不同类时新建；已是列表的段落保留层级', () => {
    const { editor, state, numbering } = list([2, -1, -1]);
    state.select({ start: at('r1'), end: at('r1') });
    state.toggleList('decimal');
    expect(numbering()[1]).toEqual({ numId: 3, level: 0 });
    state.select({ start: at('r0'), end: at('r2') });
    state.toggleList('bullet');
    expect(numbering()).toEqual([
      { numId: 1, level: 2 },
      { numId: 1, level: 0 },
      { numId: 1, level: 0 },
    ]);
    expect(editor.body.numbering?.abstract[0]?.levels[0]?.numFmt).toBe('bullet');
    // 选区里已有项目符号列表：并进去，不再新建
    state.select({ start: at('r2'), end: at('r2') });
    state.toggleList('bullet');
    state.toggleList('bullet');
    expect(Object.keys(editor.body.numbering?.instances ?? {})).toEqual(['1']);
  });
});
