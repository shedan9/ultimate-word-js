import type { Run } from '@uw/model';
import {
  createTextEditor,
  DEFAULT_SECTION_PROPS,
  paragraphText,
  rangeOfNode,
  walkParagraphs,
} from '@uw/model';
import { describe, expect, it } from 'vitest';
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
