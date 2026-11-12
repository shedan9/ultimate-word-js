import type { DocRange, TextChangeSet, TextEditor, TextHistoryOptions } from '@uw/model';
import { createTextEditor, DEFAULT_SECTION_PROPS, paragraphText, walkParagraphs } from '@uw/model';
import { describe, expect, it } from 'vitest';
import type { EditingController } from './editing.ts';
import { createEditingController } from './editing.ts';

const at = (offset: number) => ({ nodeId: 'r', contentIndex: 0, offset });
const range = (start: number, end = start): DocRange => ({ start: at(start), end: at(end) });

function setup(broadcast: boolean, options: TextHistoryOptions = {}, text = '甲乙丙丁戊') {
  const source = createTextEditor(
    {
      sections: [
        {
          id: 's',
          props: DEFAULT_SECTION_PROPS,
          blocks: [
            {
              kind: 'paragraph',
              id: 'p',
              props: {},
              runs: text ? [{ kind: 'run', id: 'r', props: {}, content: [{ kind: 'text', text }] }] : [],
            },
          ],
        },
      ],
    },
    options,
  );
  const views: EditingController[] = [];
  function publish(change: TextChangeSet | undefined) {
    if (broadcast && change) for (const view of views) view.apply(change);
    return change;
  }
  const editor: TextEditor = {
    get body() {
      return source.body;
    },
    get canUndo() {
      return source.canUndo;
    },
    get canRedo() {
      return source.canRedo;
    },
    tx: (...args) => publish(source.tx(...args)),
    undo: () => publish(source.undo()),
    redo: () => publish(source.redo()),
    breakHistory: () => source.breakHistory(),
  };
  function mount() {
    const view = createEditingController(editor);
    views.push(view);
    return view;
  }
  const state = mount();
  function travel(action: 'undo' | 'redo') {
    const change = editor[action]();
    if (!broadcast && change) state.apply(change);
    return change;
  }
  return { editor, state, mount, travel, texts: () => [...walkParagraphs(editor.body)].map(paragraphText) };
}

describe.each([false, true])('选区历史（同步广播 %s）', (broadcast) => {
  it('空段首次输入的节点随历史恢复，撤销后仍可重新输入', () => {
    const { state, travel, texts } = setup(broadcast, {}, '');
    const initial = state.selection;
    state.insert('首次');
    const after = state.selection;
    travel('undo');
    expect(state.selection).toEqual(initial);
    expect(state.focus?.nodeId).toBe('p');
    travel('redo');
    expect(state.selection).toEqual(after);
    travel('undo');
    state.insert('再次');
    expect(texts()).toEqual(['再次']);
  });

  it('替换后撤销恢复原范围与方向，移动光标不覆盖重做位置', () => {
    const { state, travel, texts } = setup(broadcast);
    state.select(range(4, 1));
    state.insert('新');
    const after = state.selection;
    expect(texts()).toEqual(['甲新戊']);
    state.select(range(0));
    travel('undo');
    expect(state.selection).toEqual(range(1, 4));
    expect(state.anchor).toEqual(at(4));
    expect(state.focus).toEqual(at(1));
    state.select(range(5));
    travel('redo');
    expect(state.selection).toEqual(after);
    expect(state.focus).toEqual(after?.end);
  });

  it('连续输入合并保留首次输入前与最后输入后的光标', () => {
    const { editor, state, travel, texts } = setup(broadcast, { now: () => 0 });
    state.select(range(2));
    state.insert('一', true);
    state.insert('二', true);
    state.insert('三', true);
    const after = state.selection;
    state.move('forward');
    travel('undo');
    expect(texts()).toEqual(['甲乙丙丁戊']);
    expect(state.selection).toEqual(range(2));
    expect(editor.canUndo).toBe(false);
    travel('redo');
    expect(texts()).toEqual(['甲乙一二三丙丁戊']);
    expect(state.selection).toEqual(after);
  });

  it('格式拆 run 后撤销恢复原节点位置，重做保持反向选区', () => {
    const { state, travel } = setup(broadcast);
    state.select(range(4, 1));
    state.toggleFormat('bold');
    const after = state.selection;
    const focus = state.focus;
    travel('undo');
    expect(state.selection).toEqual(range(1, 4));
    expect(state.focus).toEqual(at(1));
    travel('redo');
    expect(state.selection).toEqual(after);
    expect(state.focus).toEqual(focus);
    expect(state.focus).toEqual(after?.start);
  });

  it('跨段粘贴与拆段分别恢复选区，撤销后继续输入落在原处', () => {
    const { state, travel, texts } = setup(broadcast);
    state.select(range(4, 1));
    state.insert('一\n二');
    const pasted = state.selection;
    state.enter();
    const entered = state.selection;
    travel('undo');
    expect(state.selection).toEqual(pasted);
    travel('undo');
    expect(state.selection).toEqual(range(1, 4));
    expect(state.focus).toEqual(at(1));
    travel('redo');
    travel('redo');
    expect(state.selection).toEqual(entered);
    travel('undo');
    travel('undo');
    state.insert('替换');
    expect(texts()).toEqual(['甲替换戊']);
  });

  it.each(['backward', 'forward'] as const)('按 %s 删除后恢复原光标', (direction) => {
    const { state, travel } = setup(broadcast);
    state.select(range(2));
    state.delete(direction);
    const after = state.selection;
    state.select(range(0));
    travel('undo');
    expect(state.selection).toEqual(range(2));
    travel('redo');
    expect(state.selection).toEqual(after);
  });

  it('组合输入替换选区独立撤销，不与后续普通输入合并', () => {
    const { state, travel } = setup(broadcast);
    state.select(range(4, 1));
    state.compositionStart();
    state.compositionEnd('中文');
    const composed = state.selection;
    state.insert('字', true);
    travel('undo');
    expect(state.selection).toEqual(composed);
    travel('undo');
    expect(state.selection).toEqual(range(1, 4));
    expect(state.focus).toEqual(at(1));
    travel('redo');
    expect(state.selection).toEqual(composed);
  });

  it('空操作与失败事务不改历史选区，成功跳转清掉暂存格式', () => {
    let fail = false;
    const { state, travel, editor } = setup(broadcast, {
      validate: () => {
        if (fail) throw new Error('预排版失败');
      },
    });
    state.select(range(2));
    state.insert('新');
    const after = state.selection;
    state.select(range(0));
    state.insert('');
    fail = true;
    expect(() => state.insert('失败')).toThrow('预排版失败');
    expect(() => travel('undo')).toThrow('预排版失败');
    expect(state.selection).toEqual(range(0));
    expect(editor.canUndo).toBe(true);
    fail = false;
    state.toggleFormat('italic');
    expect(state.pendingFormat?.italic).toBe(true);
    travel('undo');
    expect(state.pendingFormat).toBeUndefined();
    expect(state.selection).toEqual(range(2));
    travel('redo');
    expect(state.selection).toEqual(after);
  });
});

describe('多视图与宿主事务的选区历史', () => {
  it('重复广播同一变更不重复平移光标，也不覆盖原始历史选区', () => {
    const { editor, state, travel } = setup(true);
    state.select(range(3, 1));
    const change = editor.tx((tx) => {
      tx.insertText(at(0), '新');
    });
    const after = state.selection;
    if (change) state.apply(change);
    expect(state.selection).toEqual(after);
    travel('undo');
    expect(state.selection).toEqual(range(1, 3));
    expect(state.focus).toEqual(at(1));
    travel('redo');
    expect(state.selection).toEqual(after);
  });

  it('各视图独立恢复自己的范围，宿主修改和历史跳转共用一条广播', () => {
    const { editor, state, mount, travel } = setup(true);
    const second = mount();
    state.select(range(4, 1));
    second.select(range(5));
    state.insert('新');
    const firstAfter = state.selection;
    const secondAfter = second.selection;
    state.select(range(0));
    second.select(range(0));
    travel('undo');
    expect(state.selection).toEqual(range(1, 4));
    expect(state.focus).toEqual(at(1));
    expect(second.selection).toEqual(range(5));
    travel('redo');
    expect(state.selection).toEqual(firstAfter);
    expect(second.selection).toEqual(secondAfter);
    editor.tx((tx) => {
      tx.insertText(at(0), '宿主');
    });
    travel('undo');
    expect(state.selection).toEqual(firstAfter);
    expect(second.selection).toEqual(secondAfter);
  });

  it('输入合并中途挂载的视图对未知起始选区使用位置映射', () => {
    const { state, mount, travel } = setup(true, { now: () => 0 });
    state.select(range(2));
    state.insert('一', true);
    const second = mount();
    // 新视图默认在正文开头，不移动光标，避免打断原视图的连续输入合并。
    state.insert('二', true);
    travel('undo');
    expect(state.selection).toEqual(range(2));
    expect(second.selection).toEqual(range(0));
    travel('redo');
    expect(second.selection).toEqual(range(0));
  });
});
