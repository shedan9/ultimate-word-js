import {
  createTextEditor,
  DEFAULT_SECTION_PROPS,
  paragraphText,
  rangeOfNode,
  walkParagraphs,
} from '@uw/model';
import { describe, expect, it } from 'vitest';
import { createEditingController } from './editing.ts';

function setup(text = '') {
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
            runs: text ? [{ kind: 'run', id: 'r', props: {}, content: [{ kind: 'text', text }] }] : [],
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
