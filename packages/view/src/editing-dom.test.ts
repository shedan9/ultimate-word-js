// @vitest-environment jsdom
import { createTextEditor, DEFAULT_SECTION_PROPS } from '@uw/model';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { DomView } from './dom.ts';
import { mountEditing } from './editing-dom.ts';

afterEach(() => document.body.replaceChildren());

describe('编辑输入节点的销毁', () => {
  it.each(['beforeinput', 'compositionend', 'keydown', 'paste', 'copy', 'cut'] as const)(
    '销毁后旧输入节点的 %s 不再修改文档、历史或剪贴板',
    (type) => {
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
                runs: [{ kind: 'run', id: 'r', props: {}, content: [{ kind: 'text', text: '原文' }] }],
              },
            ],
          },
        ],
      });
      const at = (offset: number) => ({ nodeId: 'r', contentIndex: 0, offset });
      editor.tx((t) => void t.insertText(at(0), '历史'));
      const container = document.body.appendChild(document.createElement('div'));
      const root = container.appendChild(document.createElement('div'));
      // 此处只验真实事件的生命周期；几何由既有浏览器断言负责。
      const view = { root, caretRect: () => null, scrollTo: () => false } as unknown as DomView;
      const unsubscribe = vi.fn();
      const editing = mountEditing(container, view, {
        editor,
        text: () => '历史原文',
        subscribe: () => unsubscribe,
      });
      editing.select({ start: at(0), end: at(4) });
      const input = container.querySelector('textarea');
      if (!input) throw new Error('缺少输入节点');
      if (type === 'compositionend')
        input.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));
      const before = editor.body;
      editing.dispose();
      expect(input.isConnected).toBe(false);
      expect(container.querySelector('[role="status"]')).toBeNull();
      expect(unsubscribe).toHaveBeenCalledTimes(1);
      const setData = vi.fn();
      let event: Event;
      const options = { bubbles: true, cancelable: true };
      if (type === 'beforeinput')
        event = new InputEvent(type, { ...options, inputType: 'insertText', data: '旧输入' });
      else if (type === 'compositionend') event = new CompositionEvent(type, { ...options, data: '组合串' });
      else if (type === 'keydown') event = new KeyboardEvent(type, { ...options, key: 'z', ctrlKey: true });
      else {
        event = new Event(type, options);
        Object.defineProperty(event, 'clipboardData', { value: { getData: () => '粘贴文字', setData } });
      }
      input.dispatchEvent(event);
      expect(editor.body).toBe(before);
      expect(editor.canUndo).toBe(true);
      expect(editor.canRedo).toBe(false);
      expect(setData).not.toHaveBeenCalled();
      expect(event.defaultPrevented).toBe(false);
    },
  );
});
