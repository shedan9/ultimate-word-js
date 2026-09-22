import type { DocPosition, DocRange, TextChangeSet, TextEditor } from '@uw/model';
import type { DomView } from './dom.ts';
import { createEditingController } from './editing.ts';
import { moveVertically } from './vertical-navigation.ts';

export interface EditingBinding {
  editor: TextEditor;
  subscribe(listener: (change: TextChangeSet) => void): () => void;
}
export interface DomEditing {
  readonly selection: DocRange | undefined;
  select(range: DocRange): void;
  focus(): void;
  refresh(): void;
  dispose(): void;
}

/** textarea 常驻容器，重排只替换页树，不能打断操作系统持有的 IME 节点。 */
export function mountEditing(container: Element, view: DomView, binding: EditingBinding): DomEditing {
  const doc = container.ownerDocument;
  const win = doc.defaultView;
  const state = createEditingController(binding.editor);
  const input = doc.createElement('textarea');
  input.dataset.uwInput = 'true';
  input.setAttribute('aria-label', '文档编辑输入');
  input.setAttribute('aria-multiline', 'true');
  input.autocomplete = 'off';
  input.spellcheck = false;
  // 坐标来自 caretRect；输入节点本身同时显示组合串与系统光标。
  input.style.cssText =
    'position:fixed;z-index:10;resize:none;padding:0;border:0;background:transparent;color:CanvasText;caret-color:CanvasText;font:16px sans-serif;overflow:hidden;min-width:2px;';
  const status = doc.createElement('span');
  status.setAttribute('role', 'status');
  status.style.cssText = 'position:fixed;bottom:8px;left:8px;background:Canvas;color:CanvasText;z-index:11';
  container.append(input, status);
  let decoration: ReturnType<DomView['decorate']> | undefined;
  let dead = false;
  let drag = false;
  let dragAnchor: DocPosition | undefined;
  let suppressCommit: string | undefined;
  let timer: ReturnType<typeof setTimeout> | undefined;
  let preferredX: number | undefined;
  function refresh(): void {
    if (dead) return;
    const range = state.selection;
    decoration?.dispose();
    decoration = undefined;
    if (!range) {
      input.style.visibility = 'hidden';
      return;
    }
    const rect = view.caretRect(state.focus ?? range.end);
    input.style.visibility = rect ? 'visible' : 'hidden';
    input.style.opacity = doc.activeElement === input ? '1' : '0';
    input.style.pointerEvents = 'none';
    if (rect) {
      input.style.left = `${rect.x}px`;
      input.style.top = `${rect.y}px`;
      input.style.height = `${Math.max(rect.height, 16)}px`;
      input.style.width = state.composing ? '240px' : '2px';
    }
    if (doc.activeElement === input)
      decoration = view.decorate(range, { style: { background: 'Highlight', opacity: '0.3' } });
  }
  function run(action: () => void, vertical = false): void {
    if (!vertical) preferredX = undefined;
    try {
      action();
      status.textContent = '';
    } catch (error) {
      status.textContent = error instanceof Error ? error.message : String(error);
    }
    if (!state.composing) input.value = '';
    refresh();
  }
  function focus(): void {
    input.style.visibility = 'visible';
    input.focus({ preventScroll: true });
    refresh();
  }
  function pointer(event: PointerEvent): void {
    if (
      event.button !== 0 ||
      state.composing ||
      win === null ||
      !(event.target instanceof win.Node) ||
      !view.root.contains(event.target)
    )
      return;
    const at = view.locate(event);
    if (!at) return;
    event.preventDefault();
    const start = event.shiftKey ? (state.anchor ?? at) : at;
    run(() => state.select({ start, end: at }));
    drag = true;
    dragAnchor = start;
    focus();
  }
  function pointerMove(event: PointerEvent): void {
    if (!drag) return;
    const at = view.locate(event);
    const range = state.selection;
    if (at && range) run(() => state.select({ start: dragAnchor ?? range.start, end: at }));
  }
  function pointerEnd(): void {
    drag = false;
  }
  function doubleClick(event: MouseEvent): void {
    if (event.button !== 0 || state.composing || win === null || !(event.target instanceof win.Node)) return;
    if (!view.root.contains(event.target)) return;
    const at = view.locate(event);
    if (!at) return;
    event.preventDefault();
    drag = false;
    const caret = view.caretRect(at);
    run(() => state.selectWord(at, caret && event.clientX < caret.x ? 'before' : 'after'));
    focus();
  }
  function beforeInput(event: InputEvent): void {
    if (event.isComposing || state.composing || event.inputType === 'insertCompositionText') return;
    event.preventDefault();
    if (
      suppressCommit !== undefined &&
      (event.inputType === 'insertFromComposition' || event.data === suppressCommit)
    ) {
      suppressCommit = undefined;
      input.value = '';
      return;
    }
    run(() => {
      switch (event.inputType) {
        case 'insertText':
        case 'insertReplacementText':
          state.insert(event.data ?? '', true);
          break;
        case 'insertParagraph':
        case 'insertLineBreak':
          state.enter();
          break;
        case 'deleteContentBackward':
          state.delete('backward');
          break;
        case 'deleteContentForward':
          state.delete('forward');
          break;
        case 'historyUndo':
          binding.editor.undo();
          break;
        case 'historyRedo':
          binding.editor.redo();
          break;
      }
    });
  }
  function keydown(event: KeyboardEvent): void {
    if (event.isComposing || state.composing || event.keyCode === 229) return;
    let action: (() => void) | undefined;
    const mod = event.ctrlKey || event.metaKey;
    if (mod && event.key.toLowerCase() === 'z')
      action = () => (event.shiftKey ? binding.editor.redo() : binding.editor.undo());
    else if (mod && event.key.toLowerCase() === 'y') action = () => binding.editor.redo();
    else if (
      !event.metaKey &&
      !(event.ctrlKey && event.altKey) &&
      (event.key === 'ArrowLeft' || event.key === 'ArrowRight')
    )
      action = () =>
        state.move(
          event.key === 'ArrowLeft' ? 'backward' : 'forward',
          event.shiftKey,
          event.ctrlKey || event.altKey ? 'word' : 'grapheme',
        );
    else if (!mod && !event.altKey) {
      if (event.key === 'ArrowUp' || event.key === 'ArrowDown')
        action = () => {
          const range = state.selection;
          if (!range) return;
          const move = moveVertically(
            view.index,
            state.focus ?? range.end,
            event.key === 'ArrowUp' ? 'up' : 'down',
            preferredX,
          );
          if (move) {
            preferredX = move.x;
            state.select({
              start: event.shiftKey ? (state.anchor ?? range.start) : move.position,
              end: move.position,
            });
          }
        };
      else if (event.key === 'Enter') action = () => state.enter();
      else if (event.key === 'Backspace' || event.key === 'Delete')
        action = () => state.delete(event.key === 'Backspace' ? 'backward' : 'forward');
    }
    if (action) {
      event.preventDefault();
      run(action, event.key === 'ArrowUp' || event.key === 'ArrowDown');
      const at = state.focus;
      if (at) view.scrollTo(at, { align: 'nearest' });
    }
  }
  function compositionStart(): void {
    suppressCommit = undefined;
    run(() => state.compositionStart());
  }
  function compositionEnd(event: CompositionEvent): void {
    if (!state.composing) return;
    suppressCommit = event.data;
    clearTimeout(timer);
    timer = setTimeout(() => {
      suppressCommit = undefined;
    }, 0);
    run(() => state.compositionEnd(event.data));
  }
  function blur(): void {
    run(() => {
      state.compositionCancel();
      binding.editor.breakHistory();
    });
  }
  function paste(event: ClipboardEvent): void {
    event.preventDefault();
    if (!state.composing) run(() => state.insert(event.clipboardData?.getData('text/plain') ?? ''));
  }
  const unsubscribe = binding.subscribe((change) => {
    preferredX = undefined;
    state.apply(change);
    refresh();
  });
  input.addEventListener('input', () => {
    if (!state.composing) input.value = '';
  });
  input.addEventListener('beforeinput', beforeInput);
  input.addEventListener('keydown', keydown);
  input.addEventListener('compositionstart', compositionStart);
  input.addEventListener('compositionend', compositionEnd);
  input.addEventListener('paste', paste);
  input.addEventListener('focus', refresh);
  input.addEventListener('blur', blur);
  container.addEventListener('pointerdown', pointer as EventListener);
  container.addEventListener('dblclick', doubleClick as EventListener);
  doc.addEventListener('pointermove', pointerMove);
  doc.addEventListener('pointerup', pointerEnd);
  doc.addEventListener('pointercancel', pointerEnd);
  doc.addEventListener('scroll', refresh, true);
  win?.addEventListener('resize', refresh);
  refresh();
  return {
    get selection() {
      return state.selection;
    },
    select(range) {
      preferredX = undefined;
      state.select(range);
      refresh();
    },
    focus,
    refresh,
    dispose() {
      dead = true;
      clearTimeout(timer);
      unsubscribe();
      decoration?.dispose();
      container.removeEventListener('pointerdown', pointer as EventListener);
      container.removeEventListener('dblclick', doubleClick as EventListener);
      doc.removeEventListener('pointermove', pointerMove);
      doc.removeEventListener('pointerup', pointerEnd);
      doc.removeEventListener('pointercancel', pointerEnd);
      doc.removeEventListener('scroll', refresh, true);
      win?.removeEventListener('resize', refresh);
      input.remove();
      status.remove();
    },
  };
}
