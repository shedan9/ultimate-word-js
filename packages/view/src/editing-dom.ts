import type {
  DocPosition,
  DocRange,
  Justification,
  RichFragment,
  TextChangeSet,
  TextEditor,
} from '@uw/model';
import { fragmentToHtml, htmlToParagraphs } from './clipboard.ts';
import type { DomView } from './dom.ts';
import type { FormatQuery, ParagraphQuery, ToggleFormat } from './editing.ts';
import { createEditingController } from './editing.ts';
import { moveVertically } from './vertical-navigation.ts';

export interface EditingBinding {
  editor: TextEditor;
  text(range: DocRange): string;
  /** 缺省时复制只写纯文本；给了就同时写一份 text/html（字体、字号、加粗…与对齐）。 */
  fragment?(range: DocRange): RichFragment;
  /** 缺省时切换只看暂存格式，选区一律按「未设置」处理。 */
  format?: FormatQuery;
  /** 缺省时对齐快捷键一律设置，不做「再按一次回到左对齐」。 */
  paragraphFormat?: ParagraphQuery;
  /** 文档的默认制表位（twips），Ctrl+M 的步长；缺省 420。 */
  tabStop?: number;
  subscribe(listener: (change: TextChangeSet) => void): () => void;
}
export interface DomEditing {
  readonly selection: DocRange | undefined;
  select(range: DocRange): void;
  focus(): void;
  refresh(): void;
  dispose(): void;
}

/** Word 与浏览器富文本共用的 B / I / U；Ctrl 与 Cmd 都认，Alt / Shift 组合留给系统与宿主。 */
const TOGGLE_KEYS: Readonly<Record<string, ToggleFormat>> = { b: 'bold', i: 'italic', u: 'underline' };
/** Word 的段落对齐快捷键；J 是两端对齐（`both`），中文公文的默认对齐。 */
const ALIGN_KEYS: Readonly<Record<string, Justification>> = { l: 'left', e: 'center', r: 'right', j: 'both' };

/** Word 的行距快捷键：Ctrl+5 是 1.5 倍，不是 5 倍。 */
const LINE_KEYS: Readonly<Record<string, 1 | 1.5 | 2>> = { '1': 1, '2': 2, '5': 1.5 };

/** textarea 常驻容器，重排只替换页树，不能打断操作系统持有的 IME 节点。 */
export function mountEditing(container: Element, view: DomView, binding: EditingBinding): DomEditing {
  const doc = container.ownerDocument;
  const win = doc.defaultView;
  const state = createEditingController(
    binding.editor,
    binding.format,
    binding.paragraphFormat,
    binding.tabStop === undefined ? {} : { tabStop: binding.tabStop },
  );
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
        case 'deleteWordBackward':
          state.delete('backward', 'word');
          break;
        case 'deleteWordForward':
          state.delete('forward', 'word');
          break;
        case 'formatBold':
          state.toggleFormat('bold');
          break;
        case 'formatItalic':
          state.toggleFormat('italic');
          break;
        case 'formatUnderline':
          state.toggleFormat('underline');
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
    // 列表段首 / 列表选区的 Tab 升降级；其余位置 Tab 插制表位（Word 的行为），
    // Shift+Tab 不接管，留给浏览器把焦点移出编辑区 —— 否则键盘用户出不去。
    if (event.key === 'Tab' && !event.ctrlKey && !event.metaKey && !event.altKey) {
      const list = state.listIndentable();
      if (!list && event.shiftKey) return;
      event.preventDefault();
      run(() => (list ? state.indentList(event.shiftKey ? 'out' : 'in') : state.insertInline('tab')));
      return;
    }
    let action: (() => void) | undefined;
    const mod = event.ctrlKey || event.metaKey;
    if (mod && event.key.toLowerCase() === 'z')
      action = () => (event.shiftKey ? binding.editor.redo() : binding.editor.undo());
    else if (mod && event.key.toLowerCase() === 'y') action = () => binding.editor.redo();
    // Word 的 Ctrl+M / Ctrl+Shift+M 与 Ctrl+1 / 2 / 5。只认 Ctrl：Mac 上 Cmd+M 最小化窗口、
    // Cmd+数字切标签页，都到不了页面。Windows 的 Ctrl+数字同样被浏览器截走，行距那三个键在那里不生效。
    else if (event.ctrlKey && !event.metaKey && !event.altKey && event.key.toLowerCase() === 'm')
      action = () => state.indent(event.shiftKey ? 'out' : 'in');
    // Ctrl+T / Ctrl+Shift+T 悬挂缩进、Ctrl+0 段前间距。同样只认 Ctrl；Windows 上 Chrome 把
    // Ctrl+T（新标签页）与 Ctrl+0（重置缩放）截走，这两个键只在 Mac 上到得了页面。
    else if (event.ctrlKey && !event.metaKey && !event.altKey && event.key.toLowerCase() === 't')
      action = () => state.hangingIndent(event.shiftKey ? 'out' : 'in');
    else if (event.ctrlKey && !event.metaKey && !event.altKey && !event.shiftKey && event.key === '0')
      action = () => state.toggleSpaceBefore();
    else if (event.ctrlKey && !event.metaKey && !event.altKey && !event.shiftKey && event.key in LINE_KEYS) {
      const multiple = LINE_KEYS[event.key] as 1 | 1.5 | 2;
      action = () => state.lineSpacing(multiple);
    }
    // Word 的 Ctrl+Shift+L 是「列表项目符号」样式；样式不一定存在，这里直接套项目符号列表。
    else if (mod && !event.altKey && event.shiftKey && event.key.toLowerCase() === 'l')
      action = () => state.toggleList('bullet');
    // Word 的 Ctrl+Enter（Mac 上 Cmd+Return）插分页符；Ctrl+Shift+Enter 是分栏符，单栏文档里没有意义，不接管。
    else if (mod && !event.altKey && !event.shiftKey && event.key === 'Enter')
      action = () => state.pageBreak();
    else if (mod && !event.altKey && !event.shiftKey && event.key.toLowerCase() === 'a')
      action = () => state.selectAll();
    else if (mod && !event.altKey && !event.shiftKey && event.key.toLowerCase() in TOGGLE_KEYS) {
      const format = TOGGLE_KEYS[event.key.toLowerCase()] as ToggleFormat;
      action = () => state.toggleFormat(format);
    } else if (mod && !event.altKey && !event.shiftKey && event.key.toLowerCase() in ALIGN_KEYS) {
      const justification = ALIGN_KEYS[event.key.toLowerCase()] as Justification;
      action = () => state.align(justification);
    } else if (
      !event.altKey &&
      ((event.ctrlKey && !event.metaKey && (event.key === 'Home' || event.key === 'End')) ||
        (event.metaKey && !event.ctrlKey && (event.key === 'ArrowUp' || event.key === 'ArrowDown')))
    )
      action = () =>
        state.moveToDocumentBoundary(
          event.key === 'Home' || event.key === 'ArrowUp' ? 'start' : 'end',
          event.shiftKey,
        );
    else if (
      !event.metaKey &&
      !(event.ctrlKey && event.altKey) &&
      (event.key === 'Backspace' || event.key === 'Delete')
    )
      action = () =>
        state.delete(
          event.key === 'Backspace' ? 'backward' : 'forward',
          event.ctrlKey || event.altKey ? 'word' : 'grapheme',
        );
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
      // Shift+Enter 是 Word 的软换行：换行不分段，段落格式与编号都不变。
      else if (event.key === 'Enter')
        action = event.shiftKey ? () => state.insertInline('lineBreak') : () => state.enter();
    }
    if (action) {
      event.preventDefault();
      run(action, !mod && !event.altKey && (event.key === 'ArrowUp' || event.key === 'ArrowDown'));
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
    if (state.composing) return;
    const data = event.clipboardData;
    const html = data?.getData('text/html') ?? '';
    const plain = data?.getData('text/plain') ?? '';
    run(() => {
      // HTML 读不出一个字（只有图片、或是空壳）时退回纯文本，别让粘贴什么都不做。
      const paragraphs = html && win ? htmlToParagraphs(html, new win.DOMParser()) : [];
      if (paragraphs.some((p) => p.runs.length)) state.insertParagraphs(paragraphs);
      else state.insert(plain);
    });
  }
  function clipboard(event: ClipboardEvent): void {
    if (event.defaultPrevented || state.composing) return;
    event.preventDefault();
    const range = state.selection;
    const data = event.clipboardData;
    if (!range || !data) return;
    run(() => {
      const text = binding.text(range);
      if (text === '') return;
      const fragment = binding.fragment?.(range);
      const write = () => {
        data.setData('text/plain', text);
        if (fragment) data.setData('text/html', fragmentToHtml(fragment));
      };
      if (event.type === 'cut') state.cut(write);
      else write();
    });
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
  input.addEventListener('copy', clipboard);
  input.addEventListener('cut', clipboard);
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
