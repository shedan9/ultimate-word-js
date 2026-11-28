import type {
  BuiltinStyleName,
  DocPosition,
  DocRange,
  Justification,
  RichFragment,
  TextChangeSet,
  TextEditor,
} from '@uw/model';
import { cssColor, defaultFontFamily } from '@uw/render-dom';
import { fragmentToHtml, htmlToParagraphs } from './clipboard.ts';
import type { DomView } from './dom.ts';
import type { FormatQuery, ParagraphQuery, ParagraphStyleInfo, ToggleFormat } from './editing.ts';
import { createEditingController } from './editing.ts';
import type { ViewFormatting } from './formatting.ts';
import { applyFormatting, selectionFormatting } from './formatting.ts';
import { moveVertically } from './vertical-navigation.ts';
import type { EditingHit } from './view.ts';

export interface EditingBinding {
  headerFooter?(relId: string): EditingBinding;
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
  /** 文档的段落样式；缺省时 Ctrl+Alt+1/2/3 总是补一份内建标题定义，回车不按 `w:next` 换样式。 */
  styles?: readonly ParagraphStyleInfo[];
  subscribe(listener: (change: TextChangeSet) => void): () => void;
}
export interface DomEditing {
  readonly formatting: ViewFormatting;
  readonly selection: DocRange | undefined;
  pageFor(position: DocPosition): number | undefined;
  select(range: DocRange): void;
  focus(): void;
  refresh(): void;
  /**
   * 选区（含折叠的光标）按值变了就通知：点击、方向键、输入后光标后移、事务映射都算，
   * 滚动与缩放只重画不通知。`undefined` 是还没放过光标。
   */
  onSelectionChange(listener: (selection: DocRange | undefined) => void): () => void;
  dispose(): void;
}

function samePosition(a: DocPosition, b: DocPosition): boolean {
  return a.nodeId === b.nodeId && a.contentIndex === b.contentIndex && a.offset === b.offset;
}
function sameRange(a: DocRange | undefined, b: DocRange | undefined): boolean {
  if (a === undefined || b === undefined) return a === b;
  return samePosition(a.start, b.start) && samePosition(a.end, b.end);
}

/** Word 与浏览器富文本共用的 B / I / U；Ctrl 与 Cmd 都认，Alt / Shift 组合留给系统与宿主。 */
const TOGGLE_KEYS: Readonly<Record<string, ToggleFormat>> = { b: 'bold', i: 'italic', u: 'underline' };
/** Word 的段落对齐快捷键；J 是两端对齐（`both`），中文公文的默认对齐。 */
const ALIGN_KEYS: Readonly<Record<string, Justification>> = { l: 'left', e: 'center', r: 'right', j: 'both' };

/** Word 的行距快捷键：Ctrl+5 是 1.5 倍，不是 5 倍。 */
const LINE_KEYS: Readonly<Record<string, 1 | 1.5 | 2>> = { '1': 1, '2': 2, '5': 1.5 };
/**
 * Word 的 Ctrl+Alt+1 / 2 / 3（Mac 上 Cmd+Option+1 / 2 / 3）套标题 1–3。按 `code` 认：
 * Mac 上 Option 会把 `key` 变成「¡ ™ £」。
 */
const HEADING_CODES: Readonly<Record<string, BuiltinStyleName>> = {
  Digit1: 'heading 1',
  Digit2: 'heading 2',
  Digit3: 'heading 3',
};

/** textarea 常驻容器，重排只替换页树，不能打断操作系统持有的 IME 节点。 */
export function mountEditing(
  container: Element,
  view: DomView,
  bodyBinding: EditingBinding,
  fontFamily = defaultFontFamily,
  formatChanged?: () => void,
): DomEditing {
  const doc = container.ownerDocument;
  const win = doc.defaultView;
  const controllerFor = (binding: EditingBinding) =>
    createEditingController(binding.editor, binding.format, binding.paragraphFormat, {
      ...(binding.tabStop === undefined ? {} : { tabStop: binding.tabStop }),
      ...(binding.styles === undefined ? {} : { styles: binding.styles }),
    });
  let binding = bodyBinding;
  const bodyState = controllerFor(binding);
  let state = bodyState;
  type ActiveFrame = NonNullable<EditingHit['frame']> & { page: number };
  let active: ActiveFrame | undefined;
  let pendingHit: EditingHit | undefined;
  const input = doc.createElement('textarea');
  // 切回预览后，旧输入节点仍可能被宿主持有或收到迟到的组合 / 剪贴板事件；移除 DOM 不会移除监听者。
  const inputEvents = new (win?.AbortController ?? AbortController)();
  const inputEventOptions = { signal: inputEvents.signal };
  input.dataset.uwInput = 'true';
  input.setAttribute('aria-label', '文档编辑输入');
  input.setAttribute('aria-multiline', 'true');
  input.autocomplete = 'off';
  input.spellcheck = false;
  input.wrap = 'off';
  // 文档内的焦点由原生插入光标和选区标明；宿主的表单样式不应在字缝画出输入框。
  input.style.cssText =
    'position:fixed;z-index:10;pointer-events:none;appearance:none;resize:none;' +
    'margin:0;padding:0;border:0;border-radius:0;outline:none;box-shadow:none;' +
    'background:transparent;color:black;caret-color:currentColor;font:16px serif;' +
    'box-sizing:border-box;overflow:hidden;min-width:2px;min-height:0;field-sizing:content;';
  const autoWidth = win?.CSS?.supports('field-sizing', 'content');
  let compositionSize: number | undefined;
  function sizeInput(): void {
    if (state.composing && autoWidth) input.style.width = 'auto';
    else {
      input.style.width = '2px';
      // 较早的浏览器没有 field-sizing，先收窄再量，退格时也能缩回。
      if (state.composing && input.value) input.style.width = `${input.scrollWidth + 1}px`;
    }
  }
  const status = doc.createElement('span');
  status.setAttribute('role', 'status');
  status.style.cssText = 'position:fixed;bottom:8px;left:8px;background:Canvas;color:CanvasText;z-index:11';
  const modeButton = doc.createElement('button');
  modeButton.type = 'button';
  const boundary = doc.createElement('div');
  boundary.style.cssText =
    'position:absolute;left:0;right:0;border-color:#708399;border-width:0;pointer-events:none;z-index:2;box-sizing:border-box';
  boundary.append(modeButton);
  modeButton.dataset.uwStory = '';
  modeButton.style.cssText =
    'position:absolute;left:4px;padding:3px 8px;font:12px/1.4 system-ui;pointer-events:auto;' +
    'border:1px solid #8192ab;border-radius:4px;background:Canvas;color:CanvasText;cursor:pointer';
  container.append(input, status);
  function showMode(): void {
    const target = pendingHit?.frame ? { ...pendingHit.frame, page: pendingHit.page } : active;
    const kind = target?.kind;
    const edge = target && view.frameBoundary(target.page, target.kind);
    if (!edge || edge.relId !== target?.relId) {
      boundary.remove();
    } else {
      const header = kind === 'header';
      boundary.style.top = header ? '0' : `${edge.percent}%`;
      boundary.style.bottom = header ? `${100 - edge.percent}%` : '0';
      boundary.style.borderTop = header ? '0' : '1px dashed #708399';
      boundary.style.borderBottom = header ? '1px dashed #708399' : '0';
      modeButton.style.top = header ? '4px' : 'auto';
      modeButton.style.bottom = header ? 'auto' : '4px';
      if (boundary.parentNode !== edge.shell) edge.shell.append(boundary);
    }
    const name = kind === 'header' ? '页眉' : '页脚';
    modeButton.textContent = pendingHit
      ? pendingHit.frame
        ? `编辑${name}`
        : '返回正文'
      : `${name} · 返回正文`;
    modeButton.title = pendingHit ? '双击文档区域或按此按钮切换' : '双击正文或按 Esc 返回';
    input.setAttribute(
      'aria-label',
      active ? `${active.kind === 'header' ? '页眉' : '页脚'}编辑输入` : '文档编辑输入',
    );
  }
  function setFrame(frame?: ActiveFrame): void {
    pendingHit = undefined;
    if (frame?.relId !== active?.relId) {
      binding.editor.breakHistory();
      binding = frame && bodyBinding.headerFooter ? bodyBinding.headerFooter(frame.relId) : bodyBinding;
      state = binding === bodyBinding ? bodyState : controllerFor(binding);
    }
    active = frame;
  }
  function pageFor(position: DocPosition): number | undefined {
    return active && position.nodeId.startsWith(`${active.relId}:`) ? active.page : undefined;
  }
  function activate(hit?: EditingHit): void {
    setFrame(hit?.frame && { ...hit.frame, page: hit.page });
    run(() => {
      if (hit?.position) state.select({ start: hit.position, end: hit.position });
    });
    focus();
    if (!hit && state.focus) view.scrollTo(state.focus, { align: 'nearest' });
  }
  modeButton.addEventListener('click', () => activate(pendingHit), inputEventOptions);
  modeButton.addEventListener(
    'keydown',
    (event) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        activate();
      }
    },
    inputEventOptions,
  );
  let decoration: ReturnType<DomView['decorate']> | undefined;
  let dead = false;
  let drag = false;
  let dragAnchor: DocPosition | undefined;
  let suppressCommit: string | undefined;
  let timer: ReturnType<typeof setTimeout> | undefined;
  let preferredX: number | undefined;
  const selectionListeners = new Set<(selection: DocRange | undefined) => void>();
  let lastSelection: DocRange | undefined;
  function refresh(): void {
    if (dead) return;
    showMode();
    const range = state.selection;
    // 所有改选区的路径最后都走到这里，在这一处比较就不必在每个命令后面各补一句通知
    if (!sameRange(range, lastSelection)) {
      lastSelection = range;
      for (const listener of [...selectionListeners]) listener(range);
    }
    decoration?.dispose();
    decoration = undefined;
    if (!range) {
      input.style.visibility = 'hidden';
      return;
    }
    const at = state.focus ?? range.end;
    const rect = view.caretRect(at, active?.page);
    input.style.visibility = rect ? 'visible' : 'hidden';
    input.style.opacity = doc.activeElement === input ? '1' : '0';
    if (rect) {
      input.style.left = `${rect.x}px`;
      input.style.top = `${rect.y}px`;
      input.style.height = `${rect.height}px`;
      input.style.lineHeight = input.style.height;
      if (state.composing && compositionSize) {
        const line = view.index.caretRect(at, active?.page);
        if (line?.height) input.style.fontSize = `${(compositionSize * rect.height) / line.height}px`;
      } else input.style.fontSize = `${Math.min(rect.height, 16)}px`;
      input.style.caretColor = state.composing ? 'currentColor' : 'black';
      input.style.maxWidth = `calc(100vw - ${rect.x}px)`;
      input.style.backgroundColor = state.composing ? 'white' : 'transparent';
      sizeInput();
    }
    if (doc.activeElement === input)
      decoration = view.decorate(range, {
        story: active?.kind ?? 'body',
        ...(active && { page: active.page }),
        style: { background: 'Highlight', opacity: '0.3' },
      });
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
    formatChanged?.();
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
      boundary.contains(event.target) ||
      !view.root.contains(event.target)
    )
      return;
    const hit = view.editingHit(event);
    if (!hit) return;
    event.preventDefault();
    if (hit.frame?.relId !== active?.relId) {
      drag = false;
      if (hit.frame && !bodyBinding.headerFooter) return;
      pendingHit = hit;
      showMode();
      modeButton.focus({ preventScroll: true });
      return;
    }
    pendingHit = undefined;
    if (active) active.page = hit.page;
    const at = hit.position;
    if (!at) return;
    const start = event.shiftKey ? (state.anchor ?? at) : at;
    run(() => state.select({ start, end: at }));
    drag = true;
    dragAnchor = start;
    focus();
  }
  function pointerMove(event: PointerEvent): void {
    if (!drag) return;
    const hit = view.editingHit(event);
    if (!hit || hit.frame?.relId !== active?.relId || (active && active.page !== hit.page)) return;
    const at = hit.position;
    const range = state.selection;
    if (at && range) run(() => state.select({ start: dragAnchor ?? range.start, end: at }));
  }
  function pointerEnd(): void {
    drag = false;
  }
  function doubleClick(event: MouseEvent): void {
    if (event.button !== 0 || state.composing || win === null || !(event.target instanceof win.Node)) return;
    if (boundary.contains(event.target) || !view.root.contains(event.target)) return;
    const hit = view.editingHit(event);
    if (!hit) return;
    event.preventDefault();
    drag = false;
    if (hit.frame?.relId !== active?.relId) {
      if (!hit.frame || bodyBinding.headerFooter) activate(hit);
      return;
    }
    const at = hit.position;
    if (!at) return;
    const caret = view.caretRect(at, active?.page);
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
        case 'formatSuperscript':
          state.toggleFormat('superscript');
          break;
        case 'formatSubscript':
          state.toggleFormat('subscript');
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
    if (event.key === 'Escape' && active) {
      event.preventDefault();
      activate();
      return;
    }
    // 列表段首 / 列表选区的 Tab 升降级；其余位置 Tab 插制表位（Word 的行为），
    // Shift+Tab 不接管，留给浏览器把焦点移出编辑区 —— 否则键盘用户出不去。
    if (event.key === 'Tab' && !event.ctrlKey && !event.metaKey && !event.altKey) {
      // 单元格里 Tab / Shift+Tab 跳格（Word 的行为），优先于列表升降级；Shift+Tab 在格里因此也接管，
      // 焦点仍能从表格外面的段落移出去。Word 在格里插制表位用 Ctrl+Tab，浏览器截走它（切标签页），没有接。
      let moved = false;
      run(() => {
        moved = state.moveCell(event.shiftKey ? 'backward' : 'forward');
      });
      if (moved) {
        event.preventDefault();
        const at = state.focus;
        if (at) view.scrollTo(at, { align: 'nearest', ...(active && { page: active.page }) });
        return;
      }
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
    } else if (
      (event.ctrlKey || event.metaKey) &&
      event.altKey &&
      !event.shiftKey &&
      event.code in HEADING_CODES
    ) {
      const name = HEADING_CODES[event.code] as BuiltinStyleName;
      action = () => state.applyBuiltinStyle(name);
    }
    // Word 的 Ctrl+Shift+N 回到正文样式。Windows / Linux 的 Chrome 截走它（隐身窗口），只在 Mac 上到得了页面。
    else if (
      event.ctrlKey &&
      !event.metaKey &&
      !event.altKey &&
      event.shiftKey &&
      event.key.toLowerCase() === 'n'
    )
      action = () => state.applyBuiltinStyle('normal');
    // Word 的 Ctrl+Shift+L 是「列表项目符号」样式；样式不一定存在，这里直接套项目符号列表。
    else if (mod && !event.altKey && event.shiftKey && event.key.toLowerCase() === 'l')
      action = () => state.toggleList('bullet');
    // Word 的 Ctrl+Enter（Mac 上 Cmd+Return）插分页符；Ctrl+Shift+Enter 是分栏符，单栏文档里没有意义，不接管。
    else if (mod && !event.altKey && !event.shiftKey && event.key === 'Enter')
      action = () => state.pageBreak();
    // 只认 Ctrl：Cmd+Q 是系统退出应用，Cmd+Space 是系统搜索。
    else if (event.ctrlKey && !event.metaKey && !event.altKey && !event.shiftKey && event.key === ' ')
      action = () => state.clearRunFormat();
    else if (
      event.ctrlKey &&
      !event.metaKey &&
      !event.altKey &&
      !event.shiftKey &&
      event.key.toLowerCase() === 'q'
    )
      action = () => state.clearParagraphFormat();
    else if (mod && !event.altKey && !event.shiftKey && event.key.toLowerCase() === 'a')
      action = () => state.selectAll();
    // Equal 的 key 在 Shift 下通常变成 +，按物理键认；合成事件没有 code 时按 key 兜底。
    else if (
      mod &&
      !event.altKey &&
      (event.code === 'Equal' || (!event.code && (event.key === '=' || event.key === '+')))
    )
      action = () => state.toggleFormat(event.shiftKey ? 'superscript' : 'subscript');
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
            active,
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
      if (at) view.scrollTo(at, { align: 'nearest', ...(active && { page: active.page }) });
    }
  }
  function compositionStart(): void {
    suppressCommit = undefined;
    run(() => {
      state.compositionStart();
      const range = state.selection;
      const format = range && binding.format?.(range, state.pendingFormat)[0];
      compositionSize = format?.size;
      input.style.fontFamily = fontFamily(format?.fonts?.ascii || format?.fonts?.eastAsia || '');
      input.style.fontWeight = format?.bold ? 'bold' : 'normal';
      input.style.fontStyle = format?.italic ? 'italic' : 'normal';
      input.style.color = cssColor(format?.color ?? 'auto');
    });
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
  const unsubscribe = bodyBinding.subscribe((change) => {
    if (state !== bodyState) bodyState.apply(change);
    pendingHit = undefined;
    preferredX = undefined;
    state.apply(change);
    refresh();
  });
  input.addEventListener(
    'input',
    () => {
      if (!state.composing) input.value = '';
      else if (!autoWidth) sizeInput();
    },
    inputEventOptions,
  );
  input.addEventListener('beforeinput', beforeInput, inputEventOptions);
  input.addEventListener('keydown', keydown, inputEventOptions);
  input.addEventListener('compositionstart', compositionStart, inputEventOptions);
  input.addEventListener('compositionend', compositionEnd, inputEventOptions);
  input.addEventListener('paste', paste, inputEventOptions);
  input.addEventListener('copy', clipboard, inputEventOptions);
  input.addEventListener('cut', clipboard, inputEventOptions);
  input.addEventListener('focus', refresh, inputEventOptions);
  input.addEventListener('blur', blur, inputEventOptions);
  container.addEventListener('pointerdown', pointer as EventListener);
  container.addEventListener('dblclick', doubleClick as EventListener);
  doc.addEventListener('pointermove', pointerMove);
  doc.addEventListener('pointerup', pointerEnd);
  doc.addEventListener('pointercancel', pointerEnd);
  doc.addEventListener('scroll', refresh, true);
  win?.addEventListener('resize', refresh);
  refresh();
  return {
    formatting: {
      read() {
        if (dead) return undefined;
        try {
          return selectionFormatting(state, binding);
        } catch (error) {
          // 修订视图可以隐藏整行。模型选区仍在，当前显示树却没有可查询的端点。
          if (error instanceof RangeError) return undefined;
          throw error;
        }
      },
      execute(command) {
        if (!dead && !state.composing) run(() => applyFormatting(state, command));
      },
    },
    get selection() {
      return state.selection;
    },
    pageFor,
    select(range) {
      preferredX = undefined;
      if (pageFor(range.start) === undefined) {
        let frame: ActiveFrame | undefined;
        // 宿主也可以传回 locate() 给的部件位置。没有点击页时选第一次出现的副本。
        const page = range.start.nodeId.includes(':') ? view.index.caretRect(range.start)?.page : undefined;
        if (page !== undefined && bodyBinding.headerFooter) {
          for (const kind of ['header', 'footer'] as const) {
            const edge = view.frameBoundary(page, kind);
            if (edge && range.start.nodeId.startsWith(`${edge.relId}:`))
              frame = { kind, relId: edge.relId, page };
          }
        }
        setFrame(frame);
      }
      state.select(range);
      refresh();
      formatChanged?.();
    },
    focus,
    refresh,
    onSelectionChange(listener) {
      selectionListeners.add(listener);
      return () => {
        selectionListeners.delete(listener);
      };
    },
    dispose() {
      dead = true;
      inputEvents.abort();
      selectionListeners.clear();
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
      boundary.remove();
    },
  };
}
