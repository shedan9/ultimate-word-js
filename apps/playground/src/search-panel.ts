/** 查找位置来自当前模型；编辑、历史跳转与重挂后都重查，避免拿旧 run 范围替换新文字。 */
import type { DecorationHandle, Disposable, DocRange, UwDocument, UwView } from 'ultimate-word';
import './search-panel.css';

export function createSearchPanel(button: HTMLButtonElement, host: HTMLElement) {
  const panel = document.createElement('section');
  panel.id = 'search-panel';
  panel.className = 'search-panel';
  panel.hidden = true;
  panel.setAttribute('aria-label', '查找与替换');
  panel.innerHTML = `
    <div class="search-row">
      <label for="search-query">查找</label><input id="search-query" type="search" placeholder="输入要查找的文字">
      <div class="search-actions"><button type="button" class="search-prev">上一处</button><button type="button" class="search-next">下一处</button></div>
      <label class="search-case"><input type="checkbox">区分大小写</label>
      <button type="button" class="search-close">关闭</button>
    </div>
    <div class="search-row">
      <label for="search-replacement">替换为</label><input id="search-replacement" type="text" placeholder="留空则删除匹配文字">
      <div class="search-actions"><button type="button" class="search-one">替换当前</button><button type="button" class="search-all">全部替换</button></div>
      <div class="search-actions"><button type="button" class="search-undo">撤销</button><button type="button" class="search-redo">重做</button></div>
    </div>
    <p class="search-summary" role="status" aria-live="polite" aria-atomic="true"></p>
    <p class="search-help">Enter 下一处，Shift+Enter 上一处，Esc 关闭。替换仅在编辑模式可用。</p>`;
  host.append(panel);
  button.setAttribute('aria-controls', panel.id);
  button.setAttribute('aria-expanded', 'false');
  const query = panel.querySelector('#search-query') as HTMLInputElement;
  const replacement = panel.querySelector('#search-replacement') as HTMLInputElement;
  const matchCase = panel.querySelector('.search-case input') as HTMLInputElement;
  const summary = panel.querySelector('.search-summary') as HTMLElement;
  const one = panel.querySelector('.search-one') as HTMLButtonElement;
  const all = panel.querySelector('.search-all') as HTMLButtonElement;
  const previous = panel.querySelector('.search-prev') as HTMLButtonElement;
  const next = panel.querySelector('.search-next') as HTMLButtonElement;
  const undo = panel.querySelector('.search-undo') as HTMLButtonElement;
  const redo = panel.querySelector('.search-redo') as HTMLButtonElement;
  let doc: UwDocument | undefined;
  let view: UwView | undefined;
  let editable = false;
  let subscription: Disposable | undefined;
  let ranges: DocRange[] = [];
  let cursor = 0;
  let marks: DecorationHandle[] = [];
  let current: DecorationHandle | undefined;
  let applying = false;

  const clear = () => {
    for (const mark of marks) mark.dispose();
    marks = [];
    current?.dispose();
    current = undefined;
  };
  const notify = (text = '', error = false) => {
    const count = !query.value
      ? '请输入查找文字。'
      : ranges.length
        ? `${cursor + 1} / ${ranges.length} 处`
        : '没有匹配结果。';
    summary.textContent = [text, count].filter(Boolean).join(' · ');
    summary.classList.toggle('search-error', error);
  };
  const controls = () => {
    previous.disabled = next.disabled = ranges.length === 0;
    one.disabled = all.disabled = !editable || ranges.length === 0;
    replacement.disabled = !editable;
    undo.disabled = !editable || !doc?.canUndo;
    redo.disabled = !editable || !doc?.canRedo;
  };
  const highlight = (scroll = false) => {
    current?.dispose();
    current = undefined;
    const range = ranges[cursor];
    if (view !== undefined && range !== undefined && !panel.hidden) {
      current = view.decorate(range, { className: 'hit hit-current' });
      if (scroll) view.scrollTo(range, { align: 'center' });
    }
    controls();
    notify();
  };
  const refresh = () => {
    clear();
    ranges =
      doc !== undefined && !panel.hidden && query.value
        ? doc.find(query.value, { matchCase: matchCase.checked })
        : [];
    cursor = Math.min(cursor, Math.max(0, ranges.length - 1));
    // 长文只铺前 500 处普通底色，导航与全部替换仍使用完整结果，当前项始终单独着色。
    const mounted = view;
    if (mounted !== undefined)
      marks = ranges.slice(0, 500).map((range) => mounted.decorate(range, { className: 'hit' }));
    highlight();
  };
  const close = () => {
    panel.hidden = true;
    clear();
    button.setAttribute('aria-expanded', 'false');
    button.focus();
  };
  button.addEventListener('click', () => {
    if (!panel.hidden) return close();
    panel.hidden = false;
    button.setAttribute('aria-expanded', 'true');
    refresh();
    query.focus();
    // 窄屏上展开的表单比入口更重要，先把工具栏滚到表单顶部，避免大半空间仍被开关占着。
    if (host.scrollHeight > host.clientHeight) host.scrollTop = panel.offsetTop;
  });
  panel.querySelector('.search-close')?.addEventListener('click', close);
  const go = (step: number) => {
    if (!ranges.length) return;
    cursor = (cursor + step + ranges.length) % ranges.length;
    highlight(true);
  };
  previous.addEventListener('click', () => go(-1));
  next.addEventListener('click', () => go(1));
  query.addEventListener('input', () => {
    cursor = 0;
    refresh();
  });
  matchCase.addEventListener('change', () => {
    cursor = 0;
    refresh();
  });
  panel.addEventListener('keydown', (event) => {
    if (event.isComposing || event.keyCode === 229) return;
    if (event.key === 'Escape') {
      event.preventDefault();
      close();
    } else if (event.key === 'Enter' && event.target === query) {
      event.preventDefault();
      go(event.shiftKey ? -1 : 1);
    }
  });
  one.addEventListener('click', () => {
    const range = ranges[cursor];
    if (!editable || doc === undefined || range === undefined) return;
    try {
      applying = true;
      let inserted: DocRange | undefined;
      doc.tx((tx) => {
        inserted = tx.replaceText(range, replacement.value);
      });
      refresh();
      // 从新文字之后继续找，替换串含查找串时也不会反复替换刚插入的同一处。
      if (inserted !== undefined) {
        const end = inserted.end;
        const opened = doc;
        const after = ranges.findIndex((item) => opened.compare(item.start, end) >= 0);
        cursor = after < 0 ? 0 : after;
      }
      highlight(true);
      notify('已替换 1 处');
    } catch (error) {
      controls();
      notify(`无法替换：${error instanceof Error ? error.message : String(error)}`, true);
    } finally {
      applying = false;
    }
  });
  all.addEventListener('click', () => {
    if (!editable || doc === undefined || !query.value) return;
    try {
      applying = true;
      const result = doc.replaceAll(query.value, replacement.value, { matchCase: matchCase.checked });
      cursor = 0;
      refresh();
      notify(`已替换 ${result.replaced} 处，跳过 ${result.skipped} 处受保护内容`);
    } catch (error) {
      controls();
      notify(`无法替换：${error instanceof Error ? error.message : String(error)}`, true);
    } finally {
      applying = false;
    }
  });
  undo.addEventListener('click', () => {
    if (editable) doc?.undo();
  });
  redo.addEventListener('click', () => {
    if (editable) doc?.redo();
  });

  return {
    attach(opened: UwDocument, mounted: UwView, edit: boolean) {
      subscription?.dispose();
      clear();
      if (doc !== opened) {
        panel.hidden = true;
        cursor = 0;
        button.setAttribute('aria-expanded', 'false');
      }
      doc = opened;
      view = mounted;
      editable = edit;
      button.disabled = false;
      subscription = opened.on('document:change', () => {
        if (!applying) refresh();
      });
      refresh();
    },
    refresh,
  };
}
