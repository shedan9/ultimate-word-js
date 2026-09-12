import { TWIP_PER_PT, twipsToPx } from '@uw/core';
import type { DocumentLayout, IndexedLine, PageLayout } from '@uw/layout';
import type { DocPosition, DocRange } from '@uw/model';
import type { MountOptions } from '@uw/render-dom/dom';
import { renderPage, toDom } from '@uw/render-dom/dom';
import { pageRect } from './annotations.ts';
import type {
  DecorationHandle,
  DecorationOptions,
  OverlayHandle,
  OverlayOptions,
} from './annotations-dom.ts';
import { createAnnotations } from './annotations-dom.ts';
import { selectedText } from './copy.ts';
import type { ScrollOptions, ScrollTarget } from './scroll.ts';
import { scrollTargetRect } from './scroll.ts';
import { buildTextLayer } from './text-layer.ts';
import type { ClientPoint, ClientRect, PageViewport } from './transform.ts';
import { createReadonlyView } from './view.ts';
import { pagesToRender } from './virtual-pages.ts';

export type { OverlayPlacement } from './annotations.ts';
export type {
  DecorationHandle,
  DecorationOptions,
  OverlayHandle,
  OverlayOptions,
} from './annotations-dom.ts';
export type { ScrollOptions, ScrollTarget } from './scroll.ts';

export interface ViewOptions extends MountOptions {
  /** 默认开启；常驻原生文字层，不因绘制页卸载而丢失选区 / 浏览器查找。 */
  textLayer?: boolean;
  /** 默认开启。没有 IntersectionObserver 时降级为全量绘制。 */
  virtualize?: boolean;
  /** 可见页前后各预绘制几页，默认 2。 */
  overscan?: number;
  /** 页间距，CSS px，默认 24。 */
  pageGap?: number;
}

export interface DomView {
  readonly root: HTMLElement;
  locate(point: ClientPoint): DocPosition | null;
  rectsOf(range: DocRange): ClientRect[];
  caretRect(position: DocPosition): ClientRect | null;
  decorate(range: DocRange, options?: DecorationOptions): DecorationHandle;
  overlay(position: DocPosition, element: HTMLElement, options?: OverlayOptions): OverlayHandle;
  /**
   * 滚到一个模型位置 / range 的首行 / 某一页。目标排不出来（空 run、越界的页号）时不动并返回 false。
   * 走的是 `scrollIntoView`，所以窗口与任意祖先滚动容器都照顾到，宿主不必告诉视图谁在滚。
   */
  scrollTo(target: ScrollTarget, options?: ScrollOptions): boolean;
  /** 只改页面占位尺寸，保留文字层、选区与已绘制的内容。 */
  setZoom(zoom: number): void;
  update(layout: DocumentLayout, options?: ViewOptions): void;
  destroy(): void;
}

interface PageSlot {
  page: PageLayout;
  shell: HTMLDivElement;
  painting: HTMLDivElement;
  viewport: SVGSVGElement;
}

function validate(options: ViewOptions): void {
  const zoom = options.zoom ?? 1;
  const overscan = options.overscan ?? 2;
  const gap = options.pageGap ?? 24;
  if (!Number.isFinite(zoom) || zoom <= 0) throw new RangeError('zoom 必须是大于 0 的有限数');
  if (!Number.isSafeInteger(overscan) || overscan < 0) throw new RangeError('overscan 必须是非负整数');
  if (!Number.isFinite(gap) || gap < 0) throw new RangeError('pageGap 必须是非负有限数');
}

/**
 * 每页保留尺寸固定的壳与原生文字 SVG，只按需挂载复杂的绘制 SVG。
 * 文字层不能用 content-visibility:hidden，也不能一起卸载，否则 Ctrl+F 与跨页选区会缺字。
 * 页面几何查询使用常驻 SVG 的屏幕矩阵；它已包含滚动，不能重复加 scrollTop。
 */
export function mountView(container: Element, layout: DocumentLayout, options: ViewOptions = {}): DomView {
  let opts = { ...options, document: options.document ?? container.ownerDocument };
  validate(opts);
  const doc = container.ownerDocument;
  const win = doc.defaultView;
  let destroyed = false;
  let observer: IntersectionObserver | undefined;
  let visible = new Set<number>();
  let printing = false;
  let selecting = false;
  let root: HTMLDivElement;
  let slots: PageSlot[] = [];
  let view = createReadonlyView(layout, measurePages);

  function makeRoot(next: DocumentLayout, settings: ViewOptions, lines: readonly IndexedLine[]) {
    const document = settings.document ?? doc;
    const prefix = settings.classPrefix ?? 'uw';
    const nextRoot = document.createElement('div');
    nextRoot.className = `${prefix}-doc`;
    nextRoot.setAttribute('role', 'document');
    nextRoot.setAttribute('aria-label', '文档预览');
    nextRoot.tabIndex = 0;
    nextRoot.style.cssText = `position:relative;isolation:isolate;display:flex;flex-direction:column;align-items:center;gap:${settings.pageGap ?? 24}px`;
    // 先分桶，不能在每一页上重新扫描整份文档的行表。
    const byPage = new Map<number, IndexedLine[]>();
    for (const line of lines) {
      const bucket = byPage.get(line.page);
      if (bucket === undefined) byPage.set(line.page, [line]);
      else bucket.push(line);
    }
    const nextSlots = next.pages.map((page): PageSlot => {
      const shell = document.createElement('div');
      shell.className = `${prefix}-page-shell${page.filler === true ? ` ${prefix}-page-filler` : ''}`;
      shell.dataset.page = String(page.index);
      shell.setAttribute('role', 'group');
      shell.setAttribute('aria-label', `第 ${page.index + 1} 页，页码 ${page.number}`);
      shell.style.cssText = 'position:relative;flex:none;background:white';
      const painting = document.createElement('div');
      painting.className = `${prefix}-painting`;
      painting.style.cssText = 'position:absolute;inset:0;pointer-events:none;user-select:none';
      // aria-hidden / user-select:none 不会阻止 Ctrl+F，inert 才能消除两层文字的重复匹配。
      painting.inert = true;
      const viewport = toDom(
        buildTextLayer(page, byPage.get(page.index) ?? [], settings, settings.textLayer !== false),
        document,
      ) as SVGSVGElement;
      shell.append(painting, viewport);
      nextRoot.append(shell);
      return { page, shell, painting, viewport };
    });
    sizePages(nextSlots, settings.zoom ?? 1);
    return { root: nextRoot, slots: nextSlots };
  }

  function sizePages(pages: PageSlot[], zoom: number): void {
    for (const { page, shell } of pages) {
      shell.style.width = `${twipsToPx(page.geometry.width, zoom)}px`;
      shell.style.height = `${twipsToPx(page.geometry.height, zoom)}px`;
    }
  }

  function paint(wanted: Set<number>): void {
    for (let i = 0; i < slots.length; i++) {
      const slot = slots[i] as PageSlot;
      if (wanted.has(i)) {
        if (slot.painting.firstChild !== null) continue;
        const svg = renderPage(slot.page, opts) as SVGSVGElement;
        svg.style.cssText = 'display:block;width:100%;height:100%';
        slot.painting.append(svg);
        slot.shell.dataset.rendered = 'true';
      } else if (slot.painting.firstChild !== null) {
        slot.painting.replaceChildren();
        slot.shell.dataset.rendered = 'false';
      }
    }
  }

  function refreshPaint(): void {
    const all = printing || opts.virtualize === false || observer === undefined;
    paint(all ? new Set(slots.map((_, i) => i)) : pagesToRender(slots.length, visible, opts.overscan ?? 2));
  }

  function observe(): void {
    observer?.disconnect();
    observer = undefined;
    visible = new Set();
    if (opts.virtualize !== false && win?.IntersectionObserver !== undefined) {
      const indices = new Map<Element, number>(slots.map((slot, index) => [slot.shell, index]));
      const activeRoot = root;
      observer = new win.IntersectionObserver((entries) => {
        if (destroyed || root !== activeRoot) return;
        for (const entry of entries) {
          const index = indices.get(entry.target);
          if (index === undefined) continue;
          if (entry.isIntersecting) visible.add(index);
          else visible.delete(index);
        }
        refreshPaint();
      });
      // root:null 让浏览器同时考虑窗口与所有祖先滚动容器的裁剪。
      for (const slot of slots) observer.observe(slot.shell);
    }
    refreshPaint();
  }

  function measurePages(): PageViewport[] {
    if (destroyed || !root.isConnected) return [];
    return slots.flatMap(({ page, viewport }) => {
      const m = viewport.getScreenCTM();
      if (m === null || viewport.getClientRects().length === 0) return [];
      return [
        {
          page: page.index,
          width: page.geometry.width,
          height: page.geometry.height,
          matrix: {
            a: m.a / TWIP_PER_PT,
            b: m.b / TWIP_PER_PT,
            c: m.c / TWIP_PER_PT,
            d: m.d / TWIP_PER_PT,
            e: m.e,
            f: m.f,
          },
        },
      ];
    });
  }

  function onCopy(event: ClipboardEvent): void {
    if (destroyed || opts.textLayer === false || event.defaultPrevented || event.clipboardData === null)
      return;
    // 批注输入框有自己的选区，不能用文档上残留的选区覆盖它。
    if (doc.activeElement?.closest('[data-overlay]')) return;
    const selection = doc.getSelection();
    if (selection === null) return;
    const text = selectedText(root, selection);
    if (text === undefined) return;
    event.clipboardData.setData('text/plain', text);
    event.preventDefault();
  }
  function onPointerDown(event: PointerEvent): void {
    selecting = win !== null && event.target instanceof win.Node && root.contains(event.target);
  }
  function onPointerEnd(): void {
    selecting = false;
  }
  function onSelectionChange(): void {
    if (destroyed || selecting || opts.textLayer === false) return;
    const selection = doc.getSelection();
    if (selection === null || selection.isCollapsed || selection.rangeCount !== 1) return;
    const range = selection.getRangeAt(0);
    // 原生查找对 SVG 会建立选区，却不一定滚动祖先容器。只照顾同一行内的选区：
    // 跨页全选不滚到文末，鼠标拖拽期间也不抢走浏览器的自动滚动。
    const firstLine = range.startContainer.parentElement?.closest('[data-copy-line]');
    const lastLine = range.endContainer.parentElement?.closest('[data-copy-line]');
    if (firstLine === undefined || firstLine === null || firstLine !== lastLine) return;
    const text = range.startContainer.parentElement;
    const layer = text?.closest('[data-text-layer="true"]');
    if (text === null || text === undefined || layer === null || layer === undefined || !root.contains(layer))
      return;
    const index = slots.findIndex((slot) => slot.viewport === layer);
    if (index < 0) return;
    text.scrollIntoView({ block: 'nearest', inline: 'nearest', behavior: 'instant' });
  }
  function beforePrint(): void {
    printing = true;
    refreshPaint();
  }
  function afterPrint(): void {
    printing = false;
    refreshPaint();
  }

  const initial = makeRoot(layout, opts, view.lines);
  root = initial.root;
  slots = initial.slots;
  container.replaceChildren(root);
  observe();
  const annotations = createAnnotations(() => ({ root, pages: slots, index: view.index }), doc);
  annotations.rebind();
  doc.addEventListener('copy', onCopy);
  doc.addEventListener('selectionchange', onSelectionChange);
  doc.addEventListener('pointerdown', onPointerDown);
  doc.addEventListener('pointerup', onPointerEnd);
  doc.addEventListener('pointercancel', onPointerEnd);
  win?.addEventListener('beforeprint', beforePrint);
  win?.addEventListener('afterprint', afterPrint);

  function assertLive(): void {
    if (destroyed) throw new Error('视图已销毁');
  }
  return {
    get root() {
      return root;
    },
    locate(point) {
      if (destroyed || ![point.clientX, point.clientY].every(Number.isFinite)) return null;
      const target = doc.elementFromPoint(point.clientX, point.clientY);
      if (target === null || !root.contains(target)) return null;
      return view.locate(point);
    },
    rectsOf: (range) => view.rectsOf(range),
    caretRect: (position) => view.caretRect(position),
    decorate: annotations.decorate,
    overlay: annotations.overlay,
    scrollTo(target, options = {}) {
      assertLive();
      const scroll: ScrollIntoViewOptions = {
        block: options.align ?? 'start',
        inline: 'nearest',
        behavior: options.behavior ?? 'auto',
      };
      if ('page' in target) {
        const slot = Number.isInteger(target.page) ? slots[target.page] : undefined;
        if (slot === undefined) return false;
        slot.shell.scrollIntoView(scroll);
        return true;
      }
      const rect = scrollTargetRect(view.index, target);
      const slot = rect === undefined ? undefined : slots[rect.page];
      if (rect === undefined || slot === undefined) return false;
      // 壳内坐标与装饰同一套换算（壳可能被宿主改过尺寸，不能按 zoom 反推）
      const style = win?.getComputedStyle(slot.viewport);
      const width = Number.parseFloat(style?.width ?? '') || slot.viewport.clientWidth;
      const height = Number.parseFloat(style?.height ?? '') || slot.viewport.clientHeight;
      if (width <= 0 || height <= 0) return false;
      const local = pageRect(rect, slot.page.geometry.width, slot.page.geometry.height, width, height);
      // 借一个临时元素让浏览器算滚动量：它认得所有祖先滚动容器，自己算只能认一个。
      // smooth 滚动的终点在调用那一刻就定了，元素随后移除不影响它
      const probe = doc.createElement('div');
      probe.style.cssText = `position:absolute;left:${local.x}px;top:${local.y}px;width:${Math.max(local.width, 1)}px;height:${Math.max(local.height, 1)}px;pointer-events:none;visibility:hidden`;
      slot.shell.append(probe);
      probe.scrollIntoView(scroll);
      probe.remove();
      return true;
    },
    setZoom(zoom) {
      assertLive();
      validate({ ...opts, zoom });
      opts = { ...opts, zoom };
      sizePages(slots, zoom);
      annotations.refresh();
      // 壳尺寸变化后 IntersectionObserver 重新报告可见页，无需重建文字节点或索引。
    },
    update(next, nextOptions = {}) {
      assertLive();
      const merged = { ...opts, ...nextOptions };
      validate(merged);
      // 先构建新索引与 DOM，调用方的字体解析器抛错时保留旧视图。
      const nextView = createReadonlyView(next, measurePages);
      const nextState = makeRoot(next, merged, nextView.lines);
      const focused = doc.activeElement;
      const restoreFocus =
        win !== null &&
        focused instanceof win.HTMLElement &&
        focused.closest('[data-overlay]') !== null &&
        root.contains(focused);
      observer?.disconnect();
      root.replaceWith(nextState.root);
      root = nextState.root;
      slots = nextState.slots;
      opts = merged;
      view = nextView;
      observe();
      annotations.rebind();
      if (restoreFocus && focused.isConnected) focused.focus({ preventScroll: true });
    },
    destroy() {
      if (destroyed) return;
      destroyed = true;
      annotations.destroy();
      observer?.disconnect();
      observer = undefined;
      visible.clear();
      doc.removeEventListener('copy', onCopy);
      doc.removeEventListener('selectionchange', onSelectionChange);
      doc.removeEventListener('pointerdown', onPointerDown);
      doc.removeEventListener('pointerup', onPointerEnd);
      doc.removeEventListener('pointercancel', onPointerEnd);
      win?.removeEventListener('beforeprint', beforePrint);
      win?.removeEventListener('afterprint', afterPrint);
      root.replaceChildren();
      root.remove();
      slots = [];
      view.update({ pages: [] });
    },
  };
}
