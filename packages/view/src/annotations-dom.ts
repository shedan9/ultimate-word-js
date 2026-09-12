import type { LayoutIndex, PageLayout } from '@uw/layout';
import type { DocPosition, DocRange } from '@uw/model';
import type { OverlayPlacement } from './annotations.ts';
import { anchorRect, pageRect, placeOverlay } from './annotations.ts';

export interface DecorationOptions {
  className?: string;
  /** 纯视觉 CSS，例如 background / borderBottom；定位由视图管理。 */
  style?: Readonly<Record<string, string>>;
  /** 当前只支持文字上方的装饰，使用透明背景避免遮住字形。 */
  layer?: 'above-text';
}
export interface OverlayOptions {
  placement?: OverlayPlacement;
  /** 页面壳内 CSS px，默认 { x: 8, y: 0 }；不随文档 zoom 放大。 */
  offset?: { x: number; y: number };
}
export interface DecorationHandle {
  update(range?: DocRange): void;
  dispose(): void;
}
export interface OverlayHandle {
  update(position?: DocPosition): void;
  dispose(): void;
}
export interface AnnotationPage {
  page: PageLayout;
  shell: HTMLDivElement;
  viewport: SVGSVGElement;
}
interface Context {
  root: HTMLElement;
  pages: readonly AnnotationPage[];
  index: LayoutIndex;
}

/**
 * 装饰属于页壳，不属于可被虚拟化卸载的绘制层。滚动和 CSS 变换自然继承，
 * 不在 scroll 回调里逐个搬动批注。只在布局 / 尺寸变化时重算页面内坐标。
 */
export function createAnnotations(context: () => Context, document: Document) {
  const win = document.defaultView;
  let dead = false;
  let frame: number | undefined;
  const decorations = new Set<{ range: DocRange; options: DecorationOptions; nodes: HTMLDivElement[] }>();
  const overlays = new Set<{
    position: DocPosition;
    options: OverlayOptions;
    wrapper: HTMLDivElement;
    element: HTMLElement;
  }>();
  const observer = win?.ResizeObserver === undefined ? undefined : new win.ResizeObserver(schedule);

  function assertLive() {
    if (dead) throw new Error('视图已销毁');
  }
  function schedule() {
    if (dead || frame !== undefined || win === null) return;
    frame = win.requestAnimationFrame(() => {
      frame = undefined;
      refresh();
    });
  }
  function refresh() {
    if (dead) return;
    const { pages, index } = context();
    // 先读全部几何，再写样式，避免逐条高亮触发同步布局。
    const geometry = new Map(
      pages.map((slot) => {
        const style = win?.getComputedStyle(slot.viewport);
        return [
          slot.page.index,
          {
            slot,
            width: Number.parseFloat(style?.width ?? '') || slot.viewport.clientWidth,
            height: Number.parseFloat(style?.height ?? '') || slot.viewport.clientHeight,
          },
        ];
      }),
    );
    const local = (rect: ReturnType<LayoutIndex['caretRect']>) => {
      if (rect === undefined) return undefined;
      const page = geometry.get(rect.page);
      if (page === undefined || page.width <= 0 || page.height <= 0) return undefined;
      return {
        slot: page.slot,
        rect: pageRect(
          rect,
          page.slot.page.geometry.width,
          page.slot.page.geometry.height,
          page.width,
          page.height,
        ),
      };
    };
    const decorationJobs = [...decorations].map((item) => ({
      item,
      rects: index.rectsOf(item.range).flatMap((rect) => {
        const at = local(rect);
        return at === undefined ? [] : [at];
      }),
    }));
    const overlayJobs = [...overlays].map((item) => {
      const placement = item.options.placement ?? 'right-of-line';
      const at = local(anchorRect(index, item.position, placement));
      return {
        item,
        at,
        point:
          at === undefined
            ? undefined
            : placeOverlay(
                at.rect,
                { width: item.wrapper.offsetWidth, height: item.wrapper.offsetHeight },
                placement,
                item.options.offset ?? { x: 8, y: 0 },
              ),
      };
    });
    for (const { item, rects } of decorationJobs) {
      for (let i = 0; i < rects.length; i++) {
        const at = rects[i];
        if (at === undefined) continue;
        let node = item.nodes[i];
        if (node === undefined) {
          node = document.createElement('div');
          node.dataset.decoration = '';
          node.setAttribute('aria-hidden', 'true');
          node.className = item.options.className ?? '';
          Object.assign(node.style, item.options.style);
          // 强制无交互，透明高亮不能拦住原生选区或进入 Tab 顺序。
          Object.assign(node.style, {
            position: 'absolute',
            pointerEvents: 'none',
            userSelect: 'none',
            boxSizing: 'border-box',
            zIndex: '1',
          });
          item.nodes.push(node);
        }
        if (node.parentElement !== at.slot.shell) at.slot.shell.append(node);
        Object.assign(node.style, {
          left: `${at.rect.x}px`,
          top: `${at.rect.y}px`,
          width: `${at.rect.width}px`,
          height: `${at.rect.height}px`,
        });
      }
      for (const node of item.nodes.splice(rects.length)) node.remove();
    }
    for (const { item, at, point } of overlayJobs) {
      if (at === undefined || point === undefined) {
        item.wrapper.style.visibility = 'hidden';
        item.wrapper.inert = true;
        continue;
      }
      // 只移动包装节点，React 挂载点与用户输入状态保持不变。
      if (item.wrapper.parentElement !== at.slot.shell) at.slot.shell.append(item.wrapper);
      item.wrapper.inert = false;
      Object.assign(item.wrapper.style, { visibility: '', left: `${point.x}px`, top: `${point.y}px` });
    }
  }
  function rebind() {
    if (dead) return;
    observer?.disconnect();
    for (const { shell } of context().pages) observer?.observe(shell);
    for (const item of overlays) observer?.observe(item.wrapper);
    refresh();
  }
  win?.addEventListener('resize', schedule);

  return {
    refresh,
    rebind,
    decorate(range: DocRange, options: DecorationOptions = {}): DecorationHandle {
      assertLive();
      if (options.layer !== undefined && options.layer !== 'above-text')
        throw new RangeError('仅支持 above-text 装饰');
      const item = {
        range: { start: { ...range.start }, end: { ...range.end } },
        options: { ...options, style: { ...options.style } },
        nodes: [] as HTMLDivElement[],
      };
      decorations.add(item);
      refresh();
      return {
        update(next = item.range) {
          if (!decorations.has(item)) return;
          item.range = { start: { ...next.start }, end: { ...next.end } };
          refresh();
        },
        dispose() {
          if (!decorations.delete(item)) return;
          for (const node of item.nodes) node.remove();
          item.nodes = [];
        },
      };
    },
    overlay(position: DocPosition, element: HTMLElement, options: OverlayOptions = {}): OverlayHandle {
      assertLive();
      const placement = options.placement ?? 'right-of-line';
      if (!['right-of-line', 'above', 'below', 'inline'].includes(placement))
        throw new RangeError('未知 overlay placement');
      if (![options.offset?.x ?? 8, options.offset?.y ?? 0].every(Number.isFinite))
        throw new RangeError('overlay offset 必须是有限数');
      if (
        element.ownerDocument !== document ||
        element.contains(context().root) ||
        context().root.contains(element) ||
        [...overlays].some(
          (item) =>
            item.element === element || item.element.contains(element) || element.contains(item.element),
        )
      ) {
        throw new Error('overlay 必须使用本 Document 内独立的元素');
      }
      const wrapper = document.createElement('div');
      wrapper.dataset.overlay = '';
      wrapper.style.cssText =
        'position:absolute;z-index:3;width:max-content;pointer-events:auto;visibility:hidden';
      wrapper.append(element);
      // 先挂载才能测到内容尺寸；缺失锚点仍保留 DOM，等待后续重排恢复。
      context().root.append(wrapper);
      const item = {
        position: { ...position },
        element,
        wrapper,
        options: { ...options, ...(options.offset === undefined ? {} : { offset: { ...options.offset } }) },
      };
      overlays.add(item);
      observer?.observe(wrapper);
      refresh();
      return {
        update(next = item.position) {
          if (!overlays.has(item)) return;
          item.position = { ...next };
          refresh();
        },
        dispose() {
          if (!overlays.delete(item)) return;
          observer?.unobserve(wrapper);
          wrapper.remove();
          element.remove();
        },
      };
    },
    destroy() {
      if (dead) return;
      dead = true;
      observer?.disconnect();
      if (frame !== undefined) win?.cancelAnimationFrame(frame);
      win?.removeEventListener('resize', schedule);
      for (const item of decorations) for (const node of item.nodes) node.remove();
      for (const item of overlays) {
        item.wrapper.remove();
        item.element.remove();
      }
      decorations.clear();
      overlays.clear();
    },
  };
}
