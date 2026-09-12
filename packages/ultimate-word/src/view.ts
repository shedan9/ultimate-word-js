/**
 * `UwView` —— 一次屏幕呈现（api.md §5 / §7 / §8）。
 *
 * 它是 `@uw/view/dom` 的 `mountView()` 外面薄薄一层，只加三样东西：
 * ① `zoom` 认 `'fit-width'` / `'fit-page'`（要看容器尺寸，那是门面才知道的事）；
 * ② fit 模式下容器变宽变窄时自动重算（`ResizeObserver`）—— 缩放永不重排（架构 §4.1），
 *    所以跟着容器走是 O(1) 的，没有理由让调用方自己监听；
 * ③ `dispose()` 而不是 `destroy()` —— 凡是「挂上去」的东西都按 `Disposable` 回收（api.md §2 ④）。
 *
 * 底层 `DomView` 的其余方法原样透出，**不另起一套签名**：装饰、overlay、滚动的语义
 * 都是那一层用浏览器回归钉死的，门面再包一层等于给同一件事写两份文档。
 */
import { twipsToPx } from '@uw/core';
import type { DocumentLayout } from '@uw/layout';
import type { DocPosition, DocRange } from '@uw/model';
import type { ClientPoint, ClientRect } from '@uw/view';
import type {
  DecorationHandle,
  DecorationOptions,
  DomView,
  ViewOptions as DomViewOptions,
  OverlayHandle,
  OverlayOptions,
  ScrollOptions,
  ScrollTarget,
} from '@uw/view/dom';
import { mountView } from '@uw/view/dom';

/** 数字是倍率（1 = 100%）；两个 fit 按**最宽 / 最高的那一页**算，混合纸张的文档不会有哪一页出界 */
export type ZoomSpec = number | 'fit-width' | 'fit-page';

export interface ViewOptions {
  /** 默认 1 */
  zoom?: ZoomSpec;
  /** 页间距，CSS px，默认 24 */
  pageGap?: number;
  /** 原生可选文本层（Ctrl+F / 划词复制 / 屏幕阅读器），默认开 */
  textLayer?: boolean;
  /** 只绘制可见页前后几页，默认开；没有 `IntersectionObserver` 时退回全量绘制 */
  virtualize?: boolean;
  /** 视口外多绘制几页，默认 2 */
  overscan?: number;
  /** class 前缀，默认 `uw`。同一页面上挂两份互不干扰的样式时改它 */
  classPrefix?: string;
  /** Word 字体名 → CSS font-family。不传走渲染层的默认映射 */
  fontFamily?: (family: string) => string;
  /** 画出版心框与行盒，调试用 */
  debug?: boolean;
}

export interface Disposable {
  dispose(): void;
}

export interface UwView extends Disposable {
  /** 视图的根元素（已经挂进容器）。挂样式、算尺寸用它，别往里面塞子节点 */
  readonly root: HTMLElement;
  /** 当前生效的倍率。fit 模式下是算出来的那个数 */
  readonly zoom: number;
  /** 屏幕坐标 → 模型位置；纸外、页间空隙、被遮挡时为 null */
  locate(point: ClientPoint): DocPosition | null;
  /** 一个 range 跨行 / 跨页会有多个矩形，CSS px，不裁剪到视口 */
  rectsOf(range: DocRange): ClientRect[];
  caretRect(position: DocPosition): ClientRect | null;
  decorate(range: DocRange, options?: DecorationOptions): DecorationHandle;
  overlay(position: DocPosition, element: HTMLElement, options?: OverlayOptions): OverlayHandle;
  /** 目标排不出来时不动并返回 false（不退到页首 —— 滚到了别处比没滚更难察觉） */
  scrollTo(target: ScrollTarget, options?: ScrollOptions): boolean;
  /** 只改页面尺寸，不重排、不重建文字层与索引 */
  setZoom(zoom: ZoomSpec): void;
}

/** 文档里最宽与最高的一页在 zoom = 1 时的 px 尺寸 —— fit 的分母 */
function extent(layout: DocumentLayout): { width: number; height: number } {
  let width = 0;
  let height = 0;
  for (const p of layout.pages) {
    width = Math.max(width, twipsToPx(p.geometry.width));
    height = Math.max(height, twipsToPx(p.geometry.height));
  }
  return { width, height };
}

/**
 * 容器的内容盒（去掉 padding 与滚动条）。`clientWidth` 已经去掉了滚动条，
 * padding 要自己减 —— 页面是贴着内容盒居中的，按 border-box 算会让 fit-width 出界。
 */
function contentBox(container: Element): { width: number; height: number } {
  const win = container.ownerDocument.defaultView;
  const cs = win?.getComputedStyle(container);
  const px = (v: string | undefined) => Number.parseFloat(v ?? '') || 0;
  return {
    width: container.clientWidth - px(cs?.paddingLeft) - px(cs?.paddingRight),
    height: container.clientHeight - px(cs?.paddingTop) - px(cs?.paddingBottom),
  };
}

function resolveZoom(spec: ZoomSpec, container: Element, layout: DocumentLayout, gap: number): number {
  if (typeof spec === 'number') return spec;
  const page = extent(layout);
  const box = contentBox(container);
  // 页壳两侧各留半个页间距，免得贴着容器边；容器还没排（尺寸为 0）时退回 1，等 ResizeObserver 补
  const fitW = (box.width - gap) / page.width;
  const fitH = (box.height - gap) / page.height;
  const z = spec === 'fit-width' ? fitW : Math.min(fitW, fitH);
  return Number.isFinite(z) && z > 0 ? z : 1;
}

export function createView(
  container: Element,
  layout: DocumentLayout,
  options: ViewOptions,
  imageHref: ((id: string) => string | undefined) | undefined,
): UwView {
  const gap = options.pageGap ?? 24;
  let spec: ZoomSpec = options.zoom ?? 1;
  let zoom = resolveZoom(spec, container, layout, gap);
  let disposed = false;

  // exactOptionalPropertyTypes：可选字段要么不写，要么给确定值
  const domOptions: DomViewOptions = { zoom, pageGap: gap };
  if (options.textLayer !== undefined) domOptions.textLayer = options.textLayer;
  if (options.virtualize !== undefined) domOptions.virtualize = options.virtualize;
  if (options.overscan !== undefined) domOptions.overscan = options.overscan;
  if (options.classPrefix !== undefined) domOptions.classPrefix = options.classPrefix;
  if (options.fontFamily !== undefined) domOptions.fontFamily = options.fontFamily;
  if (options.debug !== undefined) domOptions.debug = options.debug;
  if (imageHref !== undefined) domOptions.imageHref = imageHref;

  const inner: DomView = mountView(container, layout, domOptions);

  // fit 模式跟着容器尺寸走。观察的是**容器**不是 root：root 的尺寸是页面撑出来的，
  // 缩放一改它就变，观察它会自己触发自己
  const win = container.ownerDocument.defaultView;
  let observer: ResizeObserver | undefined;
  function observe(): void {
    if (observer !== undefined || win?.ResizeObserver === undefined || typeof spec === 'number') return;
    observer = new win.ResizeObserver(() => {
      if (disposed || typeof spec === 'number') return;
      const next = resolveZoom(spec, container, layout, gap);
      if (Math.abs(next - zoom) < 1e-6) return;
      zoom = next;
      inner.setZoom(zoom);
    });
    observer.observe(container);
  }
  observe();

  return {
    get root() {
      return inner.root;
    },
    get zoom() {
      return zoom;
    },
    locate: (point) => inner.locate(point),
    rectsOf: (range) => inner.rectsOf(range),
    caretRect: (position) => inner.caretRect(position),
    decorate: (range, opts) => inner.decorate(range, opts),
    overlay: (position, element, opts) => inner.overlay(position, element, opts),
    scrollTo: (target, opts) => inner.scrollTo(target, opts),
    setZoom(next) {
      if (disposed) throw new Error('视图已销毁');
      spec = next;
      zoom = resolveZoom(spec, container, layout, gap);
      inner.setZoom(zoom);
      if (typeof spec === 'number') {
        observer?.disconnect();
        observer = undefined;
      } else observe();
    },
    dispose() {
      if (disposed) return;
      disposed = true;
      observer?.disconnect();
      observer = undefined;
      inner.destroy();
    },
  };
}
