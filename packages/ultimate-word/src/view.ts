/**
 * `UwView` —— 一次屏幕呈现（api.md §5 / §7 / §8）。
 *
 * 它是 `@uw/view/dom` 的 `mountView()` 外面薄薄一层，只加三样东西：
 * ① `zoom` 认 `'fit-width'` / `'fit-page'`（要看容器尺寸，那是门面才知道的事）；
 * ② fit 模式下容器变宽变窄时自动重算（`ResizeObserver`）—— 缩放永不重排（架构 §4.1），
 *    所以跟着容器走是 O(1) 的，没有理由让调用方自己监听；
 * ③ `dispose()` 而不是 `destroy()` —— 凡是「挂上去」的东西都按 `Disposable` 回收（api.md §2 ④）；
 * ④ 事件（api.md §11）：`selection:change` / `viewport:change` / `click:element`，
 *    底层各给一个回调，这里收成统一的 `on()`。
 *
 * 底层 `DomView` 的其余方法原样透出，**不另起一套签名**：装饰、overlay、滚动的语义
 * 都是那一层用浏览器回归钉死的，门面再包一层等于给同一件事写两份文档。
 */
import { twipsToPx } from '@uw/core';
import type { DocumentLayout } from '@uw/layout';
import type { DirectProps, DocPosition, DocRange, QueryNode } from '@uw/model';
import type { ClientPoint, ClientRect } from '@uw/view';
import type {
  DecorationHandle,
  DecorationOptions,
  DomEditing,
  DomView,
  ViewOptions as DomViewOptions,
  EditingBinding,
  OverlayHandle,
  OverlayOptions,
  ScrollOptions,
  ScrollTarget,
} from '@uw/view/dom';
import { mountEditing, mountView } from '@uw/view/dom';
import { Emitter } from './events.ts';

/** 数字是倍率（1 = 100%）；两个 fit 按**最宽 / 最高的那一页**算，混合纸张的文档不会有哪一页出界 */
export type ZoomSpec = number | 'fit-width' | 'fit-page';

export interface ViewOptions {
  /** 默认 preview；edit 接入模型事务与 IME。 */
  mode?: 'preview' | 'edit';
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
  /**
   * Ctrl+P 时怎么印，默认 `document`：一页一张纸，纸张取文档自带的页面设置。
   * `inline` 是原地印（页面怎么排就怎么印），给「文档只是页面一角」的宿主。`print()` 不受它影响
   */
  printMode?: 'document' | 'inline';
}

export interface Disposable {
  dispose(): void;
}

/**
 * 点到的是什么。现在只有超链接 —— 图片的绘制层不接指针事件、布局索引里没有对象的矩形，
 * 内容控件（`w:sdt`）解析时已经剥掉了；等它们有了再加 `kind`，所以先按 `kind` 分支。
 */
export interface ElementHit {
  kind: 'hyperlink';
  /** 超链接所在的 run（直接格式那棵树上的，与 `doc.query('run')` 同一份） */
  node: QueryNode<DirectProps>;
  /** 外部地址，或 `#书签名`（文档内跳转）；关系表里查不到时没有这一项 */
  href?: string;
}

export interface ViewEvents {
  /**
   * 编辑态的选区（含折叠的光标）按值变了：点击、拖选、方向键、输入、撤销后的映射都算，
   * 滚动与缩放不算。预览态没有模型选区（原生文字层的划词不对应 `DocRange`），不触发
   */
  'selection:change': DocRange | null;
  /**
   * 看得见的页（页序号升序，不含预绘制的）或倍率变了。可见页来自 IntersectionObserver，
   * 所以是异步的；没有它的环境（或 `virtualize: false`）只在倍率变化时报，`visiblePages` 为空
   */
  'viewport:change': { visiblePages: readonly number[]; zoom: number };
  /**
   * 点到超链接。视图**不替宿主跳转**：预览态打开新窗口、编辑态 Ctrl+点击才跳，是宿主的决定。
   * 拖选结束时的那一下 click 不算（选区没有折叠）
   */
  'click:element': ElementHit & { position: DocPosition; originalEvent: MouseEvent };
}

export interface UwView extends Disposable {
  /** 视图的根元素（已经挂进容器）。挂样式、算尺寸用它，别往里面塞子节点 */
  readonly root: HTMLElement;
  /** 当前生效的倍率。fit 模式下是算出来的那个数 */
  readonly zoom: number;
  readonly selection?: DocRange | undefined;
  select?(range: DocRange): void;
  focus?(): void;
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
  /** 按文档自带的页面设置打印：一页一张纸、不重排、与屏幕缩放无关。走 `window.print()` */
  print(): void;
  /** 挂一个事件监听者，见 `ViewEvents`。`dispose()` 视图时一并摘掉 */
  on<K extends keyof ViewEvents>(type: K, listener: (payload: ViewEvents[K]) => void): Disposable;
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
  editing?: EditingBinding,
  updates?: { subscribe(callback: (layout: DocumentLayout) => void): () => void },
  elementAt?: (position: DocPosition) => ElementHit | undefined,
): UwView {
  const gap = options.pageGap ?? 24;
  let spec: ZoomSpec = options.zoom ?? 1;
  let zoom = resolveZoom(spec, container, layout, gap);
  let disposed = false;

  // exactOptionalPropertyTypes：可选字段要么不写，要么给确定值
  const domOptions: DomViewOptions = { zoom, pageGap: gap };
  domOptions.textLayer = options.textLayer ?? options.mode !== 'edit';
  if (options.virtualize !== undefined) domOptions.virtualize = options.virtualize;
  if (options.overscan !== undefined) domOptions.overscan = options.overscan;
  if (options.classPrefix !== undefined) domOptions.classPrefix = options.classPrefix;
  if (options.fontFamily !== undefined) domOptions.fontFamily = options.fontFamily;
  if (options.debug !== undefined) domOptions.debug = options.debug;
  if (options.printMode !== undefined) domOptions.printMode = options.printMode;
  if (imageHref !== undefined) domOptions.imageHref = imageHref;

  const inner: DomView = mountView(container, layout, domOptions);
  const events = new Emitter<ViewEvents>();
  let visiblePages: readonly number[] = [];
  function viewportChanged(): void {
    events.emit('viewport:change', { visiblePages, zoom });
  }
  const stopViewport = inner.onViewport((pages) => {
    visiblePages = pages;
    viewportChanged();
  });

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
      viewportChanged();
    });
    observer.observe(container);
  }
  observe();
  let input: DomEditing | undefined;
  if (options.mode === 'edit' && editing !== undefined) input = mountEditing(container, inner, editing);
  const stopSelection = input?.onSelectionChange((range) => events.emit('selection:change', range ?? null));
  const unsubscribe = updates?.subscribe((next) => {
    layout = next;
    const before = zoom;
    zoom = resolveZoom(spec, container, layout, gap);
    inner.update(layout, { zoom });
    input?.refresh();
    // 重排可能改了最宽的那一页，fit 模式的倍率跟着变
    if (Math.abs(zoom - before) > 1e-6) viewportChanged();
  });

  // 挂在容器上（不是 root）：重排会换掉 root，容器是不动的
  function click(event: MouseEvent): void {
    if (elementAt === undefined || !events.has('click:element') || event.button !== 0) return;
    const win = container.ownerDocument.defaultView;
    if (win === null || !(event.target instanceof win.Node) || !inner.root.contains(event.target)) return;
    // 拖选结束也会来一下 click：原生选区（预览态）或模型选区（编辑态）没折叠就不算点击
    const native = container.ownerDocument.getSelection();
    if (native !== null && !native.isCollapsed && inner.root.contains(native.anchorNode)) return;
    const selected = input?.selection;
    if (selected !== undefined && !sameSpot(selected.start, selected.end)) return;
    const position = inner.locate(event);
    const hit = position === null ? undefined : elementAt(position);
    if (position === null || hit === undefined) return;
    events.emit('click:element', { ...hit, position, originalEvent: event });
  }
  container.addEventListener('click', click as EventListener);

  return {
    get root() {
      return inner.root;
    },
    get zoom() {
      return zoom;
    },
    get selection() {
      return input?.selection;
    },
    select(range) {
      if (!input) throw new Error('视图不是编辑态');
      input.select(range);
    },
    focus() {
      if (input) input.focus();
      else inner.root.focus();
    },
    locate: (point) => inner.locate(point),
    rectsOf: (range) => inner.rectsOf(range),
    caretRect: (position) => inner.caretRect(position),
    decorate: (range, opts) => inner.decorate(range, opts),
    overlay: (position, element, opts) => inner.overlay(position, element, opts),
    scrollTo: (target, opts) => inner.scrollTo(target, opts),
    print: () => inner.print(),
    on: (type, listener) => {
      if (disposed) throw new Error('视图已销毁');
      return events.on(type, listener);
    },
    setZoom(next) {
      if (disposed) throw new Error('视图已销毁');
      spec = next;
      const before = zoom;
      zoom = resolveZoom(spec, container, layout, gap);
      inner.setZoom(zoom);
      input?.refresh();
      if (Math.abs(zoom - before) > 1e-6) viewportChanged();
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
      unsubscribe?.();
      stopViewport();
      stopSelection?.();
      container.removeEventListener('click', click as EventListener);
      events.clear();
      input?.dispose();
      inner.destroy();
    },
  };
}

function sameSpot(a: DocPosition, b: DocPosition): boolean {
  return a.nodeId === b.nodeId && a.contentIndex === b.contentIndex && a.offset === b.offset;
}
