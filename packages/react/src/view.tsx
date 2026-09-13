/**
 * `<UltimateWordView>` —— `doc.mount()` 的声明式包装（api.md §14）。
 *
 * 它只做三件事：① 把 `doc.mount(容器, 选项)` 的生命周期绑到 React 的（挂 / 卸 / 换 doc 就重挂）；
 * ② `zoom` 单独走 `setZoom()` —— 缩放永不重排，改它不该重挂；③ `overlays` 走声明式：
 * 每个 spec 一个宿主 `<div>`，内容用 portal 渲染进去，宿主交给 `view.overlay()` 定位。
 * 批注气泡因此是**正常的 React 子树**（context、事件、state 都在），视图只负责它摆在哪。
 *
 * 三处容易搞反：
 * - **构造选项变了就重挂**（`pageGap` / `textLayer` / `virtualize` / … / `fontFamily`）：
 *   `UwView` 没有 `update()`（api.md §5），门面自己就是 dispose 再 mount。所以 `fontFamily`
 *   要给稳定引用（`useCallback`），内联箭头函数会让每次渲染都重挂
 * - overlay 的宿主元素在**渲染期**造（按 key 缓存）而不是在 effect 里：portal 要在同一趟渲染里
 *   拿到目标节点，等 effect 造完再 setState 会多画一帧空白
 * - 视图重挂后旧的 `OverlayHandle` 已随旧视图一起没了，`reconcileOverlays` 按 `owner` 认出来
 *   整组重挂；宿主元素与其中的 React 子树**不动**，气泡里打了一半的字还在
 */
import type { CSSProperties, ReactNode, Ref } from 'react';
import { forwardRef, useEffect, useImperativeHandle, useLayoutEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import type { UwDocument, UwView, ViewOptions, ZoomSpec } from 'ultimate-word';
import type { OverlayEntry, OverlayKey, OverlaySpec } from './overlays.ts';
import { disposeOverlays, reconcileOverlays } from './overlays.ts';

export interface ReactOverlaySpec extends OverlaySpec {
  /** 气泡的内容。每次渲染都会重新调用，与普通子元素一样 */
  render: () => ReactNode;
}

export interface UltimateWordViewProps extends ViewOptions {
  doc: UwDocument;
  /** 声明式批注：只描述「哪些、锚在哪」，挂 / 卸 / 跟随由组件转成 `view.overlay()` */
  overlays?: readonly ReactOverlaySpec[];
  /** 容器 `<div>` 的 class 与样式。滚动容器就是它 —— 给它 `overflow: auto` 与高度 */
  className?: string;
  style?: CSSProperties;
  /** 视图挂上 / 重挂 / 销毁（给 null）时叫。想直接拿句柄用 `ref` */
  onViewChange?: (view: UwView | null) => void;
}

export const UltimateWordView = forwardRef<UwView | null, UltimateWordViewProps>(function UltimateWordView(
  props,
  ref: Ref<UwView | null>,
) {
  const {
    doc,
    overlays = [],
    className,
    style,
    onViewChange,
    pageGap,
    textLayer,
    virtualize,
    overscan,
    classPrefix,
    fontFamily,
    debug,
    printMode,
  } = props;
  const zoom: ZoomSpec = props.zoom ?? 1;
  const container = useRef<HTMLDivElement>(null);
  const [view, setView] = useState<UwView | null>(null);
  const hosts = useRef(new Map<OverlayKey, HTMLDivElement>()).current;
  const entries = useRef(new Map<OverlayKey, OverlayEntry<ReactOverlaySpec>>()).current;
  const latest = useRef({ zoom, onViewChange });
  latest.current = { zoom, onViewChange };
  /** 视图当前生效的 zoom —— 与 props 不同步的那一瞬（刚重挂）靠它判断要不要 setZoom */
  const applied = useRef<{ view: UwView | null; zoom: ZoomSpec }>({ view: null, zoom });

  // 挂载：doc 或任一构造选项变了就 dispose 再 mount —— 依赖表就是构造选项那几项。
  // zoom 从 ref 读，它不是重挂的理由；overlays 与回调也刻意不在里面
  useLayoutEffect(() => {
    const el = container.current;
    if (el === null) return;
    // exactOptionalPropertyTypes：可选字段要么不写，要么给确定值
    const options: ViewOptions = { zoom: latest.current.zoom };
    if (pageGap !== undefined) options.pageGap = pageGap;
    if (textLayer !== undefined) options.textLayer = textLayer;
    if (virtualize !== undefined) options.virtualize = virtualize;
    if (overscan !== undefined) options.overscan = overscan;
    if (classPrefix !== undefined) options.classPrefix = classPrefix;
    if (fontFamily !== undefined) options.fontFamily = fontFamily;
    if (debug !== undefined) options.debug = debug;
    if (printMode !== undefined) options.printMode = printMode;
    const next = doc.mount(el, options);
    applied.current = { view: next, zoom: options.zoom as ZoomSpec };
    setView(next);
    return () => {
      setView(null);
      next.dispose();
    };
  }, [doc, pageGap, textLayer, virtualize, overscan, classPrefix, fontFamily, debug, printMode]);

  // zoom：只改页面尺寸。挂载那一趟已经带上了当时的 zoom（记在 applied 里），这里只管之后的变化
  useLayoutEffect(() => {
    if (view === null || applied.current.view !== view || applied.current.zoom === zoom) return;
    view.setZoom(zoom);
    applied.current.zoom = zoom;
  }, [view, zoom]);

  useImperativeHandle<UwView | null, UwView | null>(ref, () => view, [view]);
  // 视图来了报一次、走了（销毁或重挂前）报一次 null；重挂就是「null → 新视图」两声
  useEffect(() => {
    if (view === null) return;
    latest.current.onViewChange?.(view);
    return () => latest.current.onViewChange?.(null);
  }, [view]);

  // overlays：按 key 调和。宿主元素在渲染期按 key 造好，portal 与 view.overlay() 共用同一个节点
  const hostFor = (key: OverlayKey): HTMLDivElement => {
    let host = hosts.get(key);
    if (host === undefined) {
      host = document.createElement('div');
      host.dataset.uwOverlayKey = String(key);
      hosts.set(key, host);
    }
    return host;
  };
  const canPortal = typeof document !== 'undefined';
  useLayoutEffect(() => {
    reconcileOverlays(entries, overlays, {
      owner: view,
      mount: (spec, options) => (view as UwView).overlay(spec.anchor, hostFor(spec.key), options),
      release: (key) => hosts.delete(key),
    });
  });
  useEffect(() => () => disposeOverlays(entries), [entries]);

  return (
    <div ref={container} className={className} style={style}>
      {canPortal
        ? overlays.map((spec) => createPortal(spec.render(), hostFor(spec.key), String(spec.key)))
        : null}
    </div>
  );
});
