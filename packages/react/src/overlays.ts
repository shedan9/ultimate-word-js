/**
 * 声明式 overlay 的调和：props 里的一列 `OverlaySpec` ↔ 视图上已经挂着的一组 `OverlayHandle`。
 *
 * 按 `key` 配对（与 React 列表同理）：新 key 挂、消失的 key 摘、锚点变了 `update()`、
 * placement / offset 变了只能摘了重挂 —— `OverlayHandle.update` 只认锚点，这是底层定的，
 * 这里不另开口子。**这一层不碰 React**，纯函数 + 回调，Node 里就能测；
 * 组件那边只负责把宿主元素与 portal 接上。
 */
import type { DocPosition, OverlayHandle, OverlayOptions, OverlayPlacement } from 'ultimate-word';
import { sameOffset, samePosition } from './same.ts';

export type OverlayKey = string | number;

export interface OverlaySpec {
  /** 同一列里唯一。批注的 id 就是它 —— 换了 key 等于换了一个气泡，输入状态不保留 */
  key: OverlayKey;
  anchor: DocPosition;
  placement?: OverlayPlacement;
  /** 页面壳内 CSS px，不随 zoom 放大 */
  offset?: { x: number; y: number };
}

export interface OverlayEntry<S extends OverlaySpec = OverlaySpec> {
  spec: S;
  handle: OverlayHandle;
  /** 挂在哪个视图上；视图重挂后旧句柄已经随视图一起没了，得重新挂 */
  owner: object;
}

export interface ReconcileHost<S extends OverlaySpec> {
  /** 当前视图的身份；变了就整组重挂 */
  owner: object | null;
  mount(spec: S, options: OverlayOptions): OverlayHandle;
  /** key 从列表里消失了，宿主元素可以扔了 */
  release(key: OverlayKey): void;
}

export function overlayOptions(spec: OverlaySpec): OverlayOptions {
  // exactOptionalPropertyTypes：可选字段要么不写，要么给确定值
  const options: OverlayOptions = {};
  if (spec.placement !== undefined) options.placement = spec.placement;
  if (spec.offset !== undefined) options.offset = { ...spec.offset };
  return options;
}

function sameOptions(a: OverlaySpec, b: OverlaySpec): boolean {
  return a.placement === b.placement && sameOffset(a.offset, b.offset);
}

export function reconcileOverlays<S extends OverlaySpec>(
  current: Map<OverlayKey, OverlayEntry<S>>,
  next: readonly S[],
  host: ReconcileHost<S>,
): void {
  const wanted = new Map<OverlayKey, S>();
  for (const spec of next) {
    if (wanted.has(spec.key)) throw new Error(`overlay 的 key 重复：${String(spec.key)}`);
    wanted.set(spec.key, spec);
  }
  // 先摘后挂：同一个宿主元素不能同时挂在两个 overlay 上
  for (const [key, entry] of current) {
    const spec = wanted.get(key);
    if (spec !== undefined && entry.owner === host.owner && sameOptions(entry.spec, spec)) continue;
    entry.handle.dispose();
    current.delete(key);
    if (spec === undefined) host.release(key);
  }
  if (host.owner === null) return;
  for (const spec of next) {
    const entry = current.get(spec.key);
    if (entry === undefined) {
      current.set(spec.key, { spec, handle: host.mount(spec, overlayOptions(spec)), owner: host.owner });
      continue;
    }
    if (!samePosition(entry.spec.anchor, spec.anchor)) entry.handle.update(spec.anchor);
    entry.spec = spec;
  }
}

export function disposeOverlays(current: Map<OverlayKey, OverlayEntry>): void {
  for (const entry of current.values()) entry.handle.dispose();
  current.clear();
}
