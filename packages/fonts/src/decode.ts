/**
 * 从**字节**解出字体 —— 浏览器侧注册字体的唯一入口。
 *
 * 单独一个模块、单独一个 `@uw/fonts/decode` 子路径，是为了让主入口不依赖 fontkit：
 * 只带度量包的部署（跨平台分发的主力形态，见 metrics-pack.ts）根本不需要解析字体文件，
 * 不该被迫把整个 fontkit 打进包里。`registry.ts` 因此也保持零 fontkit 依赖 ——
 * 注册表是索引，不是解码器。
 *
 * 与 `./node` 的分工：这里 isomorphic（`fontkit.create`，Node 与浏览器都能跑），
 * `load-node.ts` 走 `fontkit.openSync` 读文件系统，只给离线工具用。
 */
import { create } from 'fontkit';
import type { FontkitFont } from './metrics.ts';
import { readRawMetrics } from './metrics.ts';
import type { FontSource } from './registry.ts';
import { fontkitSource } from './registry.ts';

/** 字体集（`.ttc`）解出来是这个形状，`simsun.ttc` 里同时装着 SimSun 与 NSimSun */
interface FontCollectionLike {
  fonts?: FontkitFont[];
  getFont?: (postscriptName: string) => FontkitFont | null;
}

/**
 * 字体集 → 其中一款。`postscriptName` 缺省时取第一款。
 *
 * 单独抽出来是因为 `create`（字节）与 `openSync`（文件）两条路解出的东西形状一样，
 * 拆包逻辑没有理由写两遍。
 */
export function unwrapFont(opened: unknown, postscriptName: string | undefined, where: string): FontkitFont {
  const collection = opened as FontCollectionLike;
  if (!Array.isArray(collection.fonts)) return opened as FontkitFont;
  if (postscriptName !== undefined && typeof collection.getFont === 'function') {
    const picked = collection.getFont(postscriptName);
    if (picked === null) throw new Error(`${where} 里没有 ${postscriptName}`);
    return picked;
  }
  const first = collection.fonts[0];
  if (first === undefined) throw new Error(`${where} 是空的字体集`);
  return first;
}

/**
 * 字节 → fontkit 字体对象。
 *
 * `ArrayBuffer` 与 `Uint8Array` 都收：`fetch(...).then(r => r.arrayBuffer())` 拿到的是前者，
 * 而 fontkit 只吃后者（喂 ArrayBuffer 会在 DataView 构造处抛 TypeError）。
 * 把这个转换放在库里，调用方不必知道这层区别。
 */
export function decodeFont(data: ArrayBuffer | Uint8Array, postscriptName?: string): FontkitFont {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  // @types/fontkit 把参数标成 Buffer，实际实现只要求 Uint8Array —— 浏览器里没有 Buffer
  const opened = (create as unknown as (b: Uint8Array, ps?: string) => unknown)(bytes, postscriptName);
  return unwrapFont(opened, postscriptName, '这段字体字节');
}

/** 字节 → 降级链第 ①级的 `FontSource`，可直接 `registry.register(family, …)` */
export function fontSourceFromBytes(data: ArrayBuffer | Uint8Array, postscriptName?: string): FontSource {
  const font = decodeFont(data, postscriptName);
  return fontkitSource(font, readRawMetrics(font));
}
