/**
 * 全局字体注册表 —— `UltimateWord.fonts`（api.md §4）。
 *
 * 为什么是全局单例而不是每份文档传一遍：字体解析与度量缓存是纯开销，同一个页面开十份公文
 * 没有理由解析十次宋体。`LoadOptions.fonts` 仍能按文档覆盖，默认继承这一份。
 *
 * 随库那 17 款度量包在**模块加载时**就注册进去，调用方一行都不用写 —— 这正是
 * 「`load()` 替调用方做掉」的落点。走 `@uw/fonts/packs`（JSON import），不走
 * `@uw/fonts/node`（`readFileSync`，浏览器里没有）。
 *
 * `register(family, bytes)` 是**异步**的（api.md 原来写的是同步）：解码字体要 fontkit，
 * 而主入口刻意不静态依赖它 —— 只带度量包的部署不该被迫把 fontkit 打进去
 * （`@uw/fonts` 把它关在 `/decode` 子路径里就是为了这个）。动态 `import()` 让打包器
 * 自然拆出一个 chunk，第一次调 `register` 才加载。
 */
import type { FontStatus, MetricsPack } from '@uw/fonts';
import { FontRegistry } from '@uw/fonts';
import { bundledPacks } from '@uw/fonts/packs';

export interface FontsApi {
  /**
   * 注册一款真实字体（降级链第 ① 级）。`.ttc` 字体集要用 `postscriptName` 指定其中一款。
   * 同名重复注册以后来者为准 —— 通常是从度量包升级到真字体。
   */
  register(family: string, data: ArrayBuffer | Uint8Array, postscriptName?: string): Promise<void>;
  /** 注册度量包（第 ② 级）。随库 17 款已经在了，这是给清单之外的字体用的 */
  registerMetrics(pack: MetricsPack): void;
  /** 替换表：「找不到 A 时用 B」，不是「一律改用 B」—— 装了原字体就用原字体 */
  substitute(map: Readonly<Record<string, string>>): void;
  /** 降级等级。`fallback` 是替换表兜住了，`missing` 才是等宽近似 */
  status(family: string | readonly string[]): FontStatus;
  /** 底层注册表。要按文档覆盖时，拿它 `new` 一份再传给 `LoadOptions.fonts` */
  readonly registry: FontRegistry;
}

/** 造一份已经装好随库度量包的注册表 —— 全局那份与「按文档覆盖」那份都从这儿来 */
export function createRegistry(): FontRegistry {
  const registry = new FontRegistry();
  for (const pack of bundledPacks()) registry.registerMetrics(pack);
  return registry;
}

export function createFontsApi(registry: FontRegistry): FontsApi {
  return {
    registry,
    async register(family, data, postscriptName) {
      const { fontSourceFromBytes } = await import('@uw/fonts/decode');
      registry.register(family, fontSourceFromBytes(data, postscriptName));
    },
    registerMetrics: (pack) => registry.registerMetrics(pack),
    substitute: (map) => registry.substitute(map),
    status: (family) => registry.status(family),
  };
}
