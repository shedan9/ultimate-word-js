/**
 * Node 侧字体加载 —— 只给保真度实验、度量包抽取这类离线工具用。
 * 浏览器侧走字节流（内嵌字体 / 用户注册的 webfont / 度量包 JSON），不碰这个文件。
 */
// fontkit 2.x 的 ESM 入口只有具名导出，没有 default
import { openSync } from 'fontkit';
import { unwrapFont } from './decode.ts';
import type { FontkitFont, RawFontMetrics } from './metrics.ts';
import { readRawMetrics } from './metrics.ts';
import type { BuildPackOptions, MetricsPack } from './metrics-pack.ts';
import { buildMetricsPack } from './metrics-pack.ts';
import type { FontSource } from './registry.ts';
import { fontkitSource } from './registry.ts';

/** Windows 系统字体目录 —— 抽度量包与真值实验的字体来源 */
export const WINDOWS_FONT_DIR = 'C:/Windows/Fonts';

/**
 * 打开字体文件；`.ttc` 字体集需要用 postscriptName 指定其中一款
 * （如 simsun.ttc 里同时有 SimSun 与 NSimSun）。
 */
export function openFont(filePath: string, postscriptName?: string): FontkitFont {
  return unwrapFont(openSync(filePath, postscriptName as string), postscriptName, filePath);
}

/**
 * 该字体是否覆盖东亚文字 —— 决定行高走不走 1.3 系数。
 * 直接查 cmap 有没有 U+4E00「一」，比读 OS/2 的 codePageRange 可靠（老字体常填错）。
 */
export function hasEastAsianCoverage(font: FontkitFont): boolean {
  const f = font as { hasGlyphForCodePoint?: (cp: number) => boolean };
  return typeof f.hasGlyphForCodePoint === 'function' && f.hasGlyphForCodePoint(0x4e00);
}

export function readMetricsFromFile(filePath: string, postscriptName?: string): RawFontMetrics {
  return readRawMetrics(openFont(filePath, postscriptName));
}

/** 字体文件 → 降级链第 ①级的 `FontSource`，可直接 `registry.register()` */
export function fileSource(filePath: string, postscriptName?: string): FontSource {
  const font = openFont(filePath, postscriptName);
  return fontkitSource(font, readRawMetrics(font));
}

/**
 * 字体文件 → 度量包。
 *
 * **必须在 Windows 上、对着 `C:/Windows/Fonts` 里的那一份跑** —— 包的意义是把 Word 用的
 * 那份度量搬到别的平台去，抽 macOS 上的同名字体等于把误差固化进随库分发的文件里。
 */
export function buildPackFromFile(
  filePath: string,
  postscriptName?: string,
  opts: BuildPackOptions = {},
): MetricsPack {
  const font = openFont(filePath, postscriptName);
  return buildMetricsPack(font, readRawMetrics(font), opts);
}
