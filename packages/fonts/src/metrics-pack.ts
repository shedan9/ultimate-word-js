/**
 * 度量包 —— 三级降级里的第 ②级，也是 Worker 边界上的字体传输格式。
 *
 * 这两件事撞到同一个答案不是巧合（架构 §9）：Worker 传输需要一份可结构化克隆的度量快照，
 * 跨平台分发需要一份不含字形、不含授权风险的纯度量文件。所以只维护一种格式。
 *
 * 关键认识：**跨平台需要的只是度量，不是字形**。非 Windows 平台用替代字体渲染、
 * 用真实度量排版，断行点与页数就和 Word 完全一致，只是字形外观不同 ——
 * 这比想办法凑齐中文字体授权现实得多。
 *
 * 体积之所以能压到 1–2 KB：CJK 字体里汉字几乎全是 1 em 等宽，只有 ASCII 那一小段是例外。
 * 所以存一个 `defaultAdvance` + 一张只记「与默认不同」的稀疏表就够了。
 */
import type { RawFontMetrics } from './metrics.ts';

/**
 * 一款字体的纯度量快照。**必须可结构化克隆**（原则 1.1）——
 * 它要过 Worker 边界，也要能 `JSON.stringify` 成随库分发的文件。
 */
export interface MetricsPack {
  /** 格式版本。字段语义变了就 +1，老版本包直接拒绝而不是猜 */
  version: 1;
  /** 字体名。用文档里会出现的那个名字（中文文档里就是「黑体」这种本地化名） */
  family: string;
  postscriptName: string;
  unitsPerEm: number;
  os2: RawFontMetrics['os2'];
  hhea: RawFontMetrics['hhea'];
  /** 绝大多数码点的推进宽度，**字体设计单位**。CJK 字体里等于 unitsPerEm */
  defaultAdvance: number;
  /** 例外表：码点（十进制字符串）→ 推进宽度。只存与 `defaultAdvance` 不同的 */
  advances: Record<string, number>;
  /**
   * 有字形的码点区间，`[lo, hi]` 对，升序。
   * 缺省表示「没抽覆盖信息」—— 那就当作全都有，宽度一律走默认值。
   */
  coverage?: [number, number][];
}

/** 抽取度量包时默认采样的码点：ASCII 可见字符 + Latin-1，CJK 字体里例外全在这一段 */
export function defaultSampleCodePoints(): number[] {
  const out: number[] = [];
  for (let cp = 0x20; cp <= 0x7e; cp++) out.push(cp);
  for (let cp = 0xa0; cp <= 0xff; cp++) out.push(cp);
  return out;
}

/** fontkit 里我们真正用到的字形接口。@types/fontkit 没暴露，这里只声明用得着的 */
export interface GlyphSource {
  hasGlyphForCodePoint?: (cp: number) => boolean;
  glyphForCodePoint?: (cp: number) => { advanceWidth?: number } | null;
}

export interface BuildPackOptions {
  /** 覆盖字体名 —— 文档里写的是「黑体」，字体自报的是 SimHei，包要按前者建索引 */
  family?: string;
  /** 采样码点，缺省见 `defaultSampleCodePoints()` */
  sample?: Iterable<number>;
  /**
   * 判定 `defaultAdvance` 用的探针码点，缺省 U+4E00「一」。
   * 拉丁字体没有这个字形，此时退回 unitsPerEm（等宽假设），例外表会把真实宽度补回来。
   */
  probe?: number;
}

/**
 * 从一款已打开的字体抽出度量包。
 *
 * 这一步本身是跨平台的（只依赖 fontkit 对象）；**绑 Windows 的是字体来源** ——
 * 要和 Word 对齐就必须抽 `C:/Windows/Fonts` 里的那一份，macOS 上同名字体度量可能不同。
 */
export function buildMetricsPack(
  font: GlyphSource,
  metrics: RawFontMetrics,
  opts: BuildPackOptions = {},
): MetricsPack {
  const advanceOf = (cp: number): number | undefined => {
    if (font.hasGlyphForCodePoint?.(cp) === false) return undefined;
    return font.glyphForCodePoint?.(cp)?.advanceWidth;
  };

  const probe = opts.probe ?? 0x4e00;
  const defaultAdvance = advanceOf(probe) ?? metrics.unitsPerEm;

  const advances: Record<string, number> = {};
  for (const cp of opts.sample ?? defaultSampleCodePoints()) {
    const w = advanceOf(cp);
    if (w !== undefined && w !== defaultAdvance) advances[String(cp)] = w;
  }

  return {
    version: 1,
    family: opts.family ?? metrics.family,
    postscriptName: metrics.postscriptName,
    unitsPerEm: metrics.unitsPerEm,
    os2: metrics.os2,
    hhea: metrics.hhea,
    defaultAdvance,
    advances,
  };
}

/** 度量包 → 行高计算要的原始度量 */
export function packMetrics(pack: MetricsPack): RawFontMetrics {
  return {
    family: pack.family,
    postscriptName: pack.postscriptName,
    unitsPerEm: pack.unitsPerEm,
    os2: pack.os2,
    hhea: pack.hhea,
  };
}

/** 某个码点的推进宽度（设计单位）；不在覆盖范围内返回 undefined，交给上层降级 */
export function packAdvance(pack: MetricsPack, cp: number): number | undefined {
  const exact = pack.advances[String(cp)];
  if (exact !== undefined) return exact;
  if (pack.coverage !== undefined && !inCoverage(pack.coverage, cp)) return undefined;
  return pack.defaultAdvance;
}

function inCoverage(coverage: readonly (readonly [number, number])[], cp: number): boolean {
  let lo = 0;
  let hi = coverage.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const r = coverage[mid] as readonly [number, number];
    if (cp < r[0]) hi = mid - 1;
    else if (cp > r[1]) lo = mid + 1;
    else return true;
  }
  return false;
}
