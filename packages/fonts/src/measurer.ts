/**
 * `TextMeasurer` —— `@uw/layout` 唯一的字体入口。
 *
 * layout **不许 import 任何字体实现**（架构原则 1.2），只拿一个注入的 `TextMeasurer`。
 * 换 Rust/WASM、换度量来源、在 Worker 里跑，都只换这个接口背后的实现。
 *
 * 热路径（`advances`）刻意做成**纯数据进出**：码点 `Uint32Array` 进、宽度 `Float64Array` 出，
 * 不持有任何 JS 对象引用。理由写在架构 §9：断行与度量是将来抽成 WASM 的两块，
 * 接口从第一天就长成能直接跨语言的样子，比事后重构便宜。
 *
 * 单位是 **twips**（原则 1.3）。返回 px 的度量器一律视为 bug ——
 * 逐字累加的浮点漂移会让最后一个字溢出版心。
 */
import type { DiagnosticSink, Twips } from '@uw/core';
import { fontUnitsToTwips } from '@uw/core';
import type { LineMetrics, LineMetricsOptions, RawFontMetrics } from './metrics.ts';
import { lineMetrics as computeLineMetrics } from './metrics.ts';
import type { FontRegistry, FontStatus } from './registry.ts';
import { isEastAsianCodePoint } from './script.ts';

export interface TextMeasurer {
  status(family: string): FontStatus;
  /** 单倍行距的行度量。`family` 用**候选名**里的第一个即可，内部走同一套解析 */
  lineMetrics(family: string, fontSize: Twips, opts?: LineMetricsOptions): LineMetrics;
  /**
   * 批量推进宽度。`out` 必须至少和 `count`（缺省 `codePoints.length`）一样长，
   * 写入的是 twips。**不做**字距调整、字符缩放（`w:w`）与字间距（`w:spacing`）——
   * 那些是排版属性，属于 layout。
   */
  advances(family: string, fontSize: Twips, codePoints: Uint32Array, out: Float64Array, count?: number): void;
  /** 单个码点，零散调用用（制表位对齐、命中测试）。与 `advances` 共用缓存 */
  advance(family: string, fontSize: Twips, codePoint: number): Twips;
}

/**
 * 级别③ 的兜底度量。
 *
 * 形状照宋体家族取（win 跨度恰好 1 em），因为目标语料是中文公文 ——
 * 这只保证版面不崩，**不保证与 Word 一致**。真要对齐就得注册度量包，
 * 所以每命中一次都会记一条诊断。
 */
export const FALLBACK_METRICS: RawFontMetrics = {
  family: '(fallback)',
  postscriptName: '(fallback)',
  unitsPerEm: 1000,
  os2: {
    winAscent: 860,
    winDescent: 140,
    typoAscender: 860,
    typoDescender: -140,
    typoLineGap: 0,
    useTypoMetrics: false,
  },
  hhea: { ascender: 860, descender: -140, lineGap: 0 },
};

/** 兜底推进宽度：东亚字全角、其余半角。等宽假设，误差随文本长度累积 */
function fallbackAdvance(cp: number, unitsPerEm: number): number {
  return isEastAsianCodePoint(cp) ? unitsPerEm : unitsPerEm / 2;
}

export interface MeasurerOptions {
  /** 字体候选名的展开：文档里的名字 → [名字, altName…]。默认只用原名 */
  candidates?: (family: string) => readonly string[];
  /** 字体缺失只记诊断不抛（架构 §10）—— 用户要的是看到文档 */
  diagnostics?: DiagnosticSink;
  /** 行度量缓存上限，超了按最久未用淘汰 */
  lineMetricsCacheSize?: number;
}

/** 每款字体一份的推进宽度缓存，存**设计单位** —— 所有字号共用同一份 */
interface FamilyCache {
  metrics: RawFontMetrics;
  advance: (cp: number) => number;
  widths: Map<number, number>;
}

export function createTextMeasurer(registry: FontRegistry, opts: MeasurerOptions = {}): TextMeasurer {
  const candidatesOf = opts.candidates ?? ((f: string) => [f]);
  const diagnostics = opts.diagnostics;
  const families = new Map<string, FamilyCache>();
  const lineCache = new Map<string, LineMetrics>();
  const lineCacheMax = opts.lineMetricsCacheSize ?? 512;
  /** 同一款缺失字体只报一次，否则一页公文能刷出上千条一模一样的诊断 */
  const reported = new Set<string>();

  function familyCache(family: string): FamilyCache {
    const cached = families.get(family);
    if (cached !== undefined) return cached;

    const hit = registry.resolve(candidatesOf(family));
    let entry: FamilyCache;
    if (hit === undefined) {
      if (!reported.has(family)) {
        reported.add(family);
        // 对应 API 文档里的 UW_FONT_MISSING
        diagnostics?.warn(
          'font-missing',
          `字体「${family}」既没有字体文件也没有度量包，已用等宽近似，断行与页数可能与 Word 不一致`,
        );
      }
      const upm = FALLBACK_METRICS.unitsPerEm;
      entry = {
        metrics: FALLBACK_METRICS,
        advance: (cp) => fallbackAdvance(cp, upm),
        widths: new Map(),
      };
    } else {
      const src = hit.source;
      const upm = src.metrics.unitsPerEm;
      entry = {
        metrics: src.metrics,
        // 字体里没有这个字形时也得给个宽度：Word 会用替补字体画出来并占位，
        // 返回 0 会让后面所有字的 x 一起左移，比宽度略有偏差严重得多
        advance: (cp) => src.advance(cp) ?? fallbackAdvance(cp, upm),
        widths: new Map(),
      };
    }
    families.set(family, entry);
    return entry;
  }

  function designWidth(cache: FamilyCache, cp: number): number {
    const cached = cache.widths.get(cp);
    if (cached !== undefined) return cached;
    const w = cache.advance(cp);
    cache.widths.set(cp, w);
    return w;
  }

  return {
    status: (family) => registry.status(candidatesOf(family)),

    lineMetrics(family, fontSize, o = {}) {
      const key = `${family}|${fontSize}|${o.source ?? 'win'}|${o.eastAsian === true ? 1 : 0}`;
      const cached = lineCache.get(key);
      if (cached !== undefined) {
        // Map 的插入顺序就是 LRU 顺序：命中后挪到末尾
        lineCache.delete(key);
        lineCache.set(key, cached);
        return cached;
      }
      const value = computeLineMetrics(familyCache(family).metrics, fontSize, o);
      lineCache.set(key, value);
      if (lineCache.size > lineCacheMax) {
        const oldest = lineCache.keys().next();
        if (oldest.done !== true) lineCache.delete(oldest.value);
      }
      return value;
    },

    advances(family, fontSize, codePoints, out, count) {
      const cache = familyCache(family);
      const n = count ?? codePoints.length;
      if (out.length < n) {
        throw new RangeError(`宽度数组太短：需要 ${n}，只有 ${out.length}`);
      }
      const upm = cache.metrics.unitsPerEm;
      for (let i = 0; i < n; i++) {
        out[i] = fontUnitsToTwips(designWidth(cache, codePoints[i] as number), upm, fontSize);
      }
    },

    advance(family, fontSize, codePoint) {
      const cache = familyCache(family);
      return fontUnitsToTwips(designWidth(cache, codePoint), cache.metrics.unitsPerEm, fontSize);
    },
  };
}
