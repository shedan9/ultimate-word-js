/**
 * 字体注册表 —— 三级降级的落地（架构 §5.2）。
 *
 * ```
 * ① 真实字体文件（fontkit 解析 OS/2 · hmtx · cmap）         ← 与 Word 一致
 * ② 度量包 JSON（离线从 Windows 字体抽的纯度量，1–2 KB/字体） ← 与 Word 一致
 * ③ 兜底近似（未知字体）                                    ← 页数可能对不上
 * ```
 *
 * 注册表**不碰字节解码**：`FontSource` 是个只有度量与推进宽度的接口，
 * fontkit 在 `fontkitSource()` 那一层就被挡住了。这样注册表本身没有 DOM、没有 fs、
 * 也没有 fontkit 依赖，可以整个搬进 Worker，将来换 Rust/WASM 也只换 source 的实现。
 *
 * 为什么是**全局单例式的注册表**而不是每份文档传一遍：字体解析和度量缓存是纯开销，
 * 同一个页面开十份公文没有理由解析十次宋体（API 设计 §4）。
 */
import type { RawFontMetrics } from './metrics.ts';
import type { MetricsPack } from './metrics-pack.ts';
import { packAdvance, packMetrics } from './metrics-pack.ts';

/** 某个字体名当前落在降级链的哪一级 */
export type FontStatus = 'file' | 'metrics' | 'fallback' | 'missing';

/**
 * 一款可用字体的最小接口。
 *
 * `advance` 返回**字体设计单位**而不是 twips：按字号缩放是一次乘法，
 * 缓存设计单位可以让同一款字体的所有字号共用一份缓存。
 */
export interface FontSource {
  /** 这一份是从真实字体文件来的还是度量包来的 —— 决定 `status()` 报什么 */
  readonly kind: 'file' | 'metrics';
  readonly metrics: RawFontMetrics;
  /** 没有这个字形时返回 undefined，交给上层决定怎么降级 */
  advance(cp: number): number | undefined;
}

/** fontkit 字体对象里我们用到的部分 */
export interface FontkitLike {
  hasGlyphForCodePoint?: (cp: number) => boolean;
  glyphForCodePoint?: (cp: number) => { advanceWidth?: number } | null;
}

/** ① 真实字体文件 */
export function fontkitSource(font: FontkitLike, metrics: RawFontMetrics): FontSource {
  return {
    kind: 'file',
    metrics,
    advance(cp) {
      if (font.hasGlyphForCodePoint?.(cp) === false) return undefined;
      return font.glyphForCodePoint?.(cp)?.advanceWidth;
    },
  };
}

/** ② 度量包 */
export function metricsPackSource(pack: MetricsPack): FontSource {
  const metrics = packMetrics(pack);
  return {
    kind: 'metrics',
    metrics,
    advance: (cp) => packAdvance(pack, cp),
  };
}

export interface ResolvedFont {
  /** 真正命中的名字 —— 可能是替换表换过之后的，诊断里要报这个 */
  family: string;
  source: FontSource;
}

/**
 * 字体名的大小写与空白不该影响命中：文档里写 "SimSun" 与 "simsun " 指的是同一款。
 * 中文字体名没有大小写，归一化对它是恒等变换，不会误伤。
 */
function normalize(name: string): string {
  return name.trim().toLowerCase();
}

export class FontRegistry {
  readonly #sources = new Map<string, ResolvedFont>();
  /** 替换表：找不到 A 时改用 B。存归一化后的名字 */
  readonly #substitutes = new Map<string, string>();

  /** 注册一款字体。同名重复注册以后来者为准 —— 调用方通常是想升级降级等级 */
  register(family: string, source: FontSource): void {
    this.#sources.set(normalize(family), { family, source });
  }

  /** 注册度量包，字体名取包里的 `family`（那是文档里会出现的名字） */
  registerPack(pack: MetricsPack): void {
    this.register(pack.family, metricsPackSource(pack));
  }

  /**
   * 登记替换字体。`{ '仿宋_GB2312': 'Noto Serif CJK SC' }` 的意思是
   * 「找不到仿宋_GB2312 时用它」，**不是**「一律改用它」——
   * 真装了原字体就该用原字体，否则度量会偏离 Word。
   */
  substitute(map: Readonly<Record<string, string>>): void {
    for (const [from, to] of Object.entries(map)) this.#substitutes.set(normalize(from), to);
  }

  /**
   * 一串候选名 → 第一个命中的字体。
   *
   * 候选名从哪来：`@uw/model` 的 `fontNameCandidates()`（文档里的名字 + `w:altName`）。
   * 中文版 Word 写的是「黑体」，磁盘上的字体叫 SimHei，只按一个名字查必然落空 ——
   * 而落空是**静默**的，直到页数对不上才会被发现。
   */
  resolve(candidates: readonly string[]): ResolvedFont | undefined {
    for (const name of candidates) {
      const hit = this.#sources.get(normalize(name));
      if (hit !== undefined) return hit;
    }
    // 原名都没有才轮到替换表：装了真字体就用真字体
    for (const name of candidates) {
      const sub = this.#substitutes.get(normalize(name));
      if (sub === undefined) continue;
      const hit = this.#sources.get(normalize(sub));
      if (hit !== undefined) return hit;
    }
    return undefined;
  }

  /**
   * 降级等级。`fallback` 与 `missing` 的区别：前者有替换表兜着（字形还算像），
   * 后者只能靠等宽近似（字形和度量都不像）。两者都会让排版偏离 Word，
   * 但 `fallback` 值得单独报出来 —— 它是可以靠注册度量包修好的。
   */
  status(family: string | readonly string[]): FontStatus {
    // 收候选名而不是单个名字，是为了和 resolve() 保持一致：
    // 「黑体」查不到但 altName「SimHei」查得到时，报 missing 会是假警报
    const candidates = typeof family === 'string' ? [family] : family;
    for (const name of candidates) {
      const direct = this.#sources.get(normalize(name));
      if (direct !== undefined) return direct.source.kind;
    }
    for (const name of candidates) {
      const sub = this.#substitutes.get(normalize(name));
      if (sub !== undefined && this.#sources.has(normalize(sub))) return 'fallback';
    }
    return 'missing';
  }

  /** 已注册的字体名（注册时的原名，非归一化） */
  families(): string[] {
    return [...this.#sources.values()].map((v) => v.family);
  }

  clear(): void {
    this.#sources.clear();
    this.#substitutes.clear();
  }
}
