/**
 * 字体度量 —— 整个保真度的地基。
 *
 * Word 的「单倍行距」**不是** CSS 的 `line-height: normal`：它取字体 OS/2 表的
 * usWinAscent / usWinDescent（部分字体退回 hhea），乘以字号得到行高。浏览器各引擎
 * 对 normal 的算法互不相同，所以这一步必须自己读字体表，不能问浏览器。
 *
 * 行高分两问，各由一次穿刺定死：**总量**（Phase 0，13 个样本）与**基线在行高里的位置**
 * （基线穿刺，26 个样本）。第二问的答案在 `baselineOffset()`，它比第一问更容易搞反 ——
 * 行高对了但基线偏了，整页文字会整体上移或下移，而每行的间距看起来完全正常。
 *
 * ## 走哪一套，看的是**字体**不是字符（2026-09-05 实测）
 *
 * `eastAsian` 这个参数由调用方给，判据**不是**「这一行有没有东亚字符」而是
 * 「这一段文字实际用的那款字体是不是东亚字体」—— 一行只有「A2C6」四个半角字符、
 * 用等线画的，Word 照样按东亚规则算（20.32pt / 15pt 字，按拉丁规则只有 15.63pt，差 30%）。
 * 样本 `spike-script-01`（11 页 × 每页四段，`pnpm --filter @uw/fidelity spike:script`），
 * 规则表与证据在 `@uw/layout` 的 `SCRIPT_RULES`。`@uw/fonts` 这一侧的入口是
 * `TextMeasurer.eastAsianFont()`：查这款字体有没有 U+4E00 的字形。
 *
 * 原来按字符判是从 Phase 0 那 13 个样本推出去的 —— 那批样本里纯拉丁的行用的是
 * Times New Roman，「字符是拉丁的」与「字体是拉丁字体」在它们身上完全重合，分不开。
 */
import type { Twips } from '@uw/core';
import { fontUnitsToTwips } from '@uw/core';

/** 字体自带的、未经字号缩放的原始度量（单位：字体设计单位） */
export interface RawFontMetrics {
  family: string;
  postscriptName: string;
  unitsPerEm: number;
  /** OS/2 表。usWinAscent / usWinDescent 均为正数，descent 是「向下多少」 */
  os2: {
    winAscent: number;
    winDescent: number;
    typoAscender: number;
    typoDescender: number;
    typoLineGap: number;
    /** fsSelection bit 7：置位时应优先用 typo 系列度量 */
    useTypoMetrics: boolean;
  };
  /** hhea 表。ascender 正、descender 负 */
  hhea: {
    ascender: number;
    descender: number;
    lineGap: number;
  };
}

/** 某个字号下的行度量（twips） */
export interface LineMetrics {
  /** 字体的 win 上沿到基线。**不含**任何额外行距 */
  ascent: Twips;
  /** 基线到字体的 win 下沿（正数） */
  descent: Twips;
  /** 额外行距：东亚是 win 跨度的 30%，拉丁是 GDI 外部行距 */
  lineGap: Twips;
  /** 单倍行距的行高 = ascent + descent + lineGap */
  lineHeight: Twips;
  /**
   * 「核心盒」上沿到基线 —— 行高被拉大时**不动**的那一段。
   *
   * 东亚 = `ascent`（30% 额外行距上下均分，不属于核心盒）；
   * 拉丁 = `ascent + lineGap`（GDI 外部行距整块坐在基线以上，属于核心盒）。
   * 这个区别是实测出来的，见 `baselineOffset()`。
   */
  coreAbove: Twips;
}

/** 行高的来源策略。默认 win —— 与 Word 在 Windows 上的行为一致 */
export type MetricSource = 'win' | 'typo' | 'hhea';

/**
 * 东亚文字的行高系数。
 *
 * 实测（apps/fidelity 的 spike-lineheight-01 / 02）：含东亚文字的行，Word 的单倍行距
 * 是字体 win 度量跨度的 **1.3 倍**，且**不加** GDI 外部行距；纯拉丁文的行则是
 * `winAscent + winDescent + 外部行距`，没有这个系数。
 *
 * | 字体 | winSpan | 预测 | 实测 |
 * |---|---|---|---|
 * | 仿宋 / 宋体 / 黑体 / 楷体（em=256） | 1.0000 | 1.300 | 1.298 ~ 1.300 |
 * | 微软雅黑 | 1.3198 | 1.716 | 1.710 ~ 1.720 |
 * | 等线 | 1.0420 | 1.355 | 1.350 ~ 1.360 |
 * | Times New Roman | 1.1074(+0.0425 leading) | 1.150 | 1.150 |
 * | Arial | 1.1172(+0.0327 leading) | 1.150 | 1.150 |
 *
 * 误差全部 < 0.15pt（其中约 0.1pt 是 Word 导出 PDF 时的坐标取整）。
 */
export const EAST_ASIAN_LINE_FACTOR = 1.3;

/**
 * GDI 的外部行距（TEXTMETRIC.tmExternalLeading）。
 * hhea 的 lineGap 里，超出 win 跨度与 hhea 跨度之差的那部分才真正加进行高。
 */
export function gdiExternalLeading(m: RawFontMetrics): number {
  const winSpan = m.os2.winAscent + m.os2.winDescent;
  const hheaSpan = m.hhea.ascender - m.hhea.descender;
  return Math.max(0, m.hhea.lineGap - (winSpan - hheaSpan));
}

export interface LineMetricsOptions {
  source?: MetricSource;
  /**
   * 这段文字用的那款字体**是不是东亚字体** —— 决定用 1.3 系数还是 GDI 外部行距。
   * 名字里没有「line」是有意的：判据是字体不是字符，见文件头。
   */
  eastAsian?: boolean;
}

export function readRawMetrics(font: FontkitFont): RawFontMetrics {
  const os2 = font['OS/2'] as Os2Table | undefined;
  const hhea = font.hhea as HheaTable | undefined;
  if (!os2 || !hhea) {
    throw new Error(`字体缺少 OS/2 或 hhea 表：${font.postscriptName ?? '(unknown)'}`);
  }
  return {
    family: font.familyName ?? font.postscriptName ?? '',
    postscriptName: font.postscriptName ?? '',
    unitsPerEm: font.unitsPerEm,
    os2: {
      winAscent: os2.winAscent,
      winDescent: os2.winDescent,
      typoAscender: os2.typoAscender,
      typoDescender: os2.typoDescender,
      typoLineGap: os2.typoLineGap,
      // fsSelection bit 7 = USE_TYPO_METRICS
      useTypoMetrics: ((os2.fsSelection ?? 0) & 0x80) !== 0,
    },
    hhea: {
      ascender: hhea.ascent,
      descender: hhea.descent,
      lineGap: hhea.lineGap,
    },
  };
}

/**
 * 单倍行距的行度量。
 *
 * Word 在 Windows 上走 GDI 度量，而不是 CSS 的 `line-height: normal`：
 * - 拉丁：`lineHeight = (usWinAscent + usWinDescent + 外部行距) × 字号 / unitsPerEm`
 * - 东亚：`lineHeight = (usWinAscent + usWinDescent) × 1.3 × 字号 / unitsPerEm`，不加外部行距
 *
 * typo / hhea 两个来源留着做对照实验与兜底 —— 有些字体 win 度量异常大。
 *
 * `eastAsian` 由调用方按**这段文字用的那款字体**判定后传进来（见文件头），这里不自己猜 ——
 * 度量层拿到的只有一份 `RawFontMetrics`，它认不出「仿宋」和「Times」的区别，
 * 而覆盖率要问 `FontSource`。原来这里写的是「`eastAsian` 是**整行**的属性」，被
 * `spike-script-01` 推翻了：同一行里的两段文字完全可以一段走东亚规则、一段走拉丁规则。
 */
export function lineMetrics(m: RawFontMetrics, fontSize: Twips, opts: LineMetricsOptions = {}): LineMetrics {
  const source = opts.source ?? 'win';
  const u = (v: number): Twips => fontUnitsToTwips(v, m.unitsPerEm, fontSize);
  let ascent: number;
  let descent: number;
  let lineGap: number;
  switch (source) {
    case 'win':
      ascent = u(m.os2.winAscent);
      descent = u(m.os2.winDescent);
      lineGap = opts.eastAsian ? (ascent + descent) * (EAST_ASIAN_LINE_FACTOR - 1) : u(gdiExternalLeading(m));
      break;
    case 'typo':
      ascent = u(m.os2.typoAscender);
      descent = u(-m.os2.typoDescender);
      lineGap = u(m.os2.typoLineGap);
      break;
    case 'hhea':
      ascent = u(m.hhea.ascender);
      descent = u(-m.hhea.descender);
      lineGap = u(m.hhea.lineGap);
      break;
  }
  // 核心盒：行高被行距规则 / 网格拉大时**不跟着变**的那一段。拉丁的外部行距整块坐在
  // 基线以上，所以它属于核心盒；东亚那 30% 要上下均分，所以不属于。
  // typo / hhea 两条对照路径的 lineGap 是字体自报的行间距，语义与 GDI 外部行距不同，
  // 没有真值支持，一律按「不属于核心盒」处理 —— 它们本来也只用于兜底和对照实验。
  const coreAbove = source === 'win' && opts.eastAsian !== true ? ascent + lineGap : ascent;
  return { ascent, descent, lineGap, lineHeight: ascent + descent + lineGap, coreAbove };
}

/**
 * 行顶到基线的距离 —— **核心盒在最终行高里居中**。
 *
 * 实测（apps/fidelity 的 spike-baseline-01/02/03，26 个首行样本，最大误差 0.13pt）：
 * 基线位置只由核心盒与最终行高决定，多出来的空间**一律上下均分**，
 * 不管它是从哪儿来的 ——
 *
 * | 额外空间的来源 | 分配 | 样本 |
 * |---|---|---|
 * | 东亚的 30% 额外行距 | 上下均分 | 宋体 / 仿宋 / 黑体 / 楷体 / 雅黑 / 等线，12–72pt |
 * | 拉丁的 GDI 外部行距 | **全在基线以上** | Times New Roman / Arial，12 / 48 / 72pt |
 * | 行网格吸附的余量 | 上下均分（东亚与拉丁同规则） | spike-baseline-03 前四段 |
 * | 倍数行距放大的余量 | 上下均分 | spike-baseline-03 末三段，1.5 / 2.0 倍 |
 *
 * 「拉丁的外部行距全在上」不是特例而是同一条规则的推论：GDI 把 tmExternalLeading
 * 定义成**行与行之间**的空隙，Word 把它加在每行 ascent 之上（页首第一行也加），
 * 于是它进了核心盒；进了核心盒的东西自然不参与均分。
 *
 * 分辨力来自大字号：12pt 下「一半在上」与「全在上」对 Times 只差 0.26pt，
 * 与导出 PDF 的坐标噪声（Phase 0 实测 0.13pt）同量级；72pt 下差 1.53pt 才分得开。
 * 所以 spike-baseline-02 存在的理由不是「多测几款字体」，而是**把信号放大到噪声之上**。
 * 残差全为负（实测比预测小 0.03–0.13pt），量级与 Phase 0 一致，怀疑是 Word 在
 * 某处对度量取了整；没有证据之前不去凑这个系数，凑了就变成假精度。
 */
export function baselineOffset(m: LineMetrics, lineHeight: Twips): Twips {
  const core = m.coreAbove + m.descent;
  return m.coreAbove + (lineHeight - core) / 2;
}

/**
 * 固定值行距（`w:lineRule="exact"`）的基线：**行高的 80%**，与字体、字号都无关。
 *
 * 这一条推翻了「核心盒在行高里居中」在固定值行距下的适用性 —— 上面那条规则是拿
 * 单倍 / 倍数 / 网格三种行距标定的（spike-baseline-01/02/03 的 `lineSpacingPt` 全是 0），
 * 固定值那一格是空的。`spike-page-01`（固定行距 20pt）一跑就露馅：整页文字比预测低
 * 1.77pt，正好是仿宋 12pt 那 30% 额外行距（3.6pt）的一半。
 *
 * `spike-baseline-04` 把信号放大到分得开，结论比「额外行距全在基线以上」更干脆 ——
 * **基线位置只是行高的一个固定比例**：
 *
 * | 字体 @ 字号 | 固定行距 | 自然行高 | 实测基线 | 实测 ÷ 行高 |
 * |---|---|---|---|---|
 * | 仿宋 @ 48 | 80.04 | 62.40 | 64.05 | 0.8002 |
 * | 黑体 @ 48 | 80.04 | 62.40 | 64.05 | 0.8002 |
 * | 仿宋 @ 48 | 50.04 | 62.40 | 40.05 | 0.8004 |
 * | 黑体 @ 48 | 50.04 | 62.40 | 40.05 | 0.8004 |
 * | 仿宋 @ 12 | 20.04 | 15.60 | 16.05 | 0.8009 |
 * | Times @ 48 | 80.04 | 55.20 | 64.05 | 0.8002 |
 *
 * 三个证据一起看：① 两款东亚字体在同一个行高上给出**同一个**基线；
 * ② 拉丁字体也给出**同一个**基线（自然行高差 7.2pt 却毫无影响）；
 * ③ 行高比自然行高小（50.04 < 62.40，字被压）时同样是 0.8。
 * 残差恒为 +0.018pt，那是 Word 自述页边距（70.85）与真值 25mm（70.866）的取整差，
 * 不是系数没标准 —— 所以取整 0.8，不去凑第三位。
 *
 * 同一批样本里的**单倍行距对照**仍然落在「核心盒居中」上（仿宋 48pt 差 -0.06pt、
 * Times 48pt 差 -0.085pt），所以这不是把旧结论推翻，是补上它没覆盖的那一格。
 *
 * ⚠️ 未标定：`atLeast`（最小值行距）在「固定值赢了」的那一侧算哪一套没有样本。
 */
export const EXACT_LINE_BASELINE_RATIO = 0.8;

export function baselineOffsetExact(lineHeight: Twips): Twips {
  return lineHeight * EXACT_LINE_BASELINE_RATIO;
}

/**
 * 混排行的自然行高：取各字体单倍行距行高的最大值。
 *
 * ⚠️ 单字体的行**恒等于**它自己的行高，Phase 0 验的是这一格；混排行上它**是错的**
 * —— 见 `composeLineBox()` 的证据表。只剩 `maxHeight` 那条落选分支还在用它。
 */
export function naturalLineHeight(parts: readonly LineMetrics[]): Twips {
  let h = 0;
  for (const p of parts) {
    if (p.lineHeight > h) h = p.lineHeight;
  }
  return h;
}

/**
 * 多款字体合成一个基线：**每款各自在最终行高里居中，再取最大值**。
 *
 * ⚠️ 这是**输给了实测的那一条**，留着只为 `spike:script` 还能把落选的组合再跑一遍
 * （与 `HEADER_RULES` 里落选的 `footerAnchor: 'top'` 同理）。真正在用的是
 * `composeLineBox()` 的 `maxSides`，见那里的证据表。
 */
export function composeBaseline(parts: readonly LineMetrics[], lineHeight: Twips): Twips {
  let baseline = 0;
  for (const p of parts) {
    const b = baselineOffset(p, lineHeight);
    if (b > baseline) baseline = b;
  }
  return baseline;
}

/** 多款字体合成一行行盒的两种写法。实测的是 `maxSides`，见 `composeLineBox()` */
export type ComposeRule = 'maxHeight' | 'maxSides';

/** 一行的行盒：自然行高 + 行顶到基线。行距规则与网格吸附在这之后才动它 */
export interface LineBox {
  natural: Twips;
  above: Twips;
}

/**
 * 同一行里几款字体怎么合成一个行盒 —— **各自的行盒逐项取 max**（`maxSides`）。
 *
 * 实测（`spike-script-01` 的 P9/P10/P11，三页 9 个基线差，误差 0.04pt）。这一格
 * 在此之前是**判断**而不是结论，因为基线穿刺那 26 个样本里每一行的行盒都由**单独一款**
 * 字体定死（拉丁 run 的核心盒总是矮于同字号东亚 run 的，怎么合成都看不出来）。
 * 要照出它，需要两款**上下互不相让**的字体在同一行里：
 *
 * | 36pt | 基线以上 | 基线以下 | 单倍行高 |
 * |---|---|---|---|
 * | 等线（东亚规则） | 34.79 | 13.98 | 48.77 |
 * | 宋体（东亚规则） | **36.34** | 10.46 | 46.80 |
 * | 两者同行，Word 实测 | **36.31** | 13.97 | **50.28** |
 *
 * 也就是说 Word 给的行高比**两款字体各自的行高都大**（50.28 > 48.77 > 46.80）——
 * 原来的 `naturalLineHeight`「取各自行高的最大值」按定义就说不出这个数，差 1.51pt。
 * 上取宋体、下取等线，得 36.34 + 13.98 = 50.32（twips 取整后 50.31），对上了。
 *
 * 注意每款字体是在**自己的**自然行高里居中的（不是在合成之后的行高里）——
 * 后者会让矮的那一款把基线往下拽，P9 会算成 36.94（差 0.63pt）。
 * 合成之后多出来的空间（行距倍数、网格吸附）仍旧上下均分，那条规则不受影响。
 *
 * `maxHeight` 是落选的那一条，留着给 `spike:script` 重跑组合用。
 */
export function composeLineBox(parts: readonly LineMetrics[], rule: ComposeRule = 'maxSides'): LineBox {
  if (rule === 'maxHeight') {
    const natural = naturalLineHeight(parts);
    return { natural, above: composeBaseline(parts, natural) };
  }
  let above = 0;
  let below = 0;
  for (const p of parts) {
    const a = baselineOffset(p, p.lineHeight);
    if (a > above) above = a;
    const b = p.lineHeight - a;
    if (b > below) below = b;
  }
  return { natural: above + below, above };
}

// ── fontkit 的类型补丁 ────────────────────────────────────────────────────────
// @types/fontkit 没有暴露原始表，这里只声明我们真正读的字段。

interface Os2Table {
  winAscent: number;
  winDescent: number;
  typoAscender: number;
  typoDescender: number;
  typoLineGap: number;
  fsSelection?: number;
}

interface HheaTable {
  ascent: number;
  descent: number;
  lineGap: number;
}

export interface FontkitFont {
  familyName?: string;
  postscriptName?: string;
  unitsPerEm: number;
  hhea?: unknown;
  'OS/2'?: unknown;
  /** cmap 查询与字形推进宽度。度量包抽取与 ①级度量都靠这两个 */
  hasGlyphForCodePoint?: (cp: number) => boolean;
  glyphForCodePoint?: (cp: number) => { advanceWidth?: number } | null;
}
