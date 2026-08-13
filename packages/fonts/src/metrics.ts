/**
 * 字体度量 —— 整个保真度的地基。
 *
 * Word 的「单倍行距」**不是** CSS 的 `line-height: normal`：它取字体 OS/2 表的
 * usWinAscent / usWinDescent（部分字体退回 hhea），乘以字号得到行高。浏览器各引擎
 * 对 normal 的算法互不相同，所以这一步必须自己读字体表，不能问浏览器。
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
  /** 基线以上 */
  ascent: Twips;
  /** 基线以下（正数） */
  descent: Twips;
  /** 行间额外间隙 */
  lineGap: Twips;
  /** 单倍行距的行高 = ascent + descent + lineGap */
  lineHeight: Twips;
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
  /** 该行是否含东亚文字 —— 决定用 1.3 系数还是 GDI 外部行距 */
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
 * ⚠️ 未决：这 30% 的额外行距在基线上下如何分配（决定行内基线位置），
 * 需要另做一次「首行基线到版心顶」的穿刺来定，Phase 2 之前必须解决。
 * 现在一律记进 lineGap，只保证行高总量正确。
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
  return { ascent, descent, lineGap, lineHeight: ascent + descent + lineGap };
}

/**
 * 一行里混排多款字体时，行高取各字体行度量的**逐项最大值**，
 * 而不是最大行高 —— ascent 与 descent 可能来自不同字体。
 */
export function combineLineMetrics(parts: readonly LineMetrics[]): LineMetrics {
  let ascent = 0;
  let descent = 0;
  let lineGap = 0;
  for (const p of parts) {
    if (p.ascent > ascent) ascent = p.ascent;
    if (p.descent > descent) descent = p.descent;
    if (p.lineGap > lineGap) lineGap = p.lineGap;
  }
  return { ascent, descent, lineGap, lineHeight: ascent + descent + lineGap };
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
}
