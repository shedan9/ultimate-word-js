/**
 * 行高规则的回归测试。
 *
 * 这里的期望值不是从代码推的，是从 Word 实测反推的 ——
 * 见 apps/fidelity 的 spike-lineheight-01 / 02 与 spike-lineheight.ts。
 * 改动 lineMetrics 前先想清楚：真值站在测试这一边。
 */

import { ptToTwips, twipsToPt } from '@uw/core';
import { describe, expect, it } from 'vitest';
import { combineLineMetrics, gdiExternalLeading, lineMetrics, type RawFontMetrics } from './metrics.ts';

/** 仿宋 / 宋体 / 黑体 / 楷体：unitsPerEm=256，win 跨度恰好 1.0 em */
const fangSong: RawFontMetrics = {
  family: 'FangSong',
  postscriptName: 'FangSong',
  unitsPerEm: 256,
  os2: {
    winAscent: 220,
    winDescent: 36,
    typoAscender: 220,
    typoDescender: -36,
    typoLineGap: 36,
    useTypoMetrics: false,
  },
  hhea: { ascender: 220, descender: -36, lineGap: 36 },
};

/** Times New Roman：win 跨度 1.1074 em，另有 87/2048 的 GDI 外部行距 */
const timesNewRoman: RawFontMetrics = {
  family: 'Times New Roman',
  postscriptName: 'TimesNewRomanPSMT',
  unitsPerEm: 2048,
  os2: {
    winAscent: 1825,
    winDescent: 443,
    typoAscender: 1420,
    typoDescender: -442,
    typoLineGap: 307,
    useTypoMetrics: false,
  },
  hhea: { ascender: 1825, descender: -443, lineGap: 87 },
};

/** 微软雅黑：win 跨度 1.3198 em，无 lineGap —— 用来把「1.3 系数」与「固定 1.3 倍字号」区分开 */
const yaHei: RawFontMetrics = {
  family: 'Microsoft YaHei',
  postscriptName: 'MicrosoftYaHei',
  unitsPerEm: 2048,
  os2: {
    winAscent: 2167,
    winDescent: 536,
    typoAscender: 2167,
    typoDescender: -536,
    typoLineGap: 0,
    useTypoMetrics: false,
  },
  hhea: { ascender: 2167, descender: -536, lineGap: 0 },
};

const heightPt = (m: RawFontMetrics, sizePt: number, eastAsian: boolean): number =>
  twipsToPt(lineMetrics(m, ptToTwips(sizePt), { eastAsian }).lineHeight);

describe('GDI 外部行距', () => {
  it('lineGap 超出 win 与 hhea 跨度之差的部分才算数', () => {
    expect(gdiExternalLeading(timesNewRoman)).toBe(87);
  });

  it('win 跨度已经把 lineGap 吃掉时为 0', () => {
    // 仿宋：win 跨度 256 = hhea 跨度 256，lineGap 36 全额进外部行距
    expect(gdiExternalLeading(fangSong)).toBe(36);
    expect(gdiExternalLeading(yaHei)).toBe(0);
  });
});

describe('单倍行距行高（对齐 Word 实测）', () => {
  it('东亚文字：win 跨度 × 1.3，不加外部行距', () => {
    // 实测 20.76pt @ 15.96pt 字号
    expect(heightPt(fangSong, 15.96, true)).toBeCloseTo(20.75, 2);
    // 实测 20.52 ~ 20.64pt @ 12pt
    expect(heightPt(yaHei, 12, true)).toBeCloseTo(20.59, 2);
  });

  it('拉丁文字：win 跨度 + 外部行距，没有 1.3 系数', () => {
    // 实测 13.80pt @ 12pt
    expect(heightPt(timesNewRoman, 12, false)).toBeCloseTo(13.8, 2);
  });

  it('1.3 是乘在字体度量上，不是「行高 = 1.3 × 字号」', () => {
    // 若是后者，雅黑与仿宋在同字号下行高会相同 —— 实测差了 58%
    expect(heightPt(yaHei, 12, true)).toBeGreaterThan(heightPt(fangSong, 12, true) * 1.3);
  });

  it('typo / hhea 两条对照路径可用', () => {
    const typo = lineMetrics(timesNewRoman, ptToTwips(12), { source: 'typo' });
    const hhea = lineMetrics(timesNewRoman, ptToTwips(12), { source: 'hhea' });
    expect(twipsToPt(typo.lineHeight)).toBeCloseTo((12 * (1420 + 442 + 307)) / 2048, 4);
    expect(twipsToPt(hhea.lineHeight)).toBeCloseTo((12 * (1825 + 443 + 87)) / 2048, 4);
  });
});

describe('混排行', () => {
  it('ascent 与 descent 逐项取最大，可能来自不同字体', () => {
    const ea = lineMetrics(fangSong, ptToTwips(12), { eastAsian: true });
    const latin = lineMetrics(timesNewRoman, ptToTwips(12), { eastAsian: false });
    const combined = combineLineMetrics([ea, latin]);
    expect(combined.ascent).toBe(Math.max(ea.ascent, latin.ascent));
    expect(combined.descent).toBe(Math.max(ea.descent, latin.descent));
    expect(combined.lineHeight).toBe(combined.ascent + combined.descent + combined.lineGap);
  });
});
