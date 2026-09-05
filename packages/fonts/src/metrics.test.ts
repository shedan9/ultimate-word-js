/**
 * 行高规则的回归测试。
 *
 * 这里的期望值不是从代码推的，是从 Word 实测反推的 ——
 * 见 apps/fidelity 的 spike-lineheight-01 / 02 与 spike-lineheight.ts。
 * 改动 lineMetrics 前先想清楚：真值站在测试这一边。
 */

import { ptToTwips, twipsToPt } from '@uw/core';
import { describe, expect, it } from 'vitest';
import {
  baselineOffset,
  composeLineBox,
  gdiExternalLeading,
  lineMetrics,
  naturalLineHeight,
  type RawFontMetrics,
} from './metrics.ts';

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

/** 等线：win 跨度 1.0420 em，无 lineGap —— 现代中文字体里 win 跨度最接近 1 em 的一款 */
const dengXian: RawFontMetrics = {
  family: 'DengXian',
  postscriptName: 'DengXian',
  unitsPerEm: 2048,
  os2: {
    winAscent: 1659,
    winDescent: 475,
    typoAscender: 1659,
    typoDescender: -475,
    typoLineGap: 0,
    useTypoMetrics: false,
  },
  hhea: { ascender: 1659, descender: -475, lineGap: 0 },
};

/** Arial：win 跨度 1.1172 em，GDI 外部行距 67/2048 */
const arial: RawFontMetrics = {
  family: 'Arial',
  postscriptName: 'ArialMT',
  unitsPerEm: 2048,
  os2: {
    winAscent: 1854,
    winDescent: 434,
    typoAscender: 1491,
    typoDescender: -431,
    typoLineGap: 307,
    useTypoMetrics: false,
  },
  hhea: { ascender: 1854, descender: -434, lineGap: 67 },
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

/**
 * 基线在行高里的位置。
 *
 * 期望值全部是 Word 实测的「首行基线 − 版心顶」，来自 apps/fidelity 的
 * spike-baseline-01 / 02 / 03（每段用 pageBreakBefore 顶到自己那页最上面，
 * 于是这个差就是基线在行盒里的位置，与前面排了什么无关）。
 *
 * 容差 0.15pt：实测残差最大 0.13pt，与 Phase 0 的 0.132pt 同量级，都来自 Word 那一侧的取整。
 * 想收紧容差得先解释掉那个系统偏差（26 个样本的残差**全为负**），不是改这里的数字。
 * 用 `toBeCloseTo` 表达不了 0.15 这个刻度（它只认 10 的幂），所以显式写差值。
 */
const TOLERANCE_PT = 0.15;

const expectPt = (actual: number, measured: number): void => {
  expect(Math.abs(actual - measured)).toBeLessThan(TOLERANCE_PT);
};

/** 行顶到基线（pt）。`heightPtValue` 缺省表示自然行高，给了就是被规则 / 网格拉大之后的行高 */
const baselinePt = (
  m: RawFontMetrics,
  sizePt: number,
  eastAsian: boolean,
  heightPtValue?: number,
): number => {
  const lm = lineMetrics(m, ptToTwips(sizePt), { eastAsian });
  return twipsToPt(
    baselineOffset(lm, heightPtValue === undefined ? lm.lineHeight : ptToTwips(heightPtValue)),
  );
};

describe('基线位置：自然行盒（对齐 Word 实测）', () => {
  it('东亚：额外的 30% 上下均分', () => {
    // spike-baseline-01：仿宋 16pt 实测 16.050、宋体 12pt 12.070、黑体 22pt 22.170
    // （这四款字体的 win 度量完全相同，所以同一份 RawFontMetrics 就代表了它们全部）
    expectPt(baselinePt(fangSong, 16, true), 16.05);
    expectPt(baselinePt(fangSong, 12, true), 12.07);
    expectPt(baselinePt(fangSong, 22, true), 22.17);
    // spike-baseline-02 把字号放大到 48–72pt，信号是噪声的十几倍
    expectPt(baselinePt(fangSong, 72, true), 72.57);
    expectPt(baselinePt(fangSong, 48, true), 48.33);
    // 雅黑 60pt 实测 75.330、等线 72pt 实测 69.570 —— 两者 win 跨度差 27%，
    // 能把「1.3 乘在字体度量上」与「1.3 乘在字号上」彻底分开
    expectPt(baselinePt(yaHei, 60, true), 75.33);
    expectPt(baselinePt(dengXian, 72, true), 69.57);
  });

  it('拉丁：GDI 外部行距全在基线以上', () => {
    // Times 12 / 48 / 72pt 实测 11.110 / 44.730 / 67.170
    expectPt(baselinePt(timesNewRoman, 12, false), 11.11);
    expectPt(baselinePt(timesNewRoman, 48, false), 44.73);
    expectPt(baselinePt(timesNewRoman, 72, false), 67.17);
    // Arial 12 / 72pt 实测 11.230 / 67.530
    expectPt(baselinePt(arial, 12, false), 11.23);
    expectPt(baselinePt(arial, 72, false), 67.53);
  });

  it('拉丁的外部行距若也上下均分，72pt 上要差 1.5pt —— 当初 12pt 样本分辨不出来的就是这个', () => {
    const lm = lineMetrics(timesNewRoman, ptToTwips(72), { eastAsian: false });
    // 半个外部行距在 72pt 上是 1.53pt，与实测值比是 1.48pt（预测本身还差 0.05）
    expect(Math.abs(twipsToPt(lm.ascent + lm.lineGap / 2) - 67.17)).toBeGreaterThan(1.4);
    // 同一个错误在 12pt 上只差 0.26pt，与坐标噪声同量级
    const small = lineMetrics(timesNewRoman, ptToTwips(12), { eastAsian: false });
    expect(Math.abs(twipsToPt(small.ascent + small.lineGap / 2) - 11.11)).toBeLessThan(0.3);
  });

  it('东亚的额外行距若全在基线以上，宋体 72pt 上要差 10pt', () => {
    const lm = lineMetrics(fangSong, ptToTwips(72), { eastAsian: true });
    expect(Math.abs(twipsToPt(lm.ascent + lm.lineGap) - 72.57)).toBeGreaterThan(10);
  });
});

describe('基线位置：行高被拉大之后（对齐 Word 实测）', () => {
  // spike-baseline-03：行网格 linePitch 636 twips = 31.8pt
  it('网格吸附多出来的空间上下均分', () => {
    // 仿宋 16pt 吸到 31.8pt，实测基线 21.570；宋体 12pt 也吸到 31.8pt，实测 20.250
    expectPt(baselinePt(fangSong, 16, true, 31.8), 21.57);
    expectPt(baselinePt(fangSong, 12, true, 31.8), 20.25);
    // 黑体 26pt 的自然行高 33.8pt 超过一个网格行，吸到两行 63.6pt，实测 41.130
    expectPt(baselinePt(fangSong, 26, true, 63.6), 41.13);
  });

  it('拉丁行的吸附余量也上下均分 —— 与自然行盒里「外部行距全在上」是两件事', () => {
    // Times 12pt 吸到 31.8pt，实测 20.130；若吸附余量也全在上会是 29.20pt
    expectPt(baselinePt(timesNewRoman, 12, false, 31.8), 20.13);
    const lm = lineMetrics(timesNewRoman, ptToTwips(12), { eastAsian: false });
    const allAbove = twipsToPt(lm.coreAbove + (ptToTwips(31.8) - (lm.coreAbove + lm.descent)));
    expect(Math.abs(allAbove - 20.13)).toBeGreaterThan(8);
  });

  it('倍数行距放大的余量上下均分', () => {
    // 网格 31.8pt + 1.5 倍行距：行高 47.7pt（先吸附再乘，所以与字号无关）
    // 仿宋 16pt 实测 29.490、宋体 12pt 实测 28.170、2.0 倍（63.6pt）实测 37.530
    expectPt(baselinePt(fangSong, 16, true, 47.7), 29.49);
    expectPt(baselinePt(fangSong, 12, true, 47.7), 28.17);
    expectPt(baselinePt(fangSong, 16, true, 63.6), 37.53);
  });
});

describe('混排行的合成', () => {
  it('单字体的行盒恒等于它自己的行高与基线 —— 两种合成写法在这一格上同解', () => {
    const ea = lineMetrics(fangSong, ptToTwips(12), { eastAsian: true });
    for (const rule of ['maxSides', 'maxHeight'] as const) {
      const boxOf = composeLineBox([ea], rule);
      expect(boxOf.natural).toBe(ea.lineHeight);
      expect(boxOf.above).toBe(baselineOffset(ea, ea.lineHeight));
    }
  });

  it('拉丁字体走拉丁规则，所以它参与合成也顶不高东亚一侧的基线', () => {
    // spike-baseline-02 的「等 Tj 等」那一页：等线 72pt 与 Times New Roman 72pt 同排一行，
    // 实测首行基线 69.570 —— 与同字号的**纯等线**那一页一模一样。
    //
    // 原来的解释是「拉丁 run 完全不参与行盒」，`spike-script-01` 把它讲对了：
    // 它参与了，但它作为拉丁字体走**拉丁规则**（外部行距整块在基线以上、没有那 30%），
    // 核心盒上沿只有 67.22pt，赢不过等线的 69.57pt。
    const ea = lineMetrics(dengXian, ptToTwips(72), { eastAsian: true });
    const latin = lineMetrics(timesNewRoman, ptToTwips(72), { eastAsian: false });
    expectPt(twipsToPt(composeLineBox([ea, latin], 'maxSides').above), 69.57);
    expectPt(twipsToPt(composeLineBox([ea, latin], 'maxSides').natural), twipsToPt(ea.lineHeight));

    // 若照旧实现给拉丁一侧也套东亚规则（那 30% 让它的 win 跨度整整多出 6%），
    // 基线被顶高 6pt 以上。这一页是全部混排样本里唯一能分开两种做法的。
    const wrong = lineMetrics(timesNewRoman, ptToTwips(72), { eastAsian: true });
    expect(Math.abs(twipsToPt(composeLineBox([ea, wrong], 'maxSides').above) - 69.57)).toBeGreaterThan(6);
  });

  it('两款东亚字体同行：上取最高、下取最深，行高比两款各自的都大', () => {
    // spike-script-01 的 P9/P10/P11：等线画 ASCII、宋体画汉字，36pt。
    // Word 实测行高 50.28pt、首行基线距版心顶 36.31pt。
    const deng = lineMetrics(dengXian, ptToTwips(36), { eastAsian: true });
    // 宋体与仿宋的 win 度量一模一样（都是 unitsPerEm=256 / 220 / 36），P10/P11 用的正是仿宋
    const song = lineMetrics(fangSong, ptToTwips(36), { eastAsian: true });
    const box = composeLineBox([deng, song], 'maxSides');
    expectPt(twipsToPt(box.natural), 50.28);
    expectPt(twipsToPt(box.above), 36.31);

    // 「取各自行高的最大值」按定义给不出这个数：两款各自只有 48.77 与 46.80
    expect(twipsToPt(box.natural)).toBeGreaterThan(twipsToPt(Math.max(deng.lineHeight, song.lineHeight)));
    expect(twipsToPt(naturalLineHeight([deng, song]))).toBeCloseTo(48.77, 1);
  });
});
