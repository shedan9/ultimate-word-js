/**
 * 段落装配：缩进（含字符单位）、对齐、行高（含行网格）、渲染片段。
 *
 * 版心宽度一律取 `SIZE_5 * 10` —— 「一行 10 个字」，期望值可以数着字写。
 */

import type { LineMetrics, TextMeasurer } from '@uw/fonts';
import type { DocGrid } from '@uw/model';
import { DEFAULT_SETTINGS } from '@uw/model';
import { describe, expect, it } from 'vitest';
import type { LayoutParagraphOptions } from './paragraph.ts';
import { layoutParagraph } from './paragraph.ts';
import { fakeMeasurer, NO_GRID, para, paraProps, run, runOf, SIZE_5 } from './test-fixtures.ts';

const WIDTH = SIZE_5 * 10;

function opts(over: Partial<LayoutParagraphOptions> = {}): LayoutParagraphOptions {
  return {
    measurer: fakeMeasurer(),
    contentWidth: WIDTH,
    settings: DEFAULT_SETTINGS,
    docGrid: NO_GRID,
    ...over,
  };
}

/** 东亚行的单倍行高：(0.8 + 0.2) em × 1.3 */
const EA_LINE = SIZE_5 * 1.3;

describe('缩进', () => {
  it('首行缩进 2 字符 = 2 个字号宽，且第一行的可用宽度跟着少 2 个字', () => {
    const p = para([run('一二三四五六七八九十一二')], {
      indent: { ...paraProps().indent, firstLineChars: 200 },
    });
    const out = layoutParagraph(p, opts());
    expect(out.lines[0]?.x).toBe(SIZE_5 * 2);
    expect(out.lines[0]?.end).toBe(8); // 10 - 2
    expect(out.lines[1]?.x).toBe(0);
  });

  it('字符单位优先于 twips —— Word 为兼容旧版会把两个都写进去', () => {
    const p = para([run('一二')], {
      indent: { ...paraProps().indent, firstLine: 999, firstLineChars: 200 },
    });
    expect(layoutParagraph(p, opts()).lines[0]?.x).toBe(SIZE_5 * 2);
  });

  it('悬挂缩进让首行往左伸出去，其余行按 left 排', () => {
    const p = para([run('一二三四五六七八九十一二')], {
      indent: { ...paraProps().indent, left: SIZE_5 * 2, hanging: SIZE_5 },
    });
    const out = layoutParagraph(p, opts());
    expect(out.lines[0]?.x).toBe(SIZE_5);
    expect(out.lines[1]?.x).toBe(SIZE_5 * 2);
  });

  it('左右缩进一起吃掉可用宽度', () => {
    const p = para([run('一二三四五六七八九十')], {
      indent: { ...paraProps().indent, left: SIZE_5, right: SIZE_5 },
    });
    expect(layoutParagraph(p, opts()).lines[0]?.end).toBe(8);
  });
});

describe('对齐', () => {
  const short = () => para([run('一二')], { justification: 'center' });

  it('居中：整行平移，不动行内的字距', () => {
    const out = layoutParagraph(short(), opts());
    expect(out.lines[0]?.x).toBe((WIDTH - SIZE_5 * 2) / 2);
    expect(out.lines[0]?.fragments[0]?.glyphX).toEqual([840, 1050]);
  });

  it('右对齐：行尾贴住版心右边', () => {
    const p = para([run('一二')], { justification: 'right' });
    expect(layoutParagraph(p, opts()).lines[0]?.x).toBe(WIDTH - SIZE_5 * 2);
  });

  it('两端对齐把多余宽度摊进字间，但**不拉伸最后一行**', () => {
    const p = para([run('一二三四五六七八九十一二')], { justification: 'both' });
    const out = layoutParagraph(p, opts({ contentWidth: SIZE_5 * 10.5 }));
    expect(out.lines[0]?.width).toBe(SIZE_5 * 10.5); // 10 个字摊开占满 10.5 字宽
    expect(out.lines[1]?.width).toBe(SIZE_5 * 2); // 末行照常
    const glyphs = out.lines[0]?.fragments[0]?.glyphX ?? [];
    expect(glyphs[1] as number).toBeGreaterThan(SIZE_5); // 字与字之间张开了
  });

  it('分散对齐连最后一行也拉 —— 这正是它与两端对齐的区别', () => {
    const p = para([run('一二')], { justification: 'distribute' });
    expect(layoutParagraph(p, opts()).lines[0]?.width).toBe(WIDTH);
  });

  it('行尾空格不参与居中计算', () => {
    const p = para([run('一二 ')], { justification: 'center' });
    expect(layoutParagraph(p, opts()).lines[0]?.x).toBe((WIDTH - SIZE_5 * 2) / 2);
  });
});

describe('行高', () => {
  const heightOf = (over = {}, o: Partial<LayoutParagraphOptions> = {}) =>
    layoutParagraph(para([run('一二')], over), opts(o)).lines[0]?.height;

  it('单倍行距走 Phase 0 标定的东亚公式', () => {
    expect(heightOf()).toBe(EA_LINE);
  });

  it('lineRule=auto 时 w:line 是 1/240 行 —— 360 就是 1.5 倍', () => {
    expect(heightOf({ spacing: { ...paraProps().spacing, line: 360 } })).toBe(EA_LINE * 1.5);
  });

  it('lineRule=exact 时 w:line 是 twips，固定值说了算', () => {
    expect(heightOf({ spacing: { ...paraProps().spacing, line: 300, lineRule: 'exact' } })).toBe(300);
  });

  it('lineRule=atLeast 取自然行高与设定值的较大者', () => {
    const spacing = { ...paraProps().spacing, lineRule: 'atLeast' as const };
    expect(heightOf({ spacing: { ...spacing, line: 500 } })).toBe(500);
    expect(heightOf({ spacing: { ...spacing, line: 100 } })).toBe(EA_LINE);
  });

  it('行网格：行高吸到 linePitch 的整数倍 —— 公文「每页 22 行」靠它', () => {
    const grid: DocGrid = { type: 'lines', linePitch: 312, charSpace: 0 };
    expect(heightOf({}, { docGrid: grid })).toBe(312);
    // 段落关掉 snapToGrid 就不吸
    expect(heightOf({ snapToGrid: false }, { docGrid: grid })).toBe(EA_LINE);
    // 固定值行距也不吸：用户既然写死了行高，网格不该再改它
    expect(
      heightOf({ spacing: { ...paraProps().spacing, line: 300, lineRule: 'exact' } }, { docGrid: grid }),
    ).toBe(300);
  });

  it('网格吸附在行距倍数**之前** —— 于是 1.5 倍行距的行高与字号无关', () => {
    // 实测（spike-baseline-03 末三段）：网格 31.8pt 下开 1.5 倍行距，仿宋 16pt 与宋体 12pt
    // 的行高都是 47.7pt = 1.5 个网格行。反过来（先乘倍数再吸附）会得到 1 个网格行，
    // 而且结果随字号变。这里用两个字号一起断言，正是为了让顺序搞反时必然有一个挂掉。
    const grid: DocGrid = { type: 'lines', linePitch: 312, charSpace: 0 };
    const spacing = { ...paraProps().spacing, line: 360 };
    expect(heightOf({ spacing }, { docGrid: grid })).toBe(312 * 1.5);
    const line = layoutParagraph(
      para([run('一二', { size: SIZE_5 * 2 })], { spacing }),
      opts({ docGrid: grid }),
    ).lines[0];
    // 两倍字号的自然行高 546 > 312，吸到两个网格行再乘 1.5
    expect(line?.height).toBe(312 * 2 * 1.5);
  });

  it('纯拉丁行不走 1.3 系数', () => {
    expect(layoutParagraph(para([run('ab')]), opts()).lines[0]?.height).toBe(SIZE_5);
  });

  it('空段落的行高取段落标记的字符属性，且走 ascii 桶 + 拉丁规则', () => {
    // 实测（spike-baseline-01 末两页）：只有段落标记的空段落，行高是标记 **ascii** 字体的
    // 拉丁行高，不是 eastAsia 字体的 1.3 倍 —— 段落标记本身不是东亚字符。
    // 合成度量器里拉丁行高恰好是 1 em，所以期望值是 SIZE_5 而不是 EA_LINE。
    const out = layoutParagraph(para([]), opts());
    expect(out.lines).toHaveLength(1);
    expect(out.lines[0]?.height).toBe(SIZE_5);
  });

  it('段前间距的行单位按行高折算', () => {
    const p = para([run('一')], { spacing: { ...paraProps().spacing, beforeLines: 100, before: 999 } });
    expect(layoutParagraph(p, opts()).spaceBefore).toBe(EA_LINE);
  });
});

/**
 * 一款「比东亚字体伸得更高」的拉丁字体，用来测混排行的合成规则。
 *
 * 合成度量器默认所有字体同形（0.8 / 0.2 em），那样测不出「谁定行盒」——
 * 必须让拉丁一侧在参与合成时会**赢**，才能证明它没有参与。
 * 参数取 0.9 / 0.25 em 是算好的：拉丁自然行高 241.5 < 东亚的 273（不影响行高），
 * 核心盒上沿 189 < 东亚基线 199.5（不触发防切字下限），
 * 但若参与居中就会给出 204.75 —— 与实测的 199.5 差得出来。
 */
function tallLatinMeasurer(ascentEm: number, descentEm: number): TextMeasurer {
  const base = fakeMeasurer();
  return {
    ...base,
    lineMetrics(family, fontSize, o = {}): LineMetrics {
      if (family !== 'Times New Roman') return base.lineMetrics(family, fontSize, o);
      const ascent = fontSize * ascentEm;
      const descent = fontSize * descentEm;
      return { ascent, descent, lineGap: 0, lineHeight: ascent + descent, coreAbove: ascent };
    },
  };
}

describe('基线在行盒里的位置', () => {
  const lineOf = (text: string, over = {}, o: Partial<LayoutParagraphOptions> = {}) =>
    layoutParagraph(para([run(text)], over), opts(o)).lines[0];

  it('东亚行：额外的 30% 上下均分，基线落在 0.95 em', () => {
    const line = lineOf('一二');
    expect(line?.height).toBe(EA_LINE);
    expect(line?.baseline).toBe(SIZE_5 * 0.95);
  });

  it('拉丁行：外部行距全在基线以上，合成字体里它是 0，所以基线就是 ascent', () => {
    const line = lineOf('ab');
    expect(line?.height).toBe(SIZE_5);
    expect(line?.baseline).toBe(SIZE_5 * 0.8);
  });

  it('网格吸附多出来的空间上下均分 —— 行高变了，字体度量没变', () => {
    const grid: DocGrid = { type: 'lines', linePitch: 312, charSpace: 0 };
    const line = lineOf('一二', {}, { docGrid: grid });
    expect(line?.height).toBe(312);
    // ascent 168 + (312 − 核心盒 210) / 2
    expect(line?.baseline).toBe(168 + (312 - SIZE_5) / 2);
  });

  it('固定值行距把行压矮时基线跟着上移，且永不超出行高', () => {
    const line = lineOf('一二', { spacing: { ...paraProps().spacing, line: 100, lineRule: 'exact' } });
    expect(line?.height).toBe(100);
    expect(line?.baseline).toBeLessThanOrEqual(100);
  });

  it('自然行高与最终行高分开记 —— 行高不对时要能分辨是度量的锅还是规则的锅', () => {
    const grid: DocGrid = { type: 'lines', linePitch: 312, charSpace: 0 };
    const line = lineOf('一二', {}, { docGrid: grid });
    expect(line?.natural).toBe(EA_LINE);
    expect(line?.height).toBe(312);
  });
});

describe('混排行的行盒', () => {
  it('东亚行的行盒只由东亚字体定，拉丁 run 不参与', () => {
    // 实测（spike-baseline-02 的「等 Tj 等」）：混排行的基线与同字号纯东亚行**一模一样**。
    const measurer = tallLatinMeasurer(0.9, 0.25);
    const mixed = layoutParagraph(para([run('一a')]), opts({ measurer })).lines[0];
    const pure = layoutParagraph(para([run('一二')]), opts({ measurer })).lines[0];
    expect(mixed?.height).toBe(pure?.height);
    expect(mixed?.baseline).toBe(pure?.baseline);
    // 若拉丁也参与居中，基线会被顶到 204.75
    expect(mixed?.baseline).not.toBe(189 + (EA_LINE - (189 + SIZE_5 * 0.25)) / 2);
  });

  it('拉丁一侧高到会被切字时，下限把行撑开 —— 这是判断不是实测', () => {
    // 2.0 em 的 ascent 远超东亚基线，样本外的情形（12pt 汉字里嵌 72pt 英文）就长这样。
    // 没有真值，所以只保证「不切字」，不保证与 Word 一致，见 line-height.ts 的 floorBox。
    const measurer = tallLatinMeasurer(2, 0.5);
    const line = layoutParagraph(para([run('一a')]), opts({ measurer })).lines[0];
    expect(line?.baseline).toBe(SIZE_5 * 2);
    expect(line?.height).toBeGreaterThanOrEqual((line?.baseline ?? 0) + SIZE_5 * 0.5);
  });
});

describe('渲染片段', () => {
  it('同 run 同字体同字号的连续字符合成一段，中英交界处切开', () => {
    const out = layoutParagraph(para([run('中a')]), opts());
    const [cn, en] = out.lines[0]?.fragments ?? [];
    expect(cn?.text).toBe('中');
    expect(en?.text).toBe('a');
    // 中西文自动间距（1/4 em，实测，见 WIDTH_RULES）体现在拉丁段的起点上
    expect(en?.x).toBe(SIZE_5 + SIZE_5 / 4);
  });

  it('glyphX 与 text 的码点一一对应，供 SVG text 的 x 数组直接使用', () => {
    const out = layoutParagraph(para([run('一二三')]), opts());
    const f = out.lines[0]?.fragments[0];
    expect(f?.glyphX).toEqual([0, SIZE_5, SIZE_5 * 2]);
    expect([...(f?.text ?? '')]).toHaveLength(f?.glyphX.length ?? 0);
  });

  it('制表位的前导符记下起止 x —— 目录那排点靠它', () => {
    const p = para([runOf([{ kind: 'tab' }, { kind: 'text', text: '中' }])], {
      tabs: [{ pos: 1000, alignment: 'left', leader: 'dot' }],
    });
    expect(layoutParagraph(p, opts()).lines[0]?.leaders).toEqual([{ x1: 0, x2: 1000, leader: 'dot' }]);
  });

  it('布局结果整个可结构化克隆 —— 它要过 Worker 边界', () => {
    const out = layoutParagraph(para([run('中文abc')]), opts());
    expect(structuredClone(out)).toEqual(out);
  });
});
