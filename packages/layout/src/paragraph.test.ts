/**
 * 段落装配：缩进（含字符单位）、对齐、行高（含行网格）、渲染片段。
 *
 * 版心宽度一律取 `SIZE_5 * 10` —— 「一行 10 个字」，期望值可以数着字写。
 */

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

  it('纯拉丁行不走 1.3 系数', () => {
    expect(layoutParagraph(para([run('ab')]), opts()).lines[0]?.height).toBe(SIZE_5);
  });

  it('空段落的行高取段落标记的字符属性 —— 它不是摆设', () => {
    const out = layoutParagraph(para([]), opts());
    expect(out.lines).toHaveLength(1);
    expect(out.lines[0]?.height).toBe(EA_LINE);
  });

  it('段前间距的行单位按行高折算', () => {
    const p = para([run('一')], { spacing: { ...paraProps().spacing, beforeLines: 100, before: 999 } });
    expect(layoutParagraph(p, opts()).spaceBefore).toBe(EA_LINE);
  });
});

describe('渲染片段', () => {
  it('同 run 同字体同字号的连续字符合成一段，中英交界处切开', () => {
    const out = layoutParagraph(para([run('中a')]), opts());
    const [cn, en] = out.lines[0]?.fragments ?? [];
    expect(cn?.text).toBe('中');
    expect(en?.text).toBe('a');
    // 中西文自动间距体现在拉丁段的起点上
    expect(en?.x).toBe(SIZE_5 + SIZE_5 / 8);
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
