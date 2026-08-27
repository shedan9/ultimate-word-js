/**
 * 表格边框的冲突解析。
 *
 * 这里的期望值有两个来源，读的时候要分清：
 * - **层级覆盖**（单元格盖表级）有 ECMA-376 §17.4.39 撑着，是规范
 * - **相邻竞争**（谁赢）已经用 `spike-table-03` 实测（21 组 × 横竖两遍全对），
 *   规则与证据表见 table-borders.ts 的 `BORDER_CONFLICT_RULES`；
 *   逐条边的回归在 table-border-fixture.test.ts。这里只挑规则的几个拐点做单测
 */
import type { Border, CellBorders, TableBorders } from '@uw/model';
import { describe, expect, it } from 'vitest';
import type { BorderCell, BorderRow } from './table-borders.ts';
import { resolveTableBorders } from './table-borders.ts';

/** `size` 直接写 twips（`w:sz` 的 1/8 磅在解析处就转掉了）：10 = 0.5pt 细实线 */
function b(style: string, size = 10, color = 'auto'): Border {
  return { style, size, space: 0, color };
}

function c(borders: CellBorders = {}, over: Partial<BorderCell> = {}): BorderCell {
  return { gridSpan: 1, vMerge: 'none', borders, ...over };
}

function r(cells: BorderCell[], gridBefore = 0): BorderRow {
  return { gridBefore, cells };
}

/** 最常见的表：四周 1pt 粗、内部 0.5pt 细 */
const FRAMED: TableBorders = {
  top: b('single', 20),
  left: b('single', 20),
  bottom: b('single', 20),
  right: b('single', 20),
  insideH: b('single', 10),
  insideV: b('single', 10),
};

/** 只取水平边的第一段 —— 不分段的用例里它就是整条边 */
const topOf = (seg: { border: Border | undefined }[]) => seg[0]?.border;

describe('层级：单元格覆盖表级', () => {
  const grid3 = [r([c(), c(), c()]), r([c(), c(), c()]), r([c(), c(), c()])];

  it('贴着外沿的用 top/left/bottom/right，内部的用 insideH/insideV', () => {
    const out = resolveTableBorders(grid3, FRAMED, 3);

    // 左上角那格：上、左是外框（粗），下、右是内部线（细）
    const nw = out[0]?.[0];
    expect(topOf(nw?.top ?? [])?.size).toBe(20);
    expect(nw?.left?.size).toBe(20);
    expect(topOf(nw?.bottom ?? [])?.size).toBe(10);
    expect(nw?.right?.size).toBe(10);

    // 正中那格：四边全是内部线
    const mid = out[1]?.[1];
    expect(topOf(mid?.top ?? [])?.size).toBe(10);
    expect(mid?.left?.size).toBe(10);
    expect(topOf(mid?.bottom ?? [])?.size).toBe(10);
    expect(mid?.right?.size).toBe(10);

    // 右下角那格：下、右是外框
    const se = out[2]?.[2];
    expect(topOf(se?.bottom ?? [])?.size).toBe(20);
    expect(se?.right?.size).toBe(20);
  });

  it('单元格写了这条边就用它，不再看表级', () => {
    const rows = [r([c({ right: b('double', 40) }), c()]), r([c(), c()])];
    const out = resolveTableBorders(rows, FRAMED, 2);
    expect(out[0]?.[0]?.right?.style).toBe('double');
    expect(out[0]?.[0]?.right?.size).toBe(40);
  });

  it('表级什么都没定义时四周都不画', () => {
    const out = resolveTableBorders([r([c(), c()])], {}, 2);
    expect(topOf(out[0]?.[0]?.top ?? [])).toBeUndefined();
    expect(out[0]?.[0]?.left).toBeUndefined();
    expect(out[0]?.[1]?.right).toBeUndefined();
  });
});

describe('nil 的两副面孔', () => {
  it('单格写 nil 能盖掉表级 insideV —— 但邻格的 single 还在，线照画', () => {
    // 这条是两级模型的核心：合成一步（「nil 一律赢」）会让整表的内部格线被一格抹掉
    const rows = [r([c({ right: b('nil', 0) }), c()])];
    const out = resolveTableBorders(rows, FRAMED, 2);
    expect(out[0]?.[0]?.right?.size).toBe(10);
    expect(out[0]?.[1]?.left?.size).toBe(10);
  });

  it('两格都说 nil 时线才真的消失', () => {
    const rows = [r([c({ right: b('nil', 0) }), c({ left: b('nil', 0) })])];
    const out = resolveTableBorders(rows, FRAMED, 2);
    expect(out[0]?.[0]?.right).toBeUndefined();
    expect(out[0]?.[1]?.left).toBeUndefined();
  });

  it('none 与 nil 在「画不画」上同义', () => {
    const rows = [r([c({ right: b('none', 0) }), c({ left: b('nil', 0) })])];
    expect(resolveTableBorders(rows, FRAMED, 2)[0]?.[0]?.right).toBeUndefined();
  });
});

describe('竞争规则（spike-table-03 实测）', () => {
  it('同类先比线宽 —— 粗的赢', () => {
    const rows = [r([c({ right: b('single', 40) }), c({ left: b('single', 10) })])];
    const out = resolveTableBorders(rows, FRAMED, 2);
    expect(out[0]?.[0]?.right?.size).toBe(40);
    // 两侧解析出同一条，渲染层画两遍是幂等的
    expect(out[0]?.[1]?.left?.size).toBe(40);
  });

  it('同宽时 double 赢 single —— 它画出来是 sz 的三倍厚', () => {
    const rows = [r([c({ right: b('single', 20) }), c({ left: b('double', 20) })])];
    expect(resolveTableBorders(rows, FRAMED, 2)[0]?.[0]?.right?.style).toBe('double');
  });

  it('双线按画出来的厚度比：0.75pt 的双线（2.25pt 厚）输给 3pt 的单线', () => {
    // 实测的第「九」组。按 w:sz 直接比的话 15 < 120 也是单线赢，分不开两种算法；
    // 分得开的是第「丁」组：1.5pt 双线（4.5pt 厚）赢过 3pt 单线
    const nine = [r([c({ right: b('double', 15) }), c({ left: b('single', 60) })])];
    expect(resolveTableBorders(nine, FRAMED, 2)[0]?.[0]?.right?.style).toBe('single');
    const ding = [r([c({ right: b('single', 60) }), c({ left: b('double', 30) })])];
    expect(resolveTableBorders(ding, FRAMED, 2)[0]?.[0]?.right?.style).toBe('double');
  });

  it('厚度一样时样式再比一次 —— 0.5pt 的双线赢过 1.5pt 的单线', () => {
    // 实测的第「癸」组：两条都画出来 1.5pt 厚，位置规则（左上）本该让 single 赢
    const rows = [r([c({ right: b('single', 30) }), c({ left: b('double', 10) })])];
    expect(resolveTableBorders(rows, FRAMED, 2)[0]?.[0]?.right?.style).toBe('double');
  });

  it('破折类再宽也输给实线 —— 3pt 的点线输给 0.75pt 的单线', () => {
    // 实测的第「八」/「丙」组。原来照 CSS 写的「先比线宽」在这里给的是相反的答案
    const dot = [r([c({ right: b('dotted', 60) }), c({ left: b('single', 15) })])];
    expect(resolveTableBorders(dot, FRAMED, 2)[0]?.[0]?.right?.style).toBe('single');
    const dash = [r([c({ right: b('dashed', 60) }), c({ left: b('single', 10) })])];
    expect(resolveTableBorders(dash, FRAMED, 2)[0]?.[0]?.right?.style).toBe('single');
  });

  it('dashed 赢 dotted', () => {
    const d = [r([c({ right: b('dotted', 20) }), c({ left: b('dashed', 20) })])];
    expect(resolveTableBorders(d, FRAMED, 2)[0]?.[0]?.right?.style).toBe('dashed');
  });

  it('同一种破折线之间不比宽度 —— 细的那条在左上就归它', () => {
    // 实测的第「戊」/「壬」/「辛」组：dotted 0.5 与 dotted 2.25 互换位置，两次都是左上赢
    const thin = [r([c({ right: b('dotted', 10, 'FF0000') }), c({ left: b('dotted', 45, '0000FF') })])];
    expect(resolveTableBorders(thin, FRAMED, 2)[0]?.[0]?.right?.color).toBe('FF0000');
    const thick = [r([c({ right: b('dotted', 45, 'FF0000') }), c({ left: b('dotted', 10, '0000FF') })])];
    expect(resolveTableBorders(thick, FRAMED, 2)[0]?.[0]?.right?.color).toBe('FF0000');
  });

  it('认不出的线型落到 single 那一档，不会输给 dashed', () => {
    const rows = [r([c({ right: b('dashDotStroked', 20) }), c({ left: b('dashed', 20) })])];
    expect(resolveTableBorders(rows, FRAMED, 2)[0]?.[0]?.right?.style).toBe('dashDotStroked');
  });

  it('全平局时取左上者 —— 垂直边归左边那格', () => {
    const rows = [r([c({ right: b('single', 20, 'FF0000') }), c({ left: b('single', 20, '0000FF') })])];
    const out = resolveTableBorders(rows, FRAMED, 2);
    expect(out[0]?.[0]?.right?.color).toBe('FF0000');
    expect(out[0]?.[1]?.left?.color).toBe('FF0000');
  });

  it('全平局时取左上者 —— 水平边归上面那行', () => {
    const rows = [r([c({ bottom: b('single', 20, 'FF0000') })]), r([c({ top: b('single', 20, '0000FF') })])];
    const out = resolveTableBorders(rows, FRAMED, 1);
    expect(topOf(out[0]?.[0]?.bottom ?? [])?.color).toBe('FF0000');
    expect(topOf(out[1]?.[0]?.top ?? [])?.color).toBe('FF0000');
  });
});

describe('水平边分段', () => {
  it('上下格子边界对不齐时按列切段，各段各比各的', () => {
    // 第一行一格通栏跨 3 列，第二行 3 格，只有中间那格的 top 是粗的
    const rows = [r([c({}, { gridSpan: 3 })]), r([c(), c({ top: b('single', 40) }), c()])];
    const out = resolveTableBorders(rows, FRAMED, 3);

    const bottom = out[0]?.[0]?.bottom ?? [];
    expect(bottom.map((s) => [s.col, s.span, s.border?.size])).toEqual([
      [0, 1, 10],
      [1, 1, 40],
      [2, 1, 10],
    ]);
  });

  it('结果相同的相邻段合并成一条，不留一列宽的碎片', () => {
    const rows = [r([c({}, { gridSpan: 3 })]), r([c(), c(), c()])];
    const bottom = resolveTableBorders(rows, FRAMED, 3)[0]?.[0]?.bottom ?? [];
    expect(bottom).toEqual([{ col: 0, span: 3, border: b('single', 10) }]);
  });
});

describe('合并与空缺', () => {
  it('vMerge=continue 与上格之间不画线（合并区内部）', () => {
    const rows = [
      r([c(), c()]),
      r([c({}, { vMerge: 'restart' }), c()]),
      r([c({}, { vMerge: 'continue' }), c()]),
    ];
    const out = resolveTableBorders(rows, FRAMED, 2);
    // 合并区内部那条：restart 的下边与 continue 的上边都不画
    expect(topOf(out[1]?.[0]?.bottom ?? [])).toBeUndefined();
    expect(topOf(out[2]?.[0]?.top ?? [])).toBeUndefined();
    // 同一条线在右边那列照画 —— 不画只是合并区自己的事
    expect(topOf(out[1]?.[1]?.bottom ?? [])?.size).toBe(10);
  });

  it('合并区的上沿与下沿照画', () => {
    const rows = [r([c({}, { vMerge: 'restart' })]), r([c({}, { vMerge: 'continue' })])];
    const out = resolveTableBorders(rows, FRAMED, 1);
    expect(topOf(out[0]?.[0]?.top ?? [])?.size).toBe(20);
    expect(topOf(out[1]?.[0]?.bottom ?? [])?.size).toBe(20);
  });

  it('gridBefore 跳掉的列是空缺：第一格不算首列，走 insideV 而不是外框', () => {
    const rows = [r([c(), c()]), r([c()], 1)];
    const out = resolveTableBorders(rows, FRAMED, 2);
    expect(out[1]?.[0]?.left?.size).toBe(10);
    // 它右边贴着表格右沿，那条仍是外框
    expect(out[1]?.[0]?.right?.size).toBe(20);
  });
});

describe('对角线', () => {
  it('直接取本格的，不参与竞争', () => {
    const rows = [r([c({ tl2br: b('single', 20), tr2bl: b('nil', 0) })])];
    const out = resolveTableBorders(rows, FRAMED, 1);
    expect(out[0]?.[0]?.tl2br?.size).toBe(20);
    expect(out[0]?.[0]?.tr2bl).toBeUndefined();
  });
});
