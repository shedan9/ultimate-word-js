/**
 * 表格的水平几何。
 *
 * 度量器是合成的（东亚字 1 em、ASCII 半角），所以「一格里排得下几个字」可以手算 ——
 * 期望值写的是算出来的数，不是跑一遍抄回来的。
 */
import { DEFAULT_SETTINGS } from '@uw/model';
import { describe, expect, it } from 'vitest';
import type { LayoutTableOptions } from './table.ts';
import { layoutTable } from './table.ts';
import {
  cell,
  cellProps,
  fakeMeasurer,
  NO_GRID,
  para,
  row,
  rowProps,
  run,
  SIZE_5,
  table,
  tableProps,
} from './test-fixtures.ts';

function opts(availWidth: number): LayoutTableOptions {
  return { measurer: fakeMeasurer(), availWidth, settings: DEFAULT_SETTINGS, docGrid: NO_GRID };
}

/** 版心 8306 twips（A4 减 GB/T 9704 的页边距那个量级），列宽用整数好手算 */
const AVAIL = 9000;

const textCell = (text: string) => cell([para([run(text)])]);

describe('列宽', () => {
  it('w:tblGrid 是权威 —— 直接用它，一行内容都不用测', () => {
    const t = table([2000, 3000, 4000], [row([textCell('甲'), textCell('乙'), textCell('丙')])]);
    const laid = layoutTable(t, opts(AVAIL));
    expect(laid.columns).toEqual([2000, 3000, 4000]);
    expect(laid.width).toBe(9000);
    expect(laid.rows[0]?.cells.map((c) => c.x)).toEqual([0, 2000, 5000]);
  });

  it('grid 缺席时拿某一行的 w:tcW 反推', () => {
    const t = table(
      [],
      [
        row([
          cell([para([run('甲')])], { props: cellProps({ width: { value: 2500, type: 'dxa' } }) }),
          cell([para([run('乙')])], { props: cellProps({ width: { value: 6500, type: 'dxa' } }) }),
        ]),
      ],
    );
    expect(layoutTable(t, opts(AVAIL)).columns).toEqual([2500, 6500]);
  });

  it('grid 与 tcW 都没有时按整表宽度等分', () => {
    const t = table([], [row([textCell('甲'), textCell('乙'), textCell('丙')])], {
      props: tableProps({ width: { value: 2500, type: 'pct' } }), // 2500/5000 = 50%
    });
    const laid = layoutTable(t, opts(AVAIL));
    expect(laid.columns).toEqual([1500, 1500, 1500]);
    expect(laid.width).toBe(4500);
  });

  it('跨列的格子把几列的宽度加起来', () => {
    const t = table(
      [2000, 3000, 4000],
      [row([cell([para([run('通栏')])], { gridSpan: 2 }), textCell('丙')])],
    );
    const cells = layoutTable(t, opts(AVAIL)).rows[0]?.cells ?? [];
    expect(cells[0]?.width).toBe(5000);
    expect(cells[0]?.span).toBe(2);
    // 下一格的起始列是 2（不是 1）—— 列号按 gridSpan 累加
    expect(cells[1]?.col).toBe(2);
    expect(cells[1]?.x).toBe(5000);
  });

  it('w:gridBefore 把本行整体右移，跳过的列照样占位置', () => {
    const t = table(
      [2000, 3000, 4000],
      [row([textCell('右边那格')], { props: rowProps({ gridBefore: 2 }) })],
    );
    const c = layoutTable(t, opts(AVAIL)).rows[0]?.cells[0];
    expect(c?.col).toBe(2);
    expect(c?.x).toBe(5000);
    expect(c?.width).toBe(4000);
  });
});

describe('整表位置', () => {
  const t = () => table([2000, 2000], [row([textCell('甲'), textCell('乙')])]);

  it('居中时表格在可用宽里居中', () => {
    const laid = layoutTable(t(), opts(AVAIL));
    expect(laid.x).toBe(0);
    const centered = layoutTable(
      table([2000, 2000], [row([textCell('甲'), textCell('乙')])], {
        props: tableProps({ justification: 'center' }),
      }),
      opts(AVAIL),
    );
    expect(centered.x).toBe((9000 - 4000) / 2);
  });

  it('左对齐时 w:tblInd 生效', () => {
    const laid = layoutTable(
      table([2000, 2000], [row([textCell('甲'), textCell('乙')])], {
        props: tableProps({ indent: { value: 720, type: 'dxa' } }),
      }),
      opts(AVAIL),
    );
    expect(laid.x).toBe(720);
  });
});

describe('格内可用宽度', () => {
  it('可用宽 = 列宽减左右边距，段落就在这个宽度里断行', () => {
    const t = table([2000, 3000], [row([textCell('甲'), textCell('乙')])]);
    const c = layoutTable(t, opts(AVAIL)).rows[0]?.cells[0];
    expect(c?.paddingLeft).toBe(108);
    expect(c?.contentWidth).toBe(2000 - 108 - 108);
    const line = c?.blocks[0];
    expect(line?.kind).toBe('paragraph');
    expect(line?.kind === 'paragraph' ? line.layout.contentWidth : 0).toBe(1784);
  });

  it('格内的字按格宽断行 —— 1784 twips 装得下 8 个五号字', () => {
    // 五号字 210 twips/字，1784 / 210 = 8.49 → 每行 8 个
    const t = table([2000], [row([textCell('一二三四五六七八九十')])]);
    const c = layoutTable(t, opts(AVAIL)).rows[0]?.cells[0];
    const block = c?.blocks[0];
    if (block?.kind !== 'paragraph') throw new Error('格里第一个块不是段落');
    expect(block.layout.lines.length).toBe(2);
    expect(block.layout.lines[0]?.width).toBe(8 * SIZE_5);
  });

  it('边距宽过格子时可用宽夹到 0，不让负数传进断行', () => {
    const narrow = cell([para([run('挤')])], {
      props: cellProps({
        margins: {
          top: { value: 0, type: 'dxa' },
          left: { value: 500, type: 'dxa' },
          bottom: { value: 0, type: 'dxa' },
          right: { value: 500, type: 'dxa' },
        },
      }),
    });
    const c = layoutTable(table([600], [row([narrow])]), opts(AVAIL)).rows[0]?.cells[0];
    expect(c?.contentWidth).toBe(0);
  });
});

describe('行高（总量，没有基线）', () => {
  it('auto 时按内容撑开，取最高那一格', () => {
    const t = table([2000, 2000], [row([textCell('一二三四五六七八九十'), textCell('短')])]);
    const r = layoutTable(t, opts(AVAIL)).rows[0];
    // 左边那格排了两行，右边一行 —— 行高跟着高的那个
    const tall = r?.cells[0]?.contentHeight ?? 0;
    const short = r?.cells[1]?.contentHeight ?? 0;
    expect(tall).toBeGreaterThan(short);
    expect(r?.height).toBe(tall);
  });

  it('atLeast 取内容与声明值的较大者，exact 压着内容不放大', () => {
    const at = (rule: 'auto' | 'atLeast' | 'exact', value: number) =>
      layoutTable(
        table([2000], [row([textCell('甲')], { props: rowProps({ height: { value, rule } }) })]),
        opts(AVAIL),
      ).rows[0]?.height ?? 0;

    const content = at('auto', 0);
    expect(content).toBeGreaterThan(0);

    // 声明值高过内容：两种规则都听声明的
    expect(at('atLeast', 5000)).toBe(5000);
    expect(at('exact', 5000)).toBe(5000);
    // 声明值低于内容：atLeast 让内容撑开，exact 把内容压着（Word 会裁掉超出的部分）
    expect(at('atLeast', 1)).toBe(content);
    expect(at('exact', 1)).toBe(1);
  });

  it('vMerge=continue 的格子不参与行高 —— 它由上面那个 restart 撑着', () => {
    const long = cell([para([run('一二三四五六七八九十十一十二')])], { vMerge: 'continue' });
    const r = layoutTable(table([2000, 2000], [row([textCell('甲'), long])]), opts(AVAIL)).rows[0];
    expect(r?.height).toBe(r?.cells[0]?.contentHeight);
  });
});

describe('嵌套表格', () => {
  it('内层表格在外层格子的可用宽里重排', () => {
    const inner = table([1000, 1000], [row([textCell('内甲'), textCell('内乙')])]);
    const outer = table([3000, 3000], [row([cell([inner]), textCell('外乙')])]);
    const laid = layoutTable(outer, opts(AVAIL));
    const block = laid.rows[0]?.cells[0]?.blocks[0];
    if (block?.kind !== 'table') throw new Error('外层格子里第一个块不是表格');
    expect(block.layout.width).toBe(2000);
    // 内层的 x 是相对**内层表格自己**的，外层格子的 padding 不叠进来
    expect(block.layout.rows[0]?.cells[1]?.x).toBe(1000);
  });
});

describe('边框接线', () => {
  /** 规则本身在 table-borders.test.ts 里测，这里只验证接线：结果挂对了格子、列数没数错 */
  const line = (size: number) => ({ style: 'single', size, space: 0, color: 'auto' });

  it('冲突解析的结果挂在每格上，末列认得出自己贴着外框', () => {
    const t = table([3000, 3000, 3000], [row([textCell('甲'), textCell('乙'), textCell('丙')])], {
      props: tableProps({
        borders: { left: line(20), right: line(20), insideV: line(10) },
      }),
    });
    const cells = layoutTable(t, opts(AVAIL)).rows[0]?.cells ?? [];
    expect(cells[0]?.borders.left?.size).toBe(20);
    expect(cells[0]?.borders.right?.size).toBe(10);
    // 这一条同时在验 colCount 取的是列宽数组的长度 —— 数错的话末列会拿到 insideV
    expect(cells[2]?.borders.right?.size).toBe(20);
  });
});

describe('可结构化克隆（原则 1.1）', () => {
  it('整个 TableLayout 能过 Worker 边界', () => {
    const t = table([2000, 2000], [row([textCell('甲'), cell([table([500], [row([textCell('内')])])])])]);
    const laid = layoutTable(t, opts(AVAIL));
    expect(() => structuredClone(laid)).not.toThrow();
    expect(structuredClone(laid)).toEqual(laid);
  });
});
