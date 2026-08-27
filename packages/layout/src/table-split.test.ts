/**
 * 表格拆行。
 *
 * 版心与 `page.test.ts` 同构：一行 8 个字（格子可用宽 1884 = 2100 − 108 × 2）、
 * 一页 3 行（819 = 273 × 3）。期望值因此能数着行写出来 —— 「头片 3 行、尾片 2 行」
 * 失败时一眼看得出是差了一行还是差了一整页。
 *
 * 这里测的是**规则**，不是保真度。规则本身由 `spike-table-04` 实测（见
 * `table-split.ts` 的 `TABLE_SPLIT_RULES`），与真值逐页对的那一份回归在
 * `table-split-fixture.test.ts`；这里只钉住「规则落到代码上是不是这个意思」。
 */
import type { ResolvedBlock, ResolvedBody, SectionProps } from '@uw/model';
import { DEFAULT_SECTION_PROPS, DEFAULT_SETTINGS } from '@uw/model';
import { describe, expect, it } from 'vitest';
import type { DocumentLayout, LayoutDocumentOptions, PlacedRow, PlacedTable } from './page.ts';
import { layoutDocument } from './page.ts';
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

const EA_LINE = SIZE_5 * 1.3;

function sect(over: Partial<SectionProps> = {}): SectionProps {
  return {
    ...structuredClone(DEFAULT_SECTION_PROPS),
    page: { width: 3300, height: 2019, orientation: 'portrait' },
    margin: { top: 600, right: 600, bottom: 600, left: 600, header: 0, footer: 0, gutter: 0 },
    docGrid: NO_GRID,
    ...over,
  };
}

const body = (blocks: ResolvedBlock[]): ResolvedBody => ({
  sections: [{ id: 's0', props: sect(), blocks }],
});
const opts = (): LayoutDocumentOptions => ({ measurer: fakeMeasurer(), settings: DEFAULT_SETTINGS });

/** 一格 n 段，每段一行 —— 拆得开的最小单位就是「段落的一行」 */
const lines = (n: number, tag = '文') => cell(Array.from({ length: n }, (_, i) => para([run(`${tag}${i}`)])));

const tableOf = (doc: DocumentLayout, page: number): PlacedTable => doc.pages[page]?.blocks[0] as PlacedTable;

/** 一片里每个格子各留下几行（嵌套表格按行数算） */
function rowShape(placed: PlacedRow): number[] {
  return placed.row.cells.map((c) =>
    c.blocks.reduce((n, b) => n + (b.kind === 'table' ? b.layout.rows.length : b.layout.lines.length), 0),
  );
}

describe('表格拆行', () => {
  it('一行放不下时从行间切开：本页一片、下一页接一片', () => {
    const doc = layoutDocument(body([table([2100], [row([lines(5)])])]), opts());
    expect(doc.pages).toHaveLength(2);

    const head = tableOf(doc, 0);
    const tail = tableOf(doc, 1);
    expect(rowShape(head.rows[0] as PlacedRow)).toEqual([3]);
    expect(rowShape(tail.rows[0] as PlacedRow)).toEqual([2]);
    // 同一个 index 在两页各出现一次，靠 continued / splitAfter 分辨这是行的哪一截
    expect(head.rows[0]).toMatchObject({ index: 0, splitAfter: true });
    expect(head.rows[0]?.continued).toBeUndefined();
    expect(tail.rows[0]).toMatchObject({ index: 0, continued: true });
    expect(tail.rows[0]?.splitAfter).toBeUndefined();
    // 头片正好填满版心，尾片从页顶起
    expect(head.rows[0]?.height).toBe(3 * EA_LINE);
    expect(tail.rows[0]?.height).toBe(2 * EA_LINE);
    expect(tail.rows[0]?.y).toBe(0);
  });

  it('切口两侧的 first / last 说的是切口，不是表格的头尾', () => {
    const doc = layoutDocument(body([table([2100], [row([lines(5)])])]), opts());
    expect(tableOf(doc, 0)).toMatchObject({ first: true, last: false });
    expect(tableOf(doc, 1)).toMatchObject({ first: false, last: true });
  });

  it('每一格各切各的：这一片的高度按最高那一格算', () => {
    const doc = layoutDocument(body([table([2100, 2100], [row([lines(5, '甲'), lines(1, '乙')])])]), opts());
    expect(rowShape(tableOf(doc, 0).rows[0] as PlacedRow)).toEqual([3, 1]);
    // 乙格已经排完，尾片里它是个空壳（渲染层还要靠它画格线与底纹）
    expect(rowShape(tableOf(doc, 1).rows[0] as PlacedRow)).toEqual([2, 0]);
    expect(tableOf(doc, 1).rows[0]?.row.cells).toHaveLength(2);
  });

  it('一行高过一整页时接着切，不会溢出版心', () => {
    const doc = layoutDocument(body([table([2100], [row([lines(8)])])]), opts());
    expect(doc.pages).toHaveLength(3);
    expect(doc.pages.map((p) => rowShape((p.blocks[0] as PlacedTable).rows[0] as PlacedRow))).toEqual([
      [3],
      [3],
      [2],
    ]);
    // 中间那一片两头都是切口
    expect(tableOf(doc, 1).rows[0]).toMatchObject({ continued: true, splitAfter: true });
  });

  it('w:cantSplit 的行整行挪走，不切', () => {
    const t = table([2100], [row([lines(2)]), row([lines(3)], { props: rowProps({ cantSplit: true }) })]);
    const doc = layoutDocument(body([t]), opts());
    expect(doc.pages).toHaveLength(2);
    expect(tableOf(doc, 0).rows.map((r) => r.index)).toEqual([0]);
    const moved = tableOf(doc, 1).rows[0] as PlacedRow;
    expect(moved.index).toBe(1);
    expect(moved.continued).toBeUndefined();
    expect(rowShape(moved)).toEqual([3]);
  });

  it('表头行不切 —— 它每页都要重复一遍，半行表头没有意义', () => {
    const header = row([lines(4, '头')], { props: rowProps({ header: true }) });
    const doc = layoutDocument(body([table([2100], [header, row([lines(2)])])]), opts());
    const first = tableOf(doc, 0).rows[0] as PlacedRow;
    expect(first.splitAfter).toBeUndefined();
    expect(rowShape(first)).toEqual([4]); // 整行硬塞（空页上没有别处可挪）
  });

  it('w:trHeight 每一片各要一份，要不到就夹在这一页能给的高度里', () => {
    const tall = row([lines(5)], { props: rowProps({ height: { value: 2000, rule: 'atLeast' } }) });
    const doc = layoutDocument(body([table([2100], [tall])]), opts());
    const head = tableOf(doc, 0).rows[0] as PlacedRow;
    const tail = tableOf(doc, 1).rows[0] as PlacedRow;
    // 两片都要 2000，而一页只有 819 —— 于是两片各占满一整页（Word 的表乙就是这样）。
    // 按「整行算完富余归尾片」的老写法，头片只有 3 行高、尾片会长出 2000 − 5 行那一截
    expect(head.height).toBe(3 * EA_LINE);
    expect(tail.height).toBe(3 * EA_LINE);
    expect(rowShape(tail)).toEqual([2]);
  });

  it('w:trHeight 要的高度大过本页剩下的地方时整行挪走，不在这一页切', () => {
    const short = row([lines(1)]);
    const tall = row([lines(2)], { props: rowProps({ height: { value: 600, rule: 'atLeast' } }) });
    const doc = layoutDocument(body([table([2100], [short, tall])]), opts());
    // 首行占掉 273，只剩 546 < 600 —— 一片都满足不了下限，切了也白切
    expect(tableOf(doc, 0).rows.map((r) => r.index)).toEqual([0]);
    expect(tableOf(doc, 1).rows[0]).toMatchObject({ index: 1 });
    expect(tableOf(doc, 1).rows[0]?.continued).toBeUndefined();
  });

  it('单元格上下边距两片各补一整份', () => {
    const margins = cellProps({
      margins: {
        top: { value: 200, type: 'dxa' },
        left: { value: 108, type: 'dxa' },
        bottom: { value: 100, type: 'dxa' },
        right: { value: 108, type: 'dxa' },
      },
    });
    const c = cell(
      Array.from({ length: 2 }, (_, i) => para([run(`文${i}`)])),
      { props: margins },
    );
    const doc = layoutDocument(body([table([2100], [row([c])])]), opts());
    const head = tableOf(doc, 0).rows[0] as PlacedRow;
    const tail = tableOf(doc, 1).rows[0] as PlacedRow;
    // 上下边距各占一份之后只剩 519，装得下 1 行 —— 按「只补上边距」算能装下 2 行，
    // 那正是这条规则在真值里露馅的地方（Word 的表甲首页少放了整整两行）
    expect(rowShape(head)).toEqual([1]);
    expect(rowShape(tail)).toEqual([1]);
    expect(head.row.cells[0]).toMatchObject({ paddingTop: 200, paddingBottom: 100 });
    expect(tail.row.cells[0]).toMatchObject({ paddingTop: 200, paddingBottom: 100 });
    expect(head.height).toBe(200 + EA_LINE + 100);
    expect(tail.height).toBe(200 + EA_LINE + 100);
  });

  it('头片照样认 w:vAlign', () => {
    const c = cell([para([run('乙')])], { props: cellProps({ verticalAlign: 'bottom' }) });
    const doc = layoutDocument(body([table([2100, 2100], [row([lines(5, '甲'), c])])]), opts());
    const head = tableOf(doc, 0).rows[0] as PlacedRow;
    // 一律按 top 摆是原来的写法：真值里那一格的字落在片底，不在片顶
    expect(head.row.cells[1]?.verticalAlign).toBe('bottom');
  });

  it('接缝上那两条线取的是表级的上下边框，不是这一行自己的', () => {
    const borders = {
      top: { style: 'single' as const, size: 40, space: 0, color: '000000', shadow: false, frame: false },
      bottom: {
        style: 'single' as const,
        size: 60,
        space: 0,
        color: '000000',
        shadow: false,
        frame: false,
      },
      insideH: {
        style: 'single' as const,
        size: 8,
        space: 0,
        color: '000000',
        shadow: false,
        frame: false,
      },
    };
    const t = table([2100], [row([lines(1)]), row([lines(5)])], { props: tableProps({ borders }) });
    const doc = layoutDocument(body([t]), opts());
    const head = tableOf(doc, 1).rows[0] as PlacedRow; // 首页只放得下第一行
    const tail = tableOf(doc, 2).rows[0] as PlacedRow;
    expect(head.row.cells[0]?.borders.bottom[0]?.border?.size).toBe(60); // 表级 bottom
    expect(tail.row.cells[0]?.borders.top[0]?.border?.size).toBe(40); // 表级 top
    // 尾片顶上那条线像正常格线一样占高度
    expect(tail.row.gridAbove).toBe(40);
  });

  it('一行文字都放不下时不切，整行挪到下一页', () => {
    // 前面一行占掉 273，只剩 546 —— 上边距 500 之后只余 46，一行文字都放不下
    const c = cell([para([run('文')])], {
      props: cellProps({
        margins: {
          top: { value: 500, type: 'dxa' },
          left: { value: 108, type: 'dxa' },
          bottom: { value: 0, type: 'dxa' },
          right: { value: 108, type: 'dxa' },
        },
      }),
    });
    const doc = layoutDocument(body([table([2100], [row([lines(1)]), row([c])])]), opts());
    expect(tableOf(doc, 0).rows.map((r) => r.index)).toEqual([0]);
    // 整行挪到第 2 页，在那儿完整放下（500 + 273 ≤ 819），没有留下任何切口
    expect(doc.pages).toHaveLength(2);
    expect(tableOf(doc, 1).rows[0]?.continued).toBeUndefined();
    expect(tableOf(doc, 1).rows[0]?.splitAfter).toBeUndefined();
    expect(rowShape(tableOf(doc, 1).rows[0] as PlacedRow)).toEqual([1]);
  });

  it('嵌套表格按行原子处理，不再往下切', () => {
    const inner = table([1600], [row([lines(2, '内')]), row([lines(2, '里')])]);
    const outer = table([2100], [row([cell([inner])])]);
    const doc = layoutDocument(body([outer]), opts());
    expect(doc.pages).toHaveLength(2);
    // 内层两行各 2 行高：第一行放得下、第二行放不下 → 切在内层的两行之间
    expect(rowShape(tableOf(doc, 0).rows[0] as PlacedRow)).toEqual([1]);
    expect(rowShape(tableOf(doc, 1).rows[0] as PlacedRow)).toEqual([1]);
  });

  it('拆行不污染缓存的 TableLayout：同一份 body 排两遍结果一样', () => {
    const b = body([table([2100], [row([lines(5)])])]);
    const first = layoutDocument(b, opts());
    const again = layoutDocument(b, opts());
    const shapeOf = (doc: DocumentLayout) =>
      doc.pages.map((p) => rowShape((p.blocks[0] as PlacedTable).rows[0] as PlacedRow));
    expect(shapeOf(again)).toEqual(shapeOf(first));
  });
});
