/**
 * 表格的**水平几何**：列宽 → 每格的 x 与可用宽 → 格内段落布局。
 *
 * 与段落那条路一样，这里做完的是「行盒之前」的部分：格子有 x、有宽、格内的段落
 * 都排成了行，**没有 y、没有行高分配**。表格的高度方向要等基线穿刺
 * （见 `@uw/fonts` 的 metrics.ts），`contentHeight` 只给出**总量**，不含基线位置。
 *
 * ## 列宽为什么可以这么简单
 *
 * `w:tblLayout` 的 `autofit` 真算起来是一整套「最小 / 最大内容宽度 + 迭代分配」的算法
 * （已列为非目标）。但对**加载已有 docx** 这个场景有一条捷径：
 * **Word 存盘时把 autofit 的结果写进了 `w:tblGrid`** —— 那串 `w:gridCol` 就是 Word
 * 自己算完的列宽。照着用，得到的就是与 Word 一致的列宽，一行内容都不用测。
 *
 * 所以这里的策略是：**`w:tblGrid` 是权威**，`w:tblW` / `w:tcW` 只在它缺席时才上场。
 * 代价是「用户在我们这儿改了单元格内容之后列宽不会自动重算」—— 那属于编辑态
 * （Phase 7）要补的东西，不是加载路径的缺陷。
 *
 * ## 未做（写下来免得以为已经做了）
 *
 * - `w:tblCellSpacing`（单元格间距）**不消费**，几何按 0 算。它会同时改变整表宽度与
 *   每格的 x，猜一个测不了的实现不如留个洞（原则 1.5）。真实公文里几乎不用
 * - `w:tblW` 与 `w:tblGrid` 之和冲突时以 grid 为准，没有真值验证过 Word 到底听谁的
 * - 边框宽度不吃可用宽（Word 把边框画在格线上，不缩文字区）—— 同样没有真值。
 *   **画哪条线**已经解出来了（`table-borders.ts` 的冲突解析，挂在 `CellLayout.borders`），
 *   它不改任何坐标，所以不受基线穿刺阻塞
 */
import type { Twips } from '@uw/core';
import type {
  CellMargins,
  NodeId,
  ResolvedBlock,
  ResolvedTable,
  ResolvedTableCell,
  Shading,
  TableWidth,
} from '@uw/model';
import type { LayoutParagraphOptions } from './paragraph.ts';
import { layoutParagraph } from './paragraph.ts';
import type { CellBorderLayout } from './table-borders.ts';
import { borderRowsOf, resolveTableBorders } from './table-borders.ts';
import type { ParagraphLayout } from './types.ts';

/** 格内的块：段落排成行，嵌套表格递归下去 */
export type BlockLayout =
  | { kind: 'paragraph'; layout: ParagraphLayout }
  | { kind: 'table'; layout: TableLayout };

export interface CellLayout {
  cellId: NodeId;
  /** 起始网格列与占的列数 */
  col: number;
  span: number;
  /** 格子左边相对**表格左边**的 x */
  x: Twips;
  /** 格子外框宽（跨列的已经加起来了），**含**单元格边距 */
  width: Twips;
  /** 格内文字能用的宽度 = `width` 减左右边距。段落就是在这个宽度里断行的 */
  contentWidth: Twips;
  /** 左边距，格内内容的 x 要从它起算 */
  paddingLeft: Twips;
  paddingRight: Twips;
  /**
   * 上下边距。**不影响断行**（宽度才影响），但格内内容从哪个 y 起排全靠它 ——
   * 原先只把它折进 `contentHeight` 的总量里，渲染层就还原不出来了
   */
  paddingTop: Twips;
  paddingBottom: Twips;
  /** `w:vAlign`：内容不满一格时贴哪边。同样只有画的时候用得上 */
  verticalAlign: 'top' | 'center' | 'bottom';
  /** `w:shd` 原样带着（不解析成 RGB，与 model 一致），渲染层在边框之下铺一层底 */
  shading: Shading | undefined;
  /** `continue` 的格子渲染层不画内容、也不参与行高 —— 上面那个 `restart` 撑着它 */
  vMerge: 'none' | 'restart' | 'continue';
  /**
   * 冲突解析完的四条边（见 table-borders.ts）。相邻两格共享的那条线在两边都会
   * 出现且**解析结果相同**，渲染层画两遍是幂等的 —— 需要去重时按格线位置归并即可。
   */
  borders: CellBorderLayout;
  blocks: BlockLayout[];
  /**
   * 格内内容的高度**总量**（段前后间距 + 各行行高）。
   * 行高总量在 Phase 0 已标定，所以这个数是准的；但基线在行高里的位置还没定，
   * 所以它只能用来比大小（比如与 `w:trHeight` 取 max），**还不能拿来画**。
   */
  contentHeight: Twips;
}

export interface RowLayout {
  rowId: NodeId;
  cells: CellLayout[];
  /** 本行最高那一格的内容高度。行高规则（`w:trHeight`）已经并进来 */
  height: Twips;
}

export interface TableLayout {
  tableId: NodeId;
  /** 表格左边相对**版心左边**的 x（`w:tblInd` 与 `w:jc` 都算进去了） */
  x: Twips;
  /** 整表宽度 = 各列宽之和 */
  width: Twips;
  /** 每个网格列的宽度，与 `w:tblGrid` 一一对应 */
  columns: Twips[];
  rows: RowLayout[];
}

export interface LayoutTableOptions extends Omit<LayoutParagraphOptions, 'contentWidth'> {
  /** 表格能用的宽度：版心宽，嵌套时是外层单元格的 `contentWidth` */
  availWidth: Twips;
}

/**
 * `w:tblW` / `w:tcW` 那套宽度 → twips。
 *
 * `pct` 的刻度是 **1/50 个百分点**（5000 = 100%），除以 50 才是百分数。
 * `auto` 与 `nil` 都不给具体值，交给调用方决定退到什么。
 */
export function widthToTwips(w: TableWidth | undefined, avail: Twips): Twips | undefined {
  if (w === undefined) return undefined;
  if (w.type === 'dxa') return w.value;
  if (w.type === 'pct') return (w.value / 5000) * avail;
  if (w.type === 'nil') return 0;
  return undefined;
}

/** 边距只认 `dxa` / `nil`：`pct` 型的单元格边距相对谁没有定论，按没写处理 */
function marginOf(m: CellMargins[keyof CellMargins]): Twips {
  if (m === undefined) return 0;
  return m.type === 'dxa' ? m.value : 0;
}

export function layoutTable(t: ResolvedTable, opts: LayoutTableOptions): TableLayout {
  const columns = columnWidths(t, opts.availWidth);
  const width = columns.reduce((a, b) => a + b, 0);

  // 边框与几何互不影响（线不吃可用宽），所以两边各算各的，最后按下标对上
  const borders = resolveTableBorders(borderRowsOf(t.rows), t.props.borders, columns.length);

  const rows = t.rows.map((r, ri): RowLayout => {
    // 被 w:gridBefore 跳掉的列照样占位置 —— 第一格的 x 要从它们之后起算
    let col = r.props.gridBefore;
    const cells = r.cells.map((c, ci): CellLayout => {
      const cell = layoutCell(c, columns, col, opts, borders[ri]?.[ci] ?? NO_BORDERS);
      col += c.gridSpan;
      return cell;
    });
    return { rowId: r.id, cells, height: rowHeight(r.props.height, cells) };
  });

  return { tableId: t.id, x: tableX(t, width, opts.availWidth), width, columns, rows };
}

/**
 * 每个网格列的宽度。
 *
 * `w:tblGrid` 在就直接用（见文件头：那是 Word 算完的结果）。它不在时才退到
 * 「整表宽度 ÷ 列数」等分 —— 这一步只在手写 XML 或第三方生成器产的文件上才走到，
 * Word 自己存的文件总是带 `w:tblGrid`。
 */
export function columnWidths(t: ResolvedTable, avail: Twips): Twips[] {
  const declared = widthToTwips(t.props.width, avail);
  const grid = t.grid.filter((w) => w > 0);
  if (grid.length === t.grid.length && grid.length > 0) return [...t.grid];

  // grid 缺席 / 有零列：先看每行的 w:tcW 能不能凑出列宽，凑不出就等分
  const count = Math.max(t.grid.length, columnCountOf(t));
  if (count === 0) return [];
  const fromCells = widthsFromCells(t, count, avail);
  if (fromCells !== undefined) return fromCells;

  const total = declared ?? avail;
  return new Array<Twips>(count).fill(total / count);
}

function columnCountOf(t: ResolvedTable): number {
  let max = 0;
  for (const r of t.rows) {
    let n = r.props.gridBefore + r.props.gridAfter;
    for (const c of r.cells) n += c.gridSpan;
    if (n > max) max = n;
  }
  return max;
}

/**
 * 拿某一行的 `w:tcW` 反推列宽。
 *
 * 只认**每格都写了 dxa 宽度且不跨列**的行 —— 跨列的格子只知道合起来多宽，
 * 拆回每列要解方程，而这种文件本来就少见，不值得为它引入一套求解。
 */
function widthsFromCells(t: ResolvedTable, count: number, avail: Twips): Twips[] | undefined {
  for (const r of t.rows) {
    if (r.props.gridBefore > 0 || r.props.gridAfter > 0) continue;
    if (r.cells.length !== count) continue;
    if (r.cells.some((c) => c.gridSpan !== 1)) continue;
    const widths = r.cells.map((c) => widthToTwips(c.props.width, avail));
    if (widths.some((w) => w === undefined || w <= 0)) continue;
    return widths as Twips[];
  }
  return undefined;
}

/** `w:jc` 决定整表在可用宽里的位置；`w:tblInd` 只在左对齐时叠加 */
function tableX(t: ResolvedTable, width: Twips, avail: Twips): Twips {
  const jc = t.props.justification;
  if (jc === 'center') return (avail - width) / 2;
  if (jc === 'right') return avail - width;
  return widthToTwips(t.props.indent, avail) ?? 0;
}

/** 下标对不上时的兜底（正常走不到：边框结果与 `t.rows` 逐格同构） */
const NO_BORDERS: CellBorderLayout = {
  top: [],
  bottom: [],
  left: undefined,
  right: undefined,
  tl2br: undefined,
  tr2bl: undefined,
};

function layoutCell(
  c: ResolvedTableCell,
  columns: readonly Twips[],
  col: number,
  opts: LayoutTableOptions,
  borders: CellBorderLayout,
): CellLayout {
  let x = 0;
  for (let i = 0; i < col && i < columns.length; i++) x += columns[i] as Twips;
  let width = 0;
  for (let i = col; i < col + c.gridSpan && i < columns.length; i++) width += columns[i] as Twips;

  const paddingLeft = marginOf(c.props.margins.left);
  const paddingRight = marginOf(c.props.margins.right);
  const paddingTop = marginOf(c.props.margins.top);
  const paddingBottom = marginOf(c.props.margins.bottom);
  // 边距比格子还宽时可用宽度会变负 —— 夹到 0，别让负宽度传进断行算法
  const contentWidth = Math.max(0, width - paddingLeft - paddingRight);

  const blocks = c.blocks.map((b) => blockLayout(b, contentWidth, opts));
  return {
    cellId: c.id,
    col,
    span: c.gridSpan,
    x,
    width,
    contentWidth,
    paddingLeft,
    paddingRight,
    paddingTop,
    paddingBottom,
    verticalAlign: c.props.verticalAlign,
    shading: c.props.shading,
    vMerge: c.vMerge,
    borders,
    blocks,
    contentHeight: contentHeightOf(blocks) + paddingTop + paddingBottom,
  };
}

function blockLayout(b: ResolvedBlock, contentWidth: Twips, opts: LayoutTableOptions): BlockLayout {
  if (b.kind === 'table') {
    // 嵌套表格在外层格子的可用宽里重新排一遍
    return { kind: 'table', layout: layoutTable(b, { ...opts, availWidth: contentWidth }) };
  }
  return { kind: 'paragraph', layout: layoutParagraph(b, { ...opts, contentWidth }) };
}

/**
 * 一摞块的高度总量（段前后间距 + 各行行高 + 嵌套表格的行高）。
 *
 * 导出是给**渲染层**用的：格内的块自己不带 y（与段落同理，见 types.ts 的说明），
 * 要把内容按 `w:vAlign` 摆到格子里就得先知道这一摞有多高。两处各算一遍必然会漂。
 */
export function contentHeightOf(blocks: readonly BlockLayout[]): Twips {
  let h = 0;
  for (const b of blocks) {
    if (b.kind === 'table') {
      for (const r of b.layout.rows) h += r.height;
      continue;
    }
    h += b.layout.spaceBefore + b.layout.spaceAfter;
    for (const line of b.layout.lines) h += line.height;
  }
  return h;
}

/**
 * 行高：内容撑起来的高度与 `w:trHeight` 的关系。
 *
 * - `exact`：**就是**这个高度，内容再高也压着（Word 会把超出的内容裁掉）
 * - `atLeast`：取两者较大的
 * - `auto`：完全按内容
 *
 * 缺席时是 `atLeast`（见 `@uw/model` 的 parseRowProps），认成 exact 会把内容压扁。
 *
 * `vMerge="continue"` 的格子不参与：它的内容不显示，由上面那个 `restart` 撑着。
 * 反过来，`restart` 那一格的内容现在**整个算进起始行** —— 合并区跨几行、
 * 高度怎么分摊到各行还没有真值，先偏高不偏低（偏低会让文字被下一行盖住）。
 */
function rowHeight(
  h: { value: Twips; rule: 'auto' | 'atLeast' | 'exact' },
  cells: readonly CellLayout[],
): Twips {
  let content = 0;
  for (const c of cells) {
    if (c.vMerge === 'continue') continue;
    if (c.contentHeight > content) content = c.contentHeight;
  }
  if (h.rule === 'exact') return h.value;
  if (h.rule === 'atLeast') return Math.max(h.value, content);
  return content;
}
