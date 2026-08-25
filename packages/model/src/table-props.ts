/**
 * 表格 / 行 / 单元格的属性类型（Phase 4）。
 *
 * 与 `props.ts` 同一套纪律：分「部分属性」与「级联结果」两套，单位在解析处转 twips，
 * 全部可结构化克隆（原则 1.1）。
 *
 * 有一个单位**故意不转**，理由与 `w:*Chars` 同构：`TableWidth` 的 `pct` 型
 * 是相对**容器宽度**的，而容器宽度要到布局时才知道（嵌套表格里更是要一层层算下来）。
 * 在这里转等于把一个还不知道的数字硬写死。
 */
import type { Twips } from '@uw/core';
import type { Justification, ParaProps, RunProps } from './props.ts';

// ── 宽度 ──────────────────────────────────────────────────────────────────────

/**
 * `w:tblW` / `w:tcW` / `w:tblCellSpacing` / `w:tblInd` 共用的宽度写法。
 *
 * 四种刻度里 `pct` 最坑：它的值是 **1/50 个百分点**，`5000` = 100%，不是 `100`。
 * 认成百分数会让「宽度 100%」的表格缩成 2% 宽。
 */
export interface TableWidth {
  /** `dxa` 时是 twips（已转好，本来就是）；`pct` 时是 1/50 %；`auto` / `nil` 时无意义 */
  value: number;
  type: 'auto' | 'dxa' | 'pct' | 'nil';
}

// ── 边框与底纹 ────────────────────────────────────────────────────────────────

/**
 * 一条边框。
 *
 * `size` 的原始刻度是 **1/8 磅**（`w:sz="4"` = 0.5pt = 最常见的细实线），这里已转 twips，
 * 于是会出现 10 twips 这种非整数磅的值 —— 正常，不要再去取整。
 *
 * `style` 原样保留（`single` / `double` / `dotted` / `nil` / `none` …）：边框的**画法**是
 * 渲染层的事，布局只关心它占不占宽度。`nil` 与 `none` 都表示没有边框，但在
 * 边框冲突解析里两者优先级不同（`nil` 是「明确无」，`none` 是「未指定」），所以不合并。
 */
export interface Border {
  style: string;
  /** 线宽，已由 1/8 磅转 twips */
  size: Twips;
  /** `w:space`：边框与内容的间距，原始单位是**磅**（不是 1/8 磅），已转 twips */
  space: Twips;
  /** 六位十六进制或 `auto`。与主题色同理不解析成 RGB，见 props.ts 末尾 */
  color: string;
  /** `w:shadow` / `w:frame`：纯视觉，带着给渲染层 */
  shadow?: boolean;
}

/** 表级边框（`w:tblBorders`）。`insideH` / `insideV` 管的是格子之间那些线 */
export interface TableBorders {
  top?: Border;
  left?: Border;
  bottom?: Border;
  right?: Border;
  insideH?: Border;
  insideV?: Border;
}

/**
 * 单元格边框（`w:tcBorders`）。比表级多两条对角线，且**也有** `insideH` / `insideV` ——
 * 后者只在这个单元格被 `gridSpan` / `vMerge` 合并过时才有意义（合并区内部的线）。
 */
export interface CellBorders extends TableBorders {
  /** 左上到右下的对角线 */
  tl2br?: Border;
  /** 右上到左下 */
  tr2bl?: Border;
}

/**
 * `w:shd`：底纹。
 *
 * 对排版坐标零影响，所以和主题色一样**不解析成 RGB**，原样带给渲染层。
 * `fill` 才是背景色，`color` 是图案的前景色 —— 反了会让「浅色底纹」变成实心块。
 */
export interface Shading {
  /** 图案（`clear` = 纯色填充，`pct25` 之类是网点） */
  pattern: string;
  color: string;
  fill: string;
}

/** `w:tblCellMar` / `w:tcMar`：单元格四周的内边距。**直接吃掉可用宽度**，布局必须认 */
export interface CellMargins {
  top?: TableWidth;
  left?: TableWidth;
  bottom?: TableWidth;
  right?: TableWidth;
}

// ── 表格 ──────────────────────────────────────────────────────────────────────

/**
 * `w:tblLook`：哪几种条件格式生效。
 *
 * 它是**开关**不是格式本身：表格样式里可能定义了 `firstRow` 的加粗，但如果
 * `w:tblLook` 说 `firstRow="0"`，那份格式就不应用。漏了这一层，所有用了
 * 内置表格样式的表都会平白多出一行加粗表头。
 *
 * 两个 band 是**反的**：`noHBand=true` 表示**不要**行带。照字面存，别自作主张取反。
 */
export interface TableLook {
  firstRow: boolean;
  lastRow: boolean;
  firstColumn: boolean;
  lastColumn: boolean;
  noHBand: boolean;
  noVBand: boolean;
}

export interface TableProps {
  /** `w:tblStyle`：表格样式 id */
  styleId?: string;
  /** `w:tblW`：整表宽度 */
  width?: TableWidth;
  /** `w:jc`：整表在版心里的水平位置（不是单元格内文字的对齐） */
  justification?: Justification;
  /** `w:tblInd`：表格左边到版心左边的距离 */
  indent?: TableWidth;
  borders?: TableBorders;
  shading?: Shading;
  /** `w:tblCellMar`：表级默认单元格边距，可被 `w:tcMar` 逐格覆盖 */
  cellMargins?: CellMargins;
  /** `w:tblCellSpacing`：格与格之间的空隙（Word 里的「允许调整单元格间距」） */
  cellSpacing?: TableWidth;
  /**
   * `w:tblLayout`：`fixed` 按 `w:tblGrid` 与 `w:tcW` 排，`autofit` 要按内容算。
   * 缺席时是 `autofit` —— 这是规范默认，不是我们的选择。
   */
  layout?: 'fixed' | 'autofit';
  look?: TableLook;
  /** `w:tblStyleRowBandSize` / `ColBandSize`：一条带算几行 / 几列，默认 1 */
  rowBandSize?: number;
  colBandSize?: number;
}

// ── 行 ────────────────────────────────────────────────────────────────────────

/**
 * `w:trHeight`：行高。
 *
 * `rule` 缺席时是 `atLeast`（**最小值**），不是固定值 —— 认成 exact 会把内容压扁。
 * `auto` 表示完全按内容，此时 `value` 无意义。
 */
export interface RowHeight {
  value: Twips;
  rule: 'auto' | 'atLeast' | 'exact';
}

export interface RowProps {
  height?: RowHeight;
  /** `w:cantSplit`：这一行不许跨页拆开 */
  cantSplit?: boolean;
  /** `w:tblHeader`：本行是表头，跨页时在每页顶部重复 */
  header?: boolean;
  /** 行级的 `w:jc`：这一行相对整表的水平对齐（少见但合法） */
  justification?: Justification;
  /** 行级的单元格间距，覆盖表级 */
  cellSpacing?: TableWidth;
  /**
   * `w:gridBefore` / `w:gridAfter`：本行**跳过**开头 / 结尾几个网格列。
   * 这是「第一行少一格」那类表的实现方式，漏了会让整行的格子全部左移。
   */
  gridBefore?: number;
  gridAfter?: number;
  /** 跳过部分的宽度（`w:wBefore` / `w:wAfter`） */
  widthBefore?: TableWidth;
  widthAfter?: TableWidth;
}

// ── 单元格 ────────────────────────────────────────────────────────────────────

export interface CellProps {
  /** `w:tcW`：本格宽度。与 `w:tblGrid` 冲突时以谁为准见 layout 的列宽算法 */
  width?: TableWidth;
  borders?: CellBorders;
  shading?: Shading;
  /** `w:tcMar`：覆盖表级的单元格边距 */
  margins?: CellMargins;
  /** `w:vAlign`：格内内容的垂直对齐 */
  verticalAlign?: 'top' | 'center' | 'bottom';
  /** `w:noWrap`：本格不自动换行（autofit 时它还会影响列宽的算法） */
  noWrap?: boolean;
  /** `w:tcFitText`：缩放字符宽度把内容塞满本格 */
  fitText?: boolean;
  /**
   * `w:textDirection`：文字方向。竖排（`tbRl` / `btLr`）在公文的版记里偶尔出现，
   * 这里**只收不用** —— 竖排要整套另外的行盒逻辑，属于 Phase 9 之后。
   */
  textDirection?: string;
}

// ── 级联结果 ──────────────────────────────────────────────────────────────────
//
// 边框与底纹在结果里**仍然可缺席**：表格的边框冲突解析（相邻格子谁的线赢）
// 是布局层的事，把「没有边框」和「明确画一条 nil」在这里抹平，冲突解析就没得判了。

export interface ResolvedTableProps {
  styleId: string;
  width: TableWidth;
  justification: Justification;
  indent: TableWidth;
  borders: TableBorders;
  shading: Shading | undefined;
  cellMargins: Required<CellMargins>;
  cellSpacing: TableWidth;
  layout: 'fixed' | 'autofit';
  look: TableLook;
  rowBandSize: number;
  colBandSize: number;
}

export interface ResolvedRowProps {
  height: RowHeight;
  cantSplit: boolean;
  header: boolean;
  justification: Justification | undefined;
  cellSpacing: TableWidth;
  gridBefore: number;
  gridAfter: number;
  widthBefore: TableWidth;
  widthAfter: TableWidth;
  /**
   * **本行专有的表级边框**，只有 `w:tblPrEx` 改过整表边框时才有（见 nodes.ts 的 `propsEx`）。
   *
   * 边框冲突解析的第一级是「单元格没写就退到表级」，而「表级」对这一行来说
   * 是被例外改过的那一份 —— 缺了这个字段，粘进来的那一行会沿用整表的线宽。
   * 缺席表示「就用整表那一份」，不是「没有边框」。
   */
  tableBorders?: TableBorders;
}

export interface ResolvedCellProps {
  width: TableWidth;
  borders: CellBorders;
  shading: Shading | undefined;
  margins: Required<CellMargins>;
  verticalAlign: 'top' | 'center' | 'bottom';
  noWrap: boolean;
  fitText: boolean;
  textDirection: string;
}

// ── 表格样式的条件格式 ────────────────────────────────────────────────────────

/**
 * `w:tblStylePr/@w:type` 的取值。
 *
 * 别被名字骗了：`band1Horz` 是**行**带（水平的带子 = 一行行的），`band1Vert` 是列带。
 * `band1` 是奇数带（第 1、3、5…），`band2` 是偶数带。
 */
export type TableStyleOverrideType =
  | 'wholeTable'
  | 'band1Vert'
  | 'band2Vert'
  | 'band1Horz'
  | 'band2Horz'
  | 'firstCol'
  | 'lastCol'
  | 'firstRow'
  | 'lastRow'
  | 'nwCell'
  | 'neCell'
  | 'swCell'
  | 'seCell';

/**
 * 表格样式里的一份条件格式。五种属性它都能带 —— 包括 `pPr` / `rPr`：
 * 「表头行加粗」就是 `firstRow` 的 `rPr` 里一个 `w:b`，它必须能穿透到单元格里的
 * 每个 run 上去（见 cascade-table.ts 的 `cellStyleLayers`）。
 */
export interface TableStyleOverride {
  paraProps: ParaProps;
  runProps: RunProps;
  tableProps: TableProps;
  rowProps: RowProps;
  cellProps: CellProps;
}

/**
 * 一份表格样式在某个单元格位置上展开后、要铺给**格内段落**的层。
 *
 * 段落与字符属性单拎出来，是因为它们不在这里 finish —— 单元格里的段落还有自己的
 * 段落样式链要走，这两份只是排在那条链**前面**的额外层（见 cascade.ts 文件头）。
 */
export interface TableStyleLayer {
  paraProps: ParaProps;
  runProps: RunProps;
}

// ── 常量 ──────────────────────────────────────────────────────────────────────

export const AUTO_WIDTH: TableWidth = { value: 0, type: 'auto' };
export const NIL_WIDTH: TableWidth = { value: 0, type: 'nil' };

/**
 * 条件格式的**应用顺序**（ECMA-376 §17.7.6，后面的覆盖前面的）。
 *
 * 两处反直觉、也是自己实现时最容易搞错的：
 * 1. **行带在列带之后** —— 两者都命中时行带赢
 * 2. **首末行在首末列之后** —— 所以「表头行整行加粗」会盖住「首列不加粗」，
 *    这正是 Word 里表头行左上角那格跟着表头走的原因
 *
 * 角单元格排最后，它们是四个角上的最终裁决。
 */
export const CONDITIONAL_ORDER: readonly TableStyleOverrideType[] = [
  'wholeTable',
  'band1Vert',
  'band2Vert',
  'band1Horz',
  'band2Horz',
  'firstCol',
  'lastCol',
  'firstRow',
  'lastRow',
  'nwCell',
  'neCell',
  'swCell',
  'seCell',
];
