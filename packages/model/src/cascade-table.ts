/**
 * 表格的样式级联（Phase 4）—— cascade.ts 的表格版。
 *
 * 层序（后面的覆盖前面的），表 / 行 / 格 / 格内段落**共用同一条**：
 *
 * ```
 * 1. 表格样式链上每一级的自身属性（w:style/w:tblPr、w:trPr、w:tcPr、w:pPr、w:rPr）
 * 2. 命中的条件格式（w:tblStylePr），按 CONDITIONAL_ORDER；同一类型内仍按样式链序
 * 3. 直接格式（w:tbl/w:tblPr、w:tr/w:trPr、w:tc/w:tcPr）
 * ```
 *
 * 三处容易错：
 * 1. **没写 `w:tblStyle` 的表格照样吃默认表格样式**（`Normal Table`）——
 *    单元格左右各 108 twips 的默认边距就在那份样式里，不在什么规范常数里
 * 2. **`w:tblLook` 是开关**：样式里定义了 `firstRow` 的格式，但 look 说不要，
 *    那份格式就不应用。漏了它，凡是用了内置表格样式的表都会平白多出加粗表头
 * 3. **表格样式的 pPr / rPr 排在段落样式链之前**（§17.7.2）——「表头行加粗」
 *    要能被单元格里段落自己的样式盖掉，反过来就成了段落样式永远赢
 *
 * **未经真值验证的一处**：隔行带（`band1Horz` / `band2Horz` / `band1Vert` / `band2Vert`）
 * 的序号算法 —— 「首行是否计入带」「带从 0 还是 1 开始数」照规范实现，没有 Word 样本。
 * 它影响的是底纹和字重（字重会改宽度 → 改断行），所以上 Windows 时值得补一份
 * 「4 行 3 列、开隔行带、`rowBandSize=2`」的样本钉死。
 */
import type { CascadeContext } from './cascade.ts';
import type { Style } from './styles.ts';
import type {
  CellMargins,
  CellProps,
  ResolvedCellProps,
  ResolvedRowProps,
  ResolvedTableProps,
  RowProps,
  TableLook,
  TableProps,
  TableStyleLayer,
  TableStyleOverrideType,
} from './table-props.ts';
import { AUTO_WIDTH, CONDITIONAL_ORDER, NIL_WIDTH } from './table-props.ts';

/** 行在表里的位置。行级属性只问得到行的条件（首行 / 末行 / 行带） */
export interface RowPosition {
  row: number;
  rowCount: number;
}

/**
 * 单元格在表里的位置 —— 条件格式命中与否全靠它。
 *
 * `col` 是**网格列号**（把前面所有格子的 `gridSpan` 加起来），不是「第几个 `w:tc`」：
 * 一个跨 3 列的格子之后，下一个格子的列号是 3 不是 1。用错的话「末列加粗」
 * 会加在中间某一格上。
 *
 * 与 `RowPosition` 分成两个类型而不是「列那几项填 0」：`col: 0` 的含义是
 * **首列**，拿它当「没有列」用，`firstCol` 会在每一行上无条件命中。
 */
export interface CellPosition extends RowPosition {
  /** 本格起始网格列 */
  col: number;
  /** 本格占的网格列数（`gridSpan`） */
  span: number;
  colCount: number;
}

/** Word 默认模板里 `Normal Table` 的单元格边距。样式表缺失时的兜底，正常文件会盖掉它 */
const FALLBACK_CELL_MARGINS: Required<CellMargins> = {
  top: { value: 0, type: 'dxa' },
  left: { value: 108, type: 'dxa' },
  bottom: { value: 0, type: 'dxa' },
  right: { value: 108, type: 'dxa' },
};

const NO_LOOK: TableLook = {
  firstRow: false,
  lastRow: false,
  firstColumn: false,
  lastColumn: false,
  noHBand: false,
  noVBand: false,
};

/** 没写 `w:tblStyle` 时用标了 `w:default="1"` 的那个表格样式 —— 与段落吃 Normal 同构 */
export function tableStyleChain(ctx: CascadeContext, direct: TableProps | undefined): Style[] {
  const id = direct?.styleId ?? ctx.styles.defaultTableStyleId();
  return ctx.styles.chainOf(id).filter((s) => s.type === 'table');
}

// ── 表级 ──────────────────────────────────────────────────────────────────────

function applyTableLevel(acc: TableProps, level: TableProps): void {
  if (level.width !== undefined) acc.width = level.width;
  if (level.justification !== undefined) acc.justification = level.justification;
  if (level.indent !== undefined) acc.indent = level.indent;
  if (level.shading !== undefined) acc.shading = level.shading;
  if (level.cellSpacing !== undefined) acc.cellSpacing = level.cellSpacing;
  if (level.layout !== undefined) acc.layout = level.layout;
  if (level.look !== undefined) acc.look = level.look;
  if (level.rowBandSize !== undefined) acc.rowBandSize = level.rowBandSize;
  if (level.colBandSize !== undefined) acc.colBandSize = level.colBandSize;
  // 边框与边距逐边合并：样式定了四周、直接格式只改了 insideH，另外五条要留着
  if (level.borders !== undefined) acc.borders = { ...acc.borders, ...definedOnly(level.borders) };
  if (level.cellMargins !== undefined) {
    acc.cellMargins = { ...acc.cellMargins, ...definedOnly(level.cellMargins) };
  }
}

/**
 * 表级属性。
 *
 * 条件格式里只有 `wholeTable` 对整表有意义（其余那些是行 / 列 / 角上的），
 * 所以这里只展开它 —— 拿 `firstRow` 的 `tblPr` 去改整表宽度是没有意义的。
 */
export function resolveTableProps(ctx: CascadeContext, direct: TableProps | undefined): ResolvedTableProps {
  const chain = tableStyleChain(ctx, direct);
  const acc: TableProps = {};

  for (const s of chain) applyTableLevel(acc, s.tableProps);
  for (const s of chain) {
    const whole = s.conditional.get('wholeTable');
    if (whole !== undefined) applyTableLevel(acc, whole.tableProps);
  }
  if (direct !== undefined) applyTableLevel(acc, direct);

  return {
    styleId: direct?.styleId ?? ctx.styles.defaultTableStyleId(),
    width: acc.width ?? AUTO_WIDTH,
    justification: acc.justification ?? 'left',
    indent: acc.indent ?? { value: 0, type: 'dxa' },
    borders: acc.borders ?? {},
    shading: acc.shading,
    cellMargins: { ...FALLBACK_CELL_MARGINS, ...definedOnly(acc.cellMargins ?? {}) },
    cellSpacing: acc.cellSpacing ?? NIL_WIDTH,
    // 规范默认是 autofit，不是 fixed —— 公文里那些「固定列宽」的表几乎都显式写了 fixed
    layout: acc.layout ?? 'autofit',
    look: acc.look ?? NO_LOOK,
    // 一条带默认一行 / 一列
    rowBandSize: acc.rowBandSize ?? 1,
    colBandSize: acc.colBandSize ?? 1,
  };
}

// ── 条件格式的命中 ────────────────────────────────────────────────────────────

/**
 * 这个位置命中了哪些条件格式，**已按应用顺序排好**。
 *
 * 带（band）的判定要先把首末行 / 首末列排除掉：Word 里表头行不算进隔行带，
 * 否则「表头 + 隔行底纹」的表会从第二行开始错位一整行。
 */
export function conditionsAt(
  look: TableLook,
  bands: { row: number; col: number },
  pos: RowPosition | CellPosition,
): TableStyleOverrideType[] {
  const hit = new Set<TableStyleOverrideType>();
  // 只给了行的位置时，列上的那几种条件一个都不该命中 —— 行属性问的就是「整行」
  const cell = 'col' in pos ? pos : undefined;

  const isFirstRow = look.firstRow && pos.row === 0;
  const isLastRow = look.lastRow && pos.row === pos.rowCount - 1;
  const isFirstCol = cell !== undefined && look.firstColumn && cell.col === 0;
  // 跨列的格子只要**盖到**最后一列就算末列
  const isLastCol = cell !== undefined && look.lastColumn && cell.col + cell.span >= cell.colCount;

  if (!look.noHBand && !isFirstRow && !isLastRow) {
    hit.add(bandOf(pos.row - (look.firstRow ? 1 : 0), bands.row, 'band1Horz', 'band2Horz'));
  }
  if (cell !== undefined && !look.noVBand && !isFirstCol && !isLastCol) {
    hit.add(bandOf(cell.col - (look.firstColumn ? 1 : 0), bands.col, 'band1Vert', 'band2Vert'));
  }
  if (isFirstCol) hit.add('firstCol');
  if (isLastCol) hit.add('lastCol');
  if (isFirstRow) hit.add('firstRow');
  if (isLastRow) hit.add('lastRow');
  // 四个角：只有行与列的条件同时成立才算，它们是最后的裁决者
  if (isFirstRow && isFirstCol) hit.add('nwCell');
  if (isFirstRow && isLastCol) hit.add('neCell');
  if (isLastRow && isFirstCol) hit.add('swCell');
  if (isLastRow && isLastCol) hit.add('seCell');

  return CONDITIONAL_ORDER.filter((t) => hit.has(t));
}

/** 第几条带 → band1（奇数条）还是 band2（偶数条）。`size` 是一条带占几行 / 几列 */
function bandOf<T>(index: number, size: number, band1: T, band2: T): T {
  if (index < 0) return band1;
  const n = Math.floor(index / Math.max(1, size));
  return n % 2 === 0 ? band1 : band2;
}

// ── 行级 ──────────────────────────────────────────────────────────────────────

function applyRowLevel(acc: RowProps, level: RowProps): void {
  if (level.height !== undefined) acc.height = level.height;
  if (level.cantSplit !== undefined) acc.cantSplit = level.cantSplit;
  if (level.header !== undefined) acc.header = level.header;
  if (level.justification !== undefined) acc.justification = level.justification;
  if (level.cellSpacing !== undefined) acc.cellSpacing = level.cellSpacing;
  if (level.gridBefore !== undefined) acc.gridBefore = level.gridBefore;
  if (level.gridAfter !== undefined) acc.gridAfter = level.gridAfter;
  if (level.widthBefore !== undefined) acc.widthBefore = level.widthBefore;
  if (level.widthAfter !== undefined) acc.widthAfter = level.widthAfter;
}

export function resolveRowProps(
  ctx: CascadeContext,
  table: ResolvedTableProps,
  tableDirect: TableProps | undefined,
  direct: RowProps | undefined,
  pos: RowPosition,
): ResolvedRowProps {
  const chain = tableStyleChain(ctx, tableDirect);
  const types = conditionsAt(table.look, { row: table.rowBandSize, col: table.colBandSize }, pos);

  const acc: RowProps = {};
  for (const s of chain) applyRowLevel(acc, s.rowProps);
  for (const type of types) {
    for (const s of chain) {
      const o = s.conditional.get(type);
      if (o !== undefined) applyRowLevel(acc, o.rowProps);
    }
  }
  if (direct !== undefined) applyRowLevel(acc, direct);

  return {
    // 行高缺席 = 完全按内容撑开
    height: acc.height ?? { value: 0, rule: 'auto' },
    cantSplit: acc.cantSplit ?? false,
    header: acc.header ?? false,
    // 缺席表示跟随整表，不是「左对齐」—— 所以这里保留 undefined
    justification: acc.justification,
    cellSpacing: acc.cellSpacing ?? table.cellSpacing,
    gridBefore: acc.gridBefore ?? 0,
    gridAfter: acc.gridAfter ?? 0,
    widthBefore: acc.widthBefore ?? NIL_WIDTH,
    widthAfter: acc.widthAfter ?? NIL_WIDTH,
  };
}

// ── 格级 ──────────────────────────────────────────────────────────────────────

function applyCellLevel(acc: CellProps, level: CellProps): void {
  if (level.width !== undefined) acc.width = level.width;
  if (level.shading !== undefined) acc.shading = level.shading;
  if (level.verticalAlign !== undefined) acc.verticalAlign = level.verticalAlign;
  if (level.noWrap !== undefined) acc.noWrap = level.noWrap;
  if (level.fitText !== undefined) acc.fitText = level.fitText;
  if (level.textDirection !== undefined) acc.textDirection = level.textDirection;
  if (level.borders !== undefined) acc.borders = { ...acc.borders, ...definedOnly(level.borders) };
  if (level.margins !== undefined) acc.margins = { ...acc.margins, ...definedOnly(level.margins) };
}

/**
 * 单元格属性 + 这个位置上要铺给格内段落的样式层，一次算完。
 *
 * 两者共用同一串条件格式，分两次算等于把「命中哪些条件」这件事写两遍 ——
 * 而这正是最容易两边不一致的地方。
 */
export function resolveCellProps(
  ctx: CascadeContext,
  table: ResolvedTableProps,
  tableDirect: TableProps | undefined,
  direct: CellProps | undefined,
  pos: CellPosition,
): { props: ResolvedCellProps; layers: TableStyleLayer[] } {
  const chain = tableStyleChain(ctx, tableDirect);
  const types = conditionsAt(table.look, { row: table.rowBandSize, col: table.colBandSize }, pos);

  const acc: CellProps = {};
  const layers: TableStyleLayer[] = [];
  for (const s of chain) {
    applyCellLevel(acc, s.cellProps);
    layers.push({ paraProps: s.paraProps, runProps: s.runProps });
  }
  for (const type of types) {
    for (const s of chain) {
      const o = s.conditional.get(type);
      if (o === undefined) continue;
      applyCellLevel(acc, o.cellProps);
      layers.push({ paraProps: o.paraProps, runProps: o.runProps });
    }
  }
  if (direct !== undefined) applyCellLevel(acc, direct);

  return {
    props: {
      width: acc.width ?? AUTO_WIDTH,
      borders: acc.borders ?? {},
      shading: acc.shading,
      // w:tcMar 缺席时退到表级 w:tblCellMar，**不是**退到 0
      margins: { ...table.cellMargins, ...definedOnly(acc.margins ?? {}) },
      verticalAlign: acc.verticalAlign ?? 'top',
      noWrap: acc.noWrap ?? false,
      fitText: acc.fitText ?? false,
      // 空串 = lrTb（普通横排）。竖排只收不用，见 table-props.ts
      textDirection: acc.textDirection ?? '',
    },
    layers,
  };
}

/** `Object.assign` 会把显式的 undefined 也拷过去，先滤掉（同 cascade.ts） */
function definedOnly<T extends object>(o: T): Partial<T> {
  const out: Partial<T> = {};
  for (const [k, v] of Object.entries(o)) {
    if (v !== undefined) out[k as keyof T] = v as T[keyof T];
  }
  return out;
}

/**
 * 网格列数 —— `firstCol` / `lastCol` 要靠它判断「是不是最后一列」。
 *
 * `w:tblGrid` 在时直接数它；不在时退到「各行 `gridSpan` 之和的最大值」。
 * 取最大值而不是第一行：带 `gridBefore` 的行本来就比别人短，拿它当基准会让
 * 整表的末列判定全错一位。
 */
export function gridColumnCount(
  grid: readonly number[],
  rows: readonly { cells: readonly { gridSpan: number }[] }[],
): number {
  if (grid.length > 0) return grid.length;
  let max = 0;
  for (const r of rows) {
    let n = 0;
    for (const c of r.cells) n += c.gridSpan;
    if (n > max) max = n;
  }
  return max;
}
