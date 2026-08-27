/**
 * 表格边框的**冲突解析**：一条格线两边各有一个单元格，各自都可能声明了边框，谁画得算？
 *
 * 这一层不需要 y，所以能在基线穿刺之前做完 —— 它算的是「哪条边用哪个 `Border`」，
 * 至于那条边画在什么高度，等行盒装配好了直接从 `RowLayout` 的 y 拿。
 *
 * ## 两级模型（顺序不能反）
 *
 * ```
 * ① 层级覆盖：单元格自己写了这条边 → 用它；没写 → 退到表级（w:tblBorders）
 * ② 相邻竞争：共享这条线的两个格子各拿一条候选出来比，赢的那条画
 * ```
 *
 * 第 ① 步里「单元格写了 `w:val="nil"`」是**明确的无边框**，它必须赢过表级的
 * `insideH` —— 这正是 Word 里「擦掉某一格的某条格线」的实现方式。反过来在第 ② 步里
 * `nil` 是弱的：一格说 nil、邻格说 single，Word 画那条 single。把两步合成一步
 * （比如「nil 一律赢」）会让整张表的内部格线全被一格的 nil 抹掉。
 *
 * ## 竞争规则已经实测（`spike-table-03`，21 组 × 横竖两遍全对）
 *
 * 规则与证据表见下面的 `BORDER_CONFLICT_RULES`。原来照 CSS 2.1 §17.6.2 类比写的
 * 「先比线宽、再比样式权重」**错了一半**：Word 是**先分类**（点线 < 虚线 < 实线类），
 * 类不同的时候线宽一点都不管用，而且**同一种破折线之间根本不比宽度**。
 *
 * ## 已知的洞
 *
 * - 单元格自己的 `insideH` / `insideV`（合并区**内部**的线）不消费。跨列合并区内部
 *   本来就没有格子边界，跨行合并区内部按「不画」处理 —— 这是 Word 的默认行为，
 *   但显式写了 `w:tcBorders/w:insideH` 想把内部线画出来的文件我们会画不出来
 * - `w:tblCellSpacing` 非 0 时格线不再共享（两个格子各画各的框），这里仍按共享算。
 *   与 table.ts 一致：间距本身就没消费
 * - 边框宽度不吃可用宽（Word 把线画在格线上，不缩文字区），所以这一层完全不改坐标
 */
import type { Border, CellBorders, ResolvedTableRow, TableBorders } from '@uw/model';
import { borderSolidRank, borderStyleClass, borderThicknessFactor } from './uncalibrated.ts';

/**
 * 一条水平边在某段列上的解析结果。
 *
 * 上下两行的格子边界不一定对齐（表头一格跨 3 列、下面 3 格），那条线就会**分段** ——
 * 每段各自跟不同的邻格竞争。垂直边没有这个问题：它只跨本行一行，两侧就是两个格子。
 */
export interface BorderSegment {
  /** 起始网格列 */
  col: number;
  /** 占几个网格列 */
  span: number;
  /** `undefined` = 这一段不画线 */
  border: Border | undefined;
}

/** 一个单元格四周（加两条对角线）解析完的边框 */
export interface CellBorderLayout {
  top: BorderSegment[];
  bottom: BorderSegment[];
  left: Border | undefined;
  right: Border | undefined;
  /**
   * 对角线不与任何人共享，直接取本格的 `w:tl2br` / `w:tr2bl`，不参与竞争。
   * 隐藏在这里的一致性：`nil` 在这里也是「不画」，与竞争里的弱语义无关
   */
  tl2br: Border | undefined;
  tr2bl: Border | undefined;
}

/** 网格上某个位置站着哪一格：行号 + 该行里第几个 `w:tc` */
interface CellRef {
  row: number;
  index: number;
}

/** 边框解析要的最小信息，和 `ResolvedTableRow` 对齐，测试里可以直接造 */
export interface BorderCell {
  gridSpan: number;
  vMerge: 'none' | 'restart' | 'continue';
  borders: CellBorders;
}

export interface BorderRow {
  gridBefore: number;
  cells: BorderCell[];
  /**
   * 本行专有的表级边框（`w:tblPrEx`，见 model 的 `ResolvedRowProps.tableBorders`）。
   * 缺席表示用整表那一份 —— 「层级覆盖」里的「表级」对这一行说的就是它。
   *
   * 共享的那条线上两侧各按**自己那一行**的表级边框出候选：例外只改了这一行，
   * 用它去替对面那一行发言就把例外的作用范围扩大了一行。
   */
  tableBorders?: TableBorders;
}

/** 从模型的行结构里取出边框解析要的那几项 */
export function borderRowsOf(rows: readonly ResolvedTableRow[]): BorderRow[] {
  return rows.map((r) => ({
    gridBefore: r.props.gridBefore,
    cells: r.cells.map((c) => ({
      gridSpan: c.gridSpan,
      vMerge: c.vMerge,
      borders: c.props.borders,
    })),
    ...(r.props.tableBorders === undefined ? {} : { tableBorders: r.props.tableBorders }),
  }));
}

/** 这一行「退到表级」时用哪一份边框：例外改过就用例外的 */
function tableAt(rows: readonly BorderRow[], r: number, table: TableBorders): TableBorders {
  return rows[r]?.tableBorders ?? table;
}

/**
 * 解析整张表的边框，返回值与 `rows[i].cells[j]` 一一对应。
 *
 * `colCount` 用列宽数组的长度（`TableLayout.columns`），不要用「最长那行的格子数」——
 * 外框判定（`col + span === colCount` 才是最右列）依赖它，差一位整列外框就变成内部线。
 */
export function resolveTableBorders(
  rows: readonly BorderRow[],
  tableBorders: TableBorders,
  colCount: number,
  rules: BorderConflictRules = BORDER_CONFLICT_RULES,
): CellBorderLayout[][] {
  const grid = buildGrid(rows, colCount);
  const cols = columnsOf(rows, colCount);

  return rows.map((row, r) =>
    row.cells.map((cell, i) => {
      const col = cols[r]?.[i] ?? 0;
      const span = cell.gridSpan;
      return {
        top: horizontalEdge(rows, grid, tableBorders, r, cell, col, span, 'top', rules),
        bottom: horizontalEdge(rows, grid, tableBorders, r, cell, col, span, 'bottom', rules),
        left: verticalEdge(rows, grid, tableBorders, r, cell, col, span, 'left', rules),
        right: verticalEdge(rows, grid, tableBorders, r, cell, col, span, 'right', rules),
        tl2br: visible(cell.borders.tl2br) ? cell.borders.tl2br : undefined,
        tr2bl: visible(cell.borders.tr2bl) ? cell.borders.tr2bl : undefined,
      };
    }),
  );
}

// ── 网格 ──────────────────────────────────────────────────────────────────────

/**
 * 铺一张 `行 × 网格列` 的表，每格记「站在这儿的是谁」。
 *
 * `w:gridBefore` / `w:gridAfter` 跳掉的位置留 `undefined`：那儿**没有格子**，
 * 与之相邻的边只有一个候选。不留空的话第一格会被当成第 0 列，整行的外框判定全错。
 */
function buildGrid(rows: readonly BorderRow[], colCount: number): (CellRef | undefined)[][] {
  return rows.map((row, r) => {
    const line = new Array<CellRef | undefined>(colCount).fill(undefined);
    let col = row.gridBefore;
    row.cells.forEach((cell, index) => {
      for (let k = 0; k < cell.gridSpan; k++) {
        if (col + k < colCount) line[col + k] = { row: r, index };
      }
      col += cell.gridSpan;
    });
    return line;
  });
}

/** 每格的起始网格列，与 `buildGrid` 用同一套推进规则，免得两处算出不同的列号 */
function columnsOf(rows: readonly BorderRow[], colCount: number): number[][] {
  return rows.map((row) => {
    let col = row.gridBefore;
    return row.cells.map((cell) => {
      const at = Math.min(col, colCount);
      col += cell.gridSpan;
      return at;
    });
  });
}

// ── 候选 ──────────────────────────────────────────────────────────────────────

/**
 * 某格某条边的候选边框：单元格显式写了就用它（含 `nil`），没写才退到表级。
 *
 * `outer` 决定退到表级时用哪一条：贴着表格外沿的用 `top` / `left` / `bottom` / `right`，
 * 内部的用 `insideH` / `insideV`。这就是「表格四周粗、内部细」那种最常见的表的成因。
 */
function candidate(
  borders: CellBorders,
  table: TableBorders,
  side: 'top' | 'bottom' | 'left' | 'right',
  outer: boolean,
): Border | undefined {
  const own = borders[side];
  if (own !== undefined) return own;
  if (outer) return table[side];
  return side === 'top' || side === 'bottom' ? table.insideH : table.insideV;
}

// ── 水平边 ────────────────────────────────────────────────────────────────────

/**
 * 水平边（`top` / `bottom`）**按列分段**解析。
 *
 * 分段的依据是对面那一行的格子边界：本格跨 3 列、对面是 3 个格子，这条线就分成
 * 3 段，各自跟不同的邻格竞争。合并成一段再比，等于让其中一格的边框替另外两格发言。
 */
function horizontalEdge(
  rows: readonly BorderRow[],
  grid: readonly (CellRef | undefined)[][],
  table: TableBorders,
  r: number,
  cell: BorderCell,
  col: number,
  span: number,
  side: 'top' | 'bottom',
  rules: BorderConflictRules,
): BorderSegment[] {
  const otherRow = side === 'top' ? r - 1 : r + 1;
  const outer = side === 'top' ? r === 0 : r === rows.length - 1;
  const mine = candidate(cell.borders, tableAt(rows, r, table), side, outer);
  const opposite = side === 'top' ? 'bottom' : 'top';

  const out: BorderSegment[] = [];
  for (let c = col; c < col + span; c++) {
    const neighbor = grid[otherRow]?.[c];
    const border = mergedAt(rows, cell, neighbor, side)
      ? undefined
      : // 上面那格是左上者：解析本格 `top` 时邻格就在上面，平局让它赢
        winner(mine, neighborCandidate(rows, table, neighbor, opposite), side === 'top', rules);
    push(out, c, border);
  }
  return out;
}

/**
 * 这条水平边是不是**跨行合并区的内部**？是的话不画。
 *
 * 判据是「下面那格是 `vMerge=continue`」：`continue` 的格子与它上方的格子属于同一个
 * 合并区，中间那条线在 Word 里默认不存在。畸形文件（`continue` 上面站着 `none`）
 * 也按合并处理 —— 既然它声称自己是续格，那条线就不该冒出来。
 */
function mergedAt(
  rows: readonly BorderRow[],
  cell: BorderCell,
  neighbor: CellRef | undefined,
  side: 'top' | 'bottom',
): boolean {
  if (neighbor === undefined) return false;
  const lower = side === 'top' ? cell : cellAt(rows, neighbor);
  return lower?.vMerge === 'continue';
}

// ── 垂直边 ────────────────────────────────────────────────────────────────────

/** 垂直边只跨本行，两侧最多两个格子，不分段 */
function verticalEdge(
  rows: readonly BorderRow[],
  grid: readonly (CellRef | undefined)[][],
  table: TableBorders,
  r: number,
  cell: BorderCell,
  col: number,
  span: number,
  side: 'left' | 'right',
  rules: BorderConflictRules,
): Border | undefined {
  const colCount = grid[r]?.length ?? 0;
  const outer = side === 'left' ? col === 0 : col + span >= colCount;
  const mine = candidate(cell.borders, tableAt(rows, r, table), side, outer);
  const neighbor = grid[r]?.[side === 'left' ? col - 1 : col + span];
  const opposite = side === 'left' ? 'right' : 'left';
  // 左边那格是左上者：解析本格 `left` 时邻格就在左边，平局让它赢
  return winner(mine, neighborCandidate(rows, table, neighbor, opposite), side === 'left', rules);
}

function neighborCandidate(
  rows: readonly BorderRow[],
  table: TableBorders,
  ref: CellRef | undefined,
  side: 'top' | 'bottom' | 'left' | 'right',
): Border | undefined {
  const cell = cellAt(rows, ref);
  if (cell === undefined || ref === undefined) return undefined;
  // 邻格存在就说明这条线不在表格外沿上，一律走 inside*
  return candidate(cell.borders, tableAt(rows, ref.row, table), side, false);
}

function cellAt(rows: readonly BorderRow[], ref: CellRef | undefined): BorderCell | undefined {
  if (ref === undefined) return undefined;
  return rows[ref.row]?.cells[ref.index];
}

// ── 竞争 ──────────────────────────────────────────────────────────────────────

/**
 * `nil` / `none` 都是「没有边框」，只是来路不同（明确无 vs 未指定）。
 * 到了画不画这一步两者同义，所以这里合并 —— 区分它俩的是**层级覆盖**那一步，
 * 在 `candidate()` 里靠 `own !== undefined` 完成，与本函数无关。
 */
function visible(b: Border | undefined): b is Border {
  return b !== undefined && b.style !== 'nil' && b.style !== 'none';
}

/**
 * ## 相邻竞争的规则（`spike-table-03` 实测，21 组 × 横竖两遍，两个方向结论一致）
 *
 * 样本做法：给竞争的两侧各一个独一无二的颜色（左 / 上 = 红，右 / 下 = 蓝），
 * PDF 里画出来的线是什么颜色就直接说出赢家是谁 —— 不必从线宽反推（好几组同宽不同样式，
 * 线宽在那几组分不开）。冲突的 `w:tcBorders` 是**改 XML** 造出来的：Word 的对象模型里
 * 一条共享边只有一个 Border 对象，设一侧等于两侧都设，它自己造不出这个局面。
 *
 * | 组 | 左 / 上 | 右 / 下 | 画出来的 | 说明 |
 * |---|---|---|---|---|
 * | 〇 | —（不写） | —（不写） | 表级绿线 | 层级覆盖那一级，也是读数方法的对照 |
 * | 一 | single 0.5 | single 2.25 | 蓝 2.16 | 同类比厚度 |
 * | 二 | single 2.25 | single 0.5 | 红 2.16 | 一的镜像，排除「总是某一侧赢」 |
 * | 三 | single 1.5 | double 1.5 | 蓝双线 | double 画出来 4.32 = 3 × sz |
 * | 四 | double 1.5 | single 1.5 | 红双线 | 三的镜像 |
 * | 五 | single 1.5 红 | single 1.5 蓝 | **红** | 全平局取**左上** |
 * | 六 | nil | single 1.5 | 蓝 | `nil` 在竞争里是弱的 |
 * | 七 | single 1.5 | nil | 红 | 六的镜像 |
 * | 八 | dotted 3.0 | single 0.75 | 蓝 0.72 | **破折类再宽也输给实线** |
 * | 九 | double 0.75 | single 3.0 | 蓝 3.0 | 双线 2.16 厚 < 单线 3.0 厚 |
 * | 十 | dashed 1.5 | dotted 1.5 | 红虚线 | dashed > dotted |
 * | 甲 | nil | nil | **无线** | `nil` 赢过表级的绿线（层级覆盖那一级） |
 * | 乙 | dotted 1.5 | dashed 1.5 | 蓝虚线 | 十的镜像，dashed > dotted 是真的排序 |
 * | 丙 | dashed 3.0 | single 0.5 | 蓝 0.48 | 八换成 dashed，同样输 |
 * | 丁 | single 3.0 | double 1.5 | 蓝双线 | 4.32 > 3.0，钉死双线算 **3 倍**不是 2 倍 |
 * | 戊 | dotted 0.5 | dotted 2.25 | **红 0.48** | 同一种破折线之间**不比宽度** |
 * | 己 | double 1.5 | double 0.75 | 红双线 | 实线类内部照样比厚度 |
 * | 庚 | dotted 2.25 | nil | 红 2.16 虚 | 戊 / 壬 的对照：宽点线本身画得出来 |
 * | 辛 | dashed 0.5 | dashed 3.0 | 红 0.48 | 虚线之间也不比宽度 |
 * | 壬 | dotted 2.25 | dotted 0.5 | 红 2.16 | 戊的镜像，两边都是左上赢 |
 * | 癸 | single 1.5 | double 0.5 | **蓝双线** | 厚度都是 1.44 → 样式权重再比一次 |
 *
 * 被推翻的旧写法（照 CSS 2.1 §17.6.2 类比来的）是「先比线宽、再比样式权重」：
 * 八 / 丙 说它错 —— 3pt 的点线 / 虚线输给 0.5pt 的单线，跨类时线宽完全不算数；
 * 戊 / 辛 / 壬 说它还错了一处 —— 同一种破折线之间连宽度都不比，直接看位置。
 */
export interface BorderConflictRules {
  /**
   * 先按线型分类（点线 < 虚线 < 实线类）再比，还是不分类只比厚度。
   * **实测 `class`**：八 / 丙 里 3pt 的破折线输给 0.5pt 的实线。
   */
  order: 'class' | 'thickness';
  /** 同一档破折线之间比不比厚度。**实测 `false`**（戊 / 辛 / 壬） */
  brokenByThickness: boolean;
  /**
   * 厚度按什么算：`rendered` = `w:sz` × 线型倍数（双线 3 倍），`size` = 直接用 `w:sz`。
   * **实测 `rendered`**：丁组里 1.5pt 的双线（画出来 4.32pt）赢过 3pt 的单线。
   */
  thickness: 'rendered' | 'size';
  /** 厚度打平时再比一次样式权重（多线型 > 单线型）。**实测 `true`**（癸） */
  styleBreaksTie: boolean;
  /** 全平局取哪一侧。**实测 `leftTop`**（五，横竖两个方向一致） */
  position: 'leftTop' | 'rightBottom';
}

export const BORDER_CONFLICT_RULES: BorderConflictRules = {
  order: 'class',
  brokenByThickness: false,
  thickness: 'rendered',
  styleBreaksTie: true,
  position: 'leftTop',
};

/** 一条边按规则折算出来的厚度 */
function thicknessOf(b: Border, rules: BorderConflictRules): number {
  return rules.thickness === 'size' ? b.size : b.size * borderThicknessFactor(b.style);
}

/** `a` 比 `b` 强返回正数，弱返回负数，分不出高下返回 0 */
function compare(a: Border, b: Border, rules: BorderConflictRules): number {
  if (rules.order === 'class') {
    const ka = borderStyleClass(a.style);
    const kb = borderStyleClass(b.style);
    if (ka !== kb) return ka - kb;
    // 同一档破折线：Word 连宽度都不比（戊 / 辛 / 壬），直接交给位置
    if (ka < 2 && !rules.brokenByThickness) return 0;
  }
  const ta = thicknessOf(a, rules);
  const tb = thicknessOf(b, rules);
  if (ta !== tb) return ta > tb ? 1 : -1;
  if (rules.styleBreaksTie) {
    const ra = borderSolidRank(a.style);
    const rb = borderSolidRank(b.style);
    if (ra !== rb) return ra > rb ? 1 : -1;
  }
  return 0;
}

/**
 * 两个候选谁画。`neighborIsLeftTop` 为真时 `other` 是左上那一侧。
 *
 * 顺序：可见性 → `compare()` → 位置。前两条分不出高下时两条线也不一定长得一样
 * （颜色可以不同），位置那一条既是实测结论（第五组）也让结果**确定**，方便逐条比对。
 */
function winner(
  mine: Border | undefined,
  other: Border | undefined,
  neighborIsLeftTop: boolean,
  rules: BorderConflictRules,
): Border | undefined {
  const a = visible(mine) ? mine : undefined;
  const b = visible(other) ? other : undefined;
  if (a === undefined) return b;
  if (b === undefined) return a;

  const cmp = compare(a, b, rules);
  if (cmp !== 0) return cmp > 0 ? a : b;
  const otherWins = rules.position === 'leftTop' ? neighborIsLeftTop : !neighborIsLeftTop;
  return otherWins ? b : a;
}

// ── 分段合并 ──────────────────────────────────────────────────────────────────

/** 逐列算完再合并相邻的同结果段，省得渲染层拿到一堆一列宽的碎片 */
function push(out: BorderSegment[], col: number, border: Border | undefined): void {
  const last = out[out.length - 1];
  if (last !== undefined && last.col + last.span === col && sameBorder(last.border, border)) {
    last.span += 1;
    return;
  }
  out.push({ col, span: 1, border });
}

function sameBorder(a: Border | undefined, b: Border | undefined): boolean {
  if (a === undefined || b === undefined) return a === b;
  return a.style === b.style && a.size === b.size && a.color === b.color && a.space === b.space;
}
