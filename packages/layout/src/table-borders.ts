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
 * ## 未经真值验证的三处（原则 1.5：写明白，别装作测过）
 *
 * 1. **竞争规则本身**照 CSS 2.1 §17.6.2 的 collapsing borders 类比：先比线宽，
 *    再比样式权重（见 `uncalibrated.ts` 的 `BORDER_STYLE_RANK`），仍平局取**左上者**。
 *    ECMA-376 只规定了 `w:tcBorders` 覆盖 `w:tblBorders`，相邻两格的事一个字没提
 * 2. **`nil` / `none` 在竞争里输给一切**（= CSS 的 `none` 而非 `hidden`）
 * 3. **平局取左上**：水平边取上面那格的 `bottom`，垂直边取左边那格的 `right`
 *
 * 三条用同一份样本就能钉死：一张 2×2 的表，四条内部边分别让相邻两格写不同的
 * `w:val` / `w:sz` / `nil`，导出 PDF 看画出来的是哪一条。
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
import { borderStyleRank } from './uncalibrated.ts';

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
  }));
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
): CellBorderLayout[][] {
  const grid = buildGrid(rows, colCount);
  const cols = columnsOf(rows, colCount);

  return rows.map((row, r) =>
    row.cells.map((cell, i) => {
      const col = cols[r]?.[i] ?? 0;
      const span = cell.gridSpan;
      return {
        top: horizontalEdge(rows, grid, tableBorders, r, cell, col, span, 'top'),
        bottom: horizontalEdge(rows, grid, tableBorders, r, cell, col, span, 'bottom'),
        left: verticalEdge(rows, grid, tableBorders, r, cell, col, span, 'left'),
        right: verticalEdge(rows, grid, tableBorders, r, cell, col, span, 'right'),
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
): BorderSegment[] {
  const otherRow = side === 'top' ? r - 1 : r + 1;
  const outer = side === 'top' ? r === 0 : r === rows.length - 1;
  const mine = candidate(cell.borders, table, side, outer);
  const opposite = side === 'top' ? 'bottom' : 'top';

  const out: BorderSegment[] = [];
  for (let c = col; c < col + span; c++) {
    const neighbor = grid[otherRow]?.[c];
    const border = mergedAt(rows, cell, neighbor, side)
      ? undefined
      : // 上面那格是左上者：解析本格 `top` 时邻格就在上面，平局让它赢
        winner(mine, neighborCandidate(rows, table, neighbor, opposite), side === 'top');
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
): Border | undefined {
  const colCount = grid[r]?.length ?? 0;
  const outer = side === 'left' ? col === 0 : col + span >= colCount;
  const mine = candidate(cell.borders, table, side, outer);
  const neighbor = grid[r]?.[side === 'left' ? col - 1 : col + span];
  const opposite = side === 'left' ? 'right' : 'left';
  // 左边那格是左上者：解析本格 `left` 时邻格就在左边，平局让它赢
  return winner(mine, neighborCandidate(rows, table, neighbor, opposite), side === 'left');
}

function neighborCandidate(
  rows: readonly BorderRow[],
  table: TableBorders,
  ref: CellRef | undefined,
  side: 'top' | 'bottom' | 'left' | 'right',
): Border | undefined {
  const cell = cellAt(rows, ref);
  if (cell === undefined) return undefined;
  // 邻格存在就说明这条线不在表格外沿上，一律走 inside*
  return candidate(cell.borders, table, side, false);
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
 * 两个候选谁画。`preferOther` 为真时平局让 `other` 赢（左上者优先，见文件头第 3 条）。
 *
 * 顺序：可见性 → 线宽 → 样式权重 → 位置。前三条都比不出来的时候两条线长得一模一样，
 * 取谁都画出同一条 —— 位置那一条只是为了让结果**确定**，方便测试比对。
 */
function winner(
  mine: Border | undefined,
  other: Border | undefined,
  preferOther: boolean,
): Border | undefined {
  const a = visible(mine) ? mine : undefined;
  const b = visible(other) ? other : undefined;
  if (a === undefined) return b;
  if (b === undefined) return a;

  if (a.size !== b.size) return a.size > b.size ? a : b;
  const ra = borderStyleRank(a.style);
  const rb = borderStyleRank(b.style);
  if (ra !== rb) return ra > rb ? a : b;
  return preferOther ? b : a;
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
