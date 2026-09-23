/**
 * 表格结构编辑：插一行 / 删几行。纯函数，只吃直接格式那棵树，由 `text-transaction.ts` 包进事务。
 *
 * 三处容易搞反：
 * ① 新行**照着所在那一行抄结构**（格数、`gridSpan`、`tcPr`、`trPr`、`tblPrEx`），每格一个空段落，
 *    段落格式抄模板格的首段 —— Word「在上方 / 下方插入行」就是这样，格宽与边框不用重新算；
 *    `w:tblHeader` 也跟着抄（Word 抄整个 trPr；表头行下面插一行，它也成了表头 —— 按界面行为写，未对过真值）。
 * ② **纵向合并按网格列对齐**，不按格的下标：一行里有跨列格时，第 i 格与下一行第 i 格不在同一列。
 *    插在合并区中间（模板格是 continue，或者下面一行同一列还是 continue）的新格写 `continue`，
 *    合并区因此长一格；插在合并区外面写 `none`。
 * ③ 删掉合并区的**首格**所在行时，下面那一格要从 `continue` 升成 `restart`，否则它会去找一个
 *    已经不存在的上格，合并区整个消失（`vMerge=continue` 的格布局层不画内容，也不画与上格之间的线）。
 *    删光所有行就删整张表；表所在的容器因此空了（单元格 / 节里只剩它）时补一个空段落 ——
 *    单元格里至少要有一个段落，Word 打开会报损坏。
 *
 * 列比行多三处：
 * ④ 列是**网格列**不是格的下标 —— 插在格 [c0, c1) 的左边就是在网格边界 c0 插一条新列、右边是 c1。
 *    每一行各自看这条边界：有格**跨过**它（起点 < 边界 < 终点）就让那格 `gridSpan` + 1（合并区被撑宽，
 *    不拆开）；落在 `w:gridBefore` / `w:gridAfter` 跳过的那段里就让跳过的列数 + 1；
 *    否则边界正好在两格之间，插一个新格，照**目标格那一侧**的邻居抄（左插抄边界右边那格，右插抄左边那格），
 *    `vMerge` 也照抄 —— 挨着一个纵向合并区插列，新列在同样的几行里也合并着（按界面行为写，未对过真值）。
 * ⑤ 新列的宽 = 目标那一列的网格宽，**整表变宽**，不从别的列里匀 —— 固定列宽的表 Word 也是往右长出去
 *    （长过版心也照长，按界面行为写，未对过真值）。于是 `w:tblGrid` 插一个数，被撑宽的格与 `dxa` 型的
 *    `w:tblW` 各加上这一段；`pct` / `auto` 不动 —— 那是相对量，列宽到布局时才换算。
 * ⑥ 删列按网格列删：两端格的网格区间取并集，整个落在里面的格删掉，部分落在里面的格缩 `gridSpan`
 *    （`dxa` 的 `w:tcW` 跟着减）。一行删到一格不剩就删掉整行（没有格的 `w:tr` Word 报损坏），
 *    删掉的格可能是合并区的首格，于是删完再把「上面没有同列首格的续格」升成 restart（同 ③）。
 *    网格删光或行删光就删整张表。
 *
 * 合并 / 拆分再两处：
 * ⑦ 合并的区域先**撑成矩形**：两端格的网格区间取并集、行取两端之间，再反复往外扩，直到没有格
 *    只露出一半（跨列格横着撑、纵向合并区竖着撑）—— Word 的选区也是这样被合并格撑大的。
 *    内容**按行、行内从左到右**接起来，空格（只有空段落）不贡献段落，全空就留首格原来的段落；
 *    段落是**搬过去**的（id 不变），所以落在它们里面的位置不用映射，只有被丢掉的空段落要收拢。
 *    多行时首行那一格是 `restart`，下面每行留一个跨同样列数的 `continue` 格（一个空段落）——
 *    Word 自己就这么存，不删行：别的列在这些行里还有自己的格。
 * ⑧ 拆分只做**合并的逆操作**：跨列格拆回一列一格、纵向合并区的每一格变回 `none`，内容留在首格。
 *    Word「拆分单元格」对话框能拆成任意行列数，那要新增网格列 / 行，不在这一步。
 */
import type { Twips } from '@uw/core';
import type { Block, Body, NodeId, Paragraph, Table, TableCell, TableRow } from './nodes.ts';
import type { TableWidth } from './table-props.ts';

/** 一张表在树上的位置：它所在的块列表 + 下标，以及沿途的祖先（重建路径用） */
interface TableHit {
  table: Table;
  rowIndex: number;
  /** 位置所在的格在这一行里的下标 */
  cellIndex: number;
  /** 把新表（或 undefined = 删掉）放回原处，返回新 body */
  replace(next: Table | undefined, filler: () => Paragraph): Body;
  /** 删表之后光标该去的块：容器里紧跟着它的，没有就紧挨着它前面的 */
  neighbour(): { after: Block | undefined; before: Block | undefined };
}

/** 含 `nodeId`（段落或 run）的**最内层**那张表与行。不在表里答 undefined */
export function findRow(body: Body, nodeId: NodeId): TableHit | undefined {
  type Rebuild = (blocks: Block[]) => Body;
  function contains(p: Paragraph): boolean {
    return p.id === nodeId || p.runs.some((r) => r.id === nodeId);
  }
  function visit(blocks: Block[], rebuild: Rebuild): TableHit | undefined {
    for (const [bi, block] of blocks.entries()) {
      if (block.kind === 'paragraph') continue;
      for (const [ri, row] of block.rows.entries())
        for (const [ci, cell] of row.cells.entries()) {
          const inner = visit(cell.blocks, (next) => {
            const cells = row.cells.map((c, i) => (i === ci ? { ...cell, blocks: next } : c));
            const rows = block.rows.map((r, i) => (i === ri ? { ...row, cells } : r));
            return rebuild(blocks.map((b, i) => (i === bi ? { ...block, rows } : b)));
          });
          if (inner) return inner;
          if (!cell.blocks.some((b) => b.kind === 'paragraph' && contains(b))) continue;
          return {
            table: block,
            rowIndex: ri,
            cellIndex: ci,
            replace(next, filler) {
              const out = blocks.flatMap((b, i) => (i !== bi ? [b] : next === undefined ? [] : [next]));
              return rebuild(out.length === 0 ? [filler()] : out);
            },
            neighbour: () => ({ after: blocks[bi + 1], before: blocks[bi - 1] }),
          };
        }
    }
    return undefined;
  }
  for (const [si, section] of body.sections.entries()) {
    const hit = visit(section.blocks, (next) => ({
      ...body,
      sections: body.sections.map((s, i) => (i === si ? { ...section, blocks: next } : s)),
    }));
    if (hit) return hit;
  }
  return undefined;
}

/** 第 i 格从第几条网格列开始（`w:gridBefore` 跳过的列也算） */
function columnOf(row: TableRow, index: number): number {
  let col = row.props.gridBefore ?? 0;
  for (let i = 0; i < index; i++) col += row.cells[i]?.gridSpan ?? 1;
  return col;
}

/** 这一行里从网格列 `col` 开始的那一格；没有格恰好从这一列开始答 undefined */
function cellAtColumn(row: TableRow | undefined, col: number): TableCell | undefined {
  if (row === undefined) return undefined;
  for (let i = 0, c = row.props.gridBefore ?? 0; i < row.cells.length; i++) {
    if (c === col) return row.cells[i];
    c += row.cells[i]?.gridSpan ?? 1;
  }
  return undefined;
}

/** 照着第 `rowIndex` 行在它上 / 下方造一行（见文件头 ①②），返回新表与新行 */
export function withInsertedRow(
  table: Table,
  rowIndex: number,
  side: 'above' | 'below',
  newId: () => NodeId,
): { table: Table; row: TableRow } {
  const template = table.rows[rowIndex];
  if (template === undefined) throw new RangeError(`表格没有第 ${rowIndex} 行`);
  const below = table.rows[rowIndex + 1];
  const cells = template.cells.map((cell, i): TableCell => {
    let vMerge: TableCell['vMerge'] = 'none';
    if (side === 'above') {
      if (cell.vMerge === 'continue') vMerge = 'continue';
    } else if (cell.vMerge !== 'none' && cellAtColumn(below, columnOf(template, i))?.vMerge === 'continue')
      vMerge = 'continue';
    const first = cell.blocks.find((b): b is Paragraph => b.kind === 'paragraph');
    const paragraph: Paragraph = { kind: 'paragraph', id: newId(), props: first?.props ?? {}, runs: [] };
    return {
      kind: 'cell',
      id: newId(),
      props: cell.props,
      gridSpan: cell.gridSpan,
      vMerge,
      blocks: [paragraph],
    };
  });
  const row: TableRow = { kind: 'row', id: newId(), props: template.props, cells };
  if (template.propsEx !== undefined) row.propsEx = template.propsEx;
  const at = side === 'above' ? rowIndex : rowIndex + 1;
  return { table: { ...table, rows: [...table.rows.slice(0, at), row, ...table.rows.slice(at)] }, row };
}

/**
 * 删掉 `[from, to]` 这几行（含两端），修好被切断的纵向合并（见文件头 ③）。
 * 全删光答 undefined —— 调用方删整张表。
 */
export function withoutRows(table: Table, from: number, to: number): Table | undefined {
  if (from < 0 || to >= table.rows.length || from > to) throw new RangeError('行范围越界');
  if (from === 0 && to === table.rows.length - 1) return undefined;
  const above = table.rows[from - 1];
  const next = table.rows[to + 1];
  let fixed = next;
  if (next !== undefined) {
    const cells = next.cells.map((cell, i): TableCell => {
      if (cell.vMerge !== 'continue') return cell;
      const head = cellAtColumn(above, columnOf(next, i));
      return head === undefined || head.vMerge === 'none' ? { ...cell, vMerge: 'restart' } : cell;
    });
    if (cells.some((c, i) => c !== next.cells[i])) fixed = { ...next, cells };
  }
  const rows = table.rows.flatMap((r, i) =>
    i >= from && i <= to ? [] : i === to + 1 && fixed ? [fixed] : [r],
  );
  return { ...table, rows };
}

/** 第 `index` 格占的网格区间 [start, end) */
export function cellColumns(row: TableRow, index: number): { start: number; end: number } {
  const start = columnOf(row, index);
  return { start, end: start + (row.cells[index]?.gridSpan ?? 1) };
}

/** `dxa` 型宽度加减一段；其余类型是相对量或没写，不动（见文件头 ⑤） */
function widen<T extends { width?: TableWidth }>(props: T, delta: number): T {
  if (props.width?.type !== 'dxa') return props;
  return { ...props, width: { ...props.width, value: Math.max(0, props.width.value + delta) } };
}

/** 一行在网格边界 `at` 插进 `width` 宽的一列：撑宽跨过它的格、加跳过的列数，或者插一个新格（文件头 ④） */
function withColumnAt(
  row: TableRow,
  at: number,
  side: 'left' | 'right',
  width: Twips,
  newId: () => NodeId,
): TableRow {
  const before = row.props.gridBefore ?? 0;
  const end = before + row.cells.reduce((n, c) => n + c.gridSpan, 0);
  if (at < before) return { ...row, props: { ...row.props, gridBefore: before + 1 } };
  if (at > end) return { ...row, props: { ...row.props, gridAfter: (row.props.gridAfter ?? 0) + 1 } };
  const cells: TableCell[] = [];
  let inserted = false;
  for (const [i, cell] of row.cells.entries()) {
    const { start, end: stop } = cellColumns(row, i);
    if (start < at && at < stop) {
      cells.push({ ...cell, gridSpan: cell.gridSpan + 1, props: widen(cell.props, width) });
      inserted = true;
      continue;
    }
    cells.push(cell);
  }
  if (inserted) return { ...row, cells };
  // 边界正好在两格之间（或一行的首末）：目标格那一侧的邻居作模板，没有就用另一侧
  const right = row.cells.findIndex((_, i) => cellColumns(row, i).start === at);
  const left = row.cells.findIndex((_, i) => cellColumns(row, i).end === at);
  const tplIndex = side === 'left' ? (right >= 0 ? right : left) : left >= 0 ? left : right;
  const template = row.cells[tplIndex] as TableCell;
  const first = template.blocks.find((b): b is Paragraph => b.kind === 'paragraph');
  const props =
    template.props.width?.type === 'dxa'
      ? { ...template.props, width: { value: width, type: 'dxa' as const } }
      : template.props;
  const fresh: TableCell = {
    kind: 'cell',
    id: newId(),
    props,
    gridSpan: 1,
    vMerge: template.vMerge,
    blocks: [{ kind: 'paragraph', id: newId(), props: first?.props ?? {}, runs: [] }],
  };
  const index = right >= 0 ? right : row.cells.length;
  return { ...row, cells: [...row.cells.slice(0, index), fresh, ...row.cells.slice(index)] };
}

/**
 * 在第 `rowIndex` 行第 `cellIndex` 格的左 / 右边插一列（见文件头 ④⑤），返回新表与新格（按行序，
 * 被撑宽而没有新格的行不在里面）。表格没写 `w:tblGrid` 时拒绝 —— 没有网格就说不出「一列」有多宽。
 */
export function withInsertedColumn(
  table: Table,
  rowIndex: number,
  cellIndex: number,
  side: 'left' | 'right',
  newId: () => NodeId,
): { table: Table; cells: TableCell[] } {
  const row = table.rows[rowIndex];
  if (row?.cells[cellIndex] === undefined)
    throw new RangeError(`表格没有第 ${rowIndex} 行第 ${cellIndex} 格`);
  if (table.grid.length === 0) throw new Error('表格没有列网格（w:tblGrid），无法插列');
  const { start, end } = cellColumns(row, cellIndex);
  const at = side === 'left' ? start : end;
  const width = table.grid[side === 'left' ? start : end - 1] ?? 0;
  const known = new Set(table.rows.flatMap((r) => r.cells.map((c) => c.id)));
  const rows = table.rows.map((r) => withColumnAt(r, at, side, width, newId));
  const cells = rows.flatMap((r) => r.cells.filter((c) => !known.has(c.id)));
  const grid = [...table.grid.slice(0, at), width, ...table.grid.slice(at)];
  return { table: { ...table, props: widen(table.props, width), grid, rows }, cells };
}

/** 续格上面同一列没有合并区（首格被删了、或那一行被删了）就升成 restart（见文件头 ③⑥） */
function repairVMerge(rows: TableRow[]): TableRow[] {
  const out: TableRow[] = [];
  for (const row of rows) {
    const above = out[out.length - 1];
    const cells = row.cells.map((cell, i): TableCell => {
      if (cell.vMerge !== 'continue') return cell;
      const head = cellAtColumn(above, columnOf(row, i));
      return head === undefined || head.vMerge === 'none' ? { ...cell, vMerge: 'restart' } : cell;
    });
    out.push(cells.some((c, i) => c !== row.cells[i]) ? { ...row, cells } : row);
  }
  return out;
}

/**
 * 删掉网格列 [from, to)（见文件头 ⑥），返回新表与删掉的格。网格或行删光答 `table: undefined` ——
 * 调用方删整张表。
 */
export function withoutColumns(
  table: Table,
  from: number,
  to: number,
): { table: Table | undefined; removed: TableCell[] } {
  if (from < 0 || to > table.grid.length || from >= to) throw new RangeError('列范围越界');
  const removed: TableCell[] = [];
  const rows: TableRow[] = [];
  const widthOf = (a: number, b: number) => table.grid.slice(a, b).reduce((n, w) => n + w, 0);
  for (const row of table.rows) {
    const props = { ...row.props };
    const before = row.props.gridBefore ?? 0;
    if (before > 0) props.gridBefore = before - Math.max(0, Math.min(before, to) - from);
    const cells: TableCell[] = [];
    for (const [i, cell] of row.cells.entries()) {
      const { start, end } = cellColumns(row, i);
      const lo = Math.max(start, from);
      const hi = Math.min(end, to);
      if (lo >= hi) cells.push(cell);
      else if (lo === start && hi === end) removed.push(cell);
      else
        cells.push({
          ...cell,
          gridSpan: cell.gridSpan - (hi - lo),
          props: widen(cell.props, -widthOf(lo, hi)),
        });
    }
    const last = before + row.cells.reduce((n, c) => n + c.gridSpan, 0);
    const after = row.props.gridAfter ?? 0;
    if (after > 0) props.gridAfter = after - Math.max(0, Math.min(last + after, to) - Math.max(last, from));
    if (props.gridBefore === 0) delete props.gridBefore;
    if (props.gridAfter === 0) delete props.gridAfter;
    if (cells.length === 0) continue; // 一格不剩的行整行删掉
    const same = cells.length === row.cells.length && cells.every((c, i) => c === row.cells[i]);
    rows.push(
      same && props.gridBefore === row.props.gridBefore && props.gridAfter === row.props.gridAfter
        ? row
        : { ...row, props, cells },
    );
  }
  const grid = [...table.grid.slice(0, from), ...table.grid.slice(to)];
  if (grid.length === 0 || rows.length === 0)
    return { table: undefined, removed: table.rows.flatMap((r) => r.cells) };
  return {
    table: { ...table, props: widen(table.props, -widthOf(from, to)), grid, rows: repairVMerge(rows) },
    removed,
  };
}

/** 只有空段落（没有字、没有对象）的格 —— 合并时不贡献内容（见文件头 ⑦） */
function isBlank(cell: TableCell): boolean {
  return cell.blocks.every(
    (b) =>
      b.kind === 'paragraph' &&
      b.runs.every((r) => r.content.every((c) => c.kind === 'text' && c.text === '')),
  );
}

/** 这一行里与网格区间 [from, to) 有交集的格的下标 */
function cellsIn(row: TableRow, from: number, to: number): number[] {
  return row.cells.flatMap((_, i) => {
    const { start, end } = cellColumns(row, i);
    return start < to && end > from ? [i] : [];
  });
}

/**
 * 两个格撑成的矩形（见文件头 ⑦）：行 [top, bottom]、网格列 [from, to)。
 * 区域里有没有格的空缺（`w:gridBefore` / `w:gridAfter` 跳过的列）答 undefined —— 那不是一块能合并的矩形。
 */
export function cellRect(
  table: Table,
  a: { rowIndex: number; cellIndex: number },
  b: { rowIndex: number; cellIndex: number },
): { top: number; bottom: number; from: number; to: number } | undefined {
  const ca = cellColumns(table.rows[a.rowIndex] as TableRow, a.cellIndex);
  const cb = cellColumns(table.rows[b.rowIndex] as TableRow, b.cellIndex);
  let top = Math.min(a.rowIndex, b.rowIndex);
  let bottom = Math.max(a.rowIndex, b.rowIndex);
  let from = Math.min(ca.start, cb.start);
  let to = Math.max(ca.end, cb.end);
  for (let changed = true; changed; ) {
    changed = false;
    for (let r = top; r <= bottom; r++) {
      const row = table.rows[r] as TableRow;
      for (const i of cellsIn(row, from, to)) {
        const { start, end } = cellColumns(row, i);
        const cell = row.cells[i] as TableCell;
        if (start < from || end > to) {
          from = Math.min(from, start);
          to = Math.max(to, end);
          changed = true;
        }
        if (r === top && cell.vMerge === 'continue' && top > 0) {
          top--;
          changed = true;
        }
        const below = cellAtColumn(table.rows[r + 1], start);
        if (r === bottom && cell.vMerge !== 'none' && below?.vMerge === 'continue') {
          bottom++;
          changed = true;
        }
      }
    }
  }
  for (let r = top; r <= bottom; r++) {
    const row = table.rows[r] as TableRow;
    const covered = cellsIn(row, from, to).reduce((n, i) => n + (row.cells[i] as TableCell).gridSpan, 0);
    if (covered !== to - from) return undefined;
  }
  return { top, bottom, from, to };
}

/**
 * 把矩形里的格并成一格（见文件头 ⑦），返回新表、合并后的首格、被丢掉的空段落与续格里新造的空段落。
 * 矩形里只有一格时答 undefined（没什么可合并的）。
 */
export function withMergedCells(
  table: Table,
  rect: { top: number; bottom: number; from: number; to: number },
  newId: () => NodeId,
): { table: Table; head: TableCell; dropped: Paragraph[]; added: NodeId[] } | undefined {
  const { top, bottom, from, to } = rect;
  const picked = table.rows
    .slice(top, bottom + 1)
    .map((row) => cellsIn(row, from, to).map((i) => row.cells[i] as TableCell));
  if (picked.flat().length <= 1) return undefined;
  const all = picked.flat();
  const filled = all.filter((c) => !isBlank(c));
  const first = all[0] as TableCell;
  const blocks = filled.length === 0 ? first.blocks : filled.flatMap((c) => c.blocks);
  const kept = new Set(blocks.map((b) => b.id));
  const dropped = all.flatMap((c) =>
    [...c.blocks].filter((b): b is Paragraph => b.kind === 'paragraph' && !kept.has(b.id)),
  );
  const span = to - from;
  const width = table.grid.length ? table.grid.slice(from, to).reduce((n, w) => n + w, 0) : undefined;
  const sized = (c: TableCell) =>
    width === undefined || c.props.width?.type !== 'dxa'
      ? c.props
      : { ...c.props, width: { value: width, type: 'dxa' as const } };
  const head: TableCell = {
    ...first,
    props: sized(first),
    gridSpan: span,
    vMerge: bottom > top ? 'restart' : 'none',
    blocks,
  };
  const added: NodeId[] = [];
  const rows = table.rows.map((row, r) => {
    if (r < top || r > bottom) return row;
    const own = picked[r - top] as TableCell[];
    const lead = own[0] as TableCell;
    let cell = head;
    if (r > top) {
      const para = lead.blocks.find((b): b is Paragraph => b.kind === 'paragraph');
      const id = newId();
      added.push(id);
      cell = {
        ...lead,
        props: sized(lead),
        gridSpan: span,
        vMerge: 'continue',
        blocks: [{ kind: 'paragraph', id, props: para?.props ?? {}, runs: [] }],
      };
    }
    const at = row.cells.indexOf(lead);
    return { ...row, cells: [...row.cells.slice(0, at), cell, ...row.cells.slice(at + own.length)] };
  });
  return { table: { ...table, rows }, head, dropped, added };
}

/**
 * 拆回合并前的格（见文件头 ⑧）：位置所在的格跨列就拆成一列一格，是纵向合并区的首格就把整个区解开。
 * 新格照原格抄属性（`dxa` 宽按网格列重算），每格一个空段落。没有合并过答 undefined。
 */
export function withSplitCell(
  table: Table,
  rowIndex: number,
  cellIndex: number,
  newId: () => NodeId,
): { table: Table; cells: TableCell[] } | undefined {
  const row = table.rows[rowIndex];
  const target = row?.cells[cellIndex];
  if (row === undefined || target === undefined)
    throw new RangeError(`表格没有第 ${rowIndex} 行第 ${cellIndex} 格`);
  if (target.gridSpan === 1 && target.vMerge !== 'restart') return undefined;
  const { start } = cellColumns(row, cellIndex);
  const added: TableCell[] = [];
  const sized = (c: TableCell, col: number) => {
    const w = table.grid[col];
    return w === undefined || c.props.width?.type !== 'dxa'
      ? c.props
      : { ...c.props, width: { value: w, type: 'dxa' as const } };
  };
  const splitRow = (r: TableRow, cell: TableCell): TableRow => {
    const para = cell.blocks.find((b): b is Paragraph => b.kind === 'paragraph');
    const extra = Array.from({ length: cell.gridSpan - 1 }, (_, k): TableCell => {
      const c: TableCell = {
        kind: 'cell',
        id: newId(),
        props: sized(cell, start + k + 1),
        gridSpan: 1,
        vMerge: 'none',
        blocks: [{ kind: 'paragraph', id: newId(), props: para?.props ?? {}, runs: [] }],
      };
      added.push(c);
      return c;
    });
    const self: TableCell = { ...cell, props: sized(cell, start), gridSpan: 1, vMerge: 'none' };
    const at = r.cells.indexOf(cell);
    return { ...r, cells: [...r.cells.slice(0, at), self, ...extra, ...r.cells.slice(at + 1)] };
  };
  const rows = [...table.rows];
  rows[rowIndex] = splitRow(row, target);
  if (target.vMerge === 'restart')
    for (let r = rowIndex + 1; r < rows.length; r++) {
      const cur = rows[r] as TableRow;
      const cont = cellAtColumn(cur, start);
      if (cont?.vMerge !== 'continue') break;
      rows[r] = splitRow(cur, cont);
    }
  return { table: { ...table, rows }, cells: added };
}
