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
 */
import type { Block, Body, NodeId, Paragraph, Table, TableCell, TableRow } from './nodes.ts';

/** 一张表在树上的位置：它所在的块列表 + 下标，以及沿途的祖先（重建路径用） */
interface TableHit {
  table: Table;
  rowIndex: number;
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
