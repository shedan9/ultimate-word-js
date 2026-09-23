/**
 * Word 表格里的 Tab / Shift+Tab：跳到下一个 / 上一个单元格并选中它的全部内容。
 *
 * 两处容易搞反：
 * ① 找的是**最内层**那张表的格 —— 嵌套表格里按 Tab 在内表里走，走到内表末格就停，
 *    不会跳回外表（Word 也是这样，末格 Tab 在内表加一行）；
 * ② `vMerge=continue` 的格跳过：它的内容不显示，由上面那一格占位，选中它等于选中一块看不见的地方。
 *    行与行之间照行序接着走（第一行末格的下一格是第二行首格）。
 */
import type { BlockNode, DocumentBody, NodeId, PropSet, TableCellNode } from './nodes.ts';
import { rangeOfNode } from './order.ts';
import type { DocRange } from './position.ts';

interface CellHit<S extends PropSet> {
  cells: TableCellNode<S>[];
  index: number;
}

/**
 * `nodeId`（段落或 run）所在单元格的上 / 下一格的内容范围。
 * `undefined`：不在单元格里（Tab 照常插制表位）；`null`：已经在首 / 末格 ——
 * Word 在末格按 Tab 会加一行，表格结构编辑还没有，调用方应当什么都不做而不是插制表位。
 */
export function adjacentCellRange<S extends PropSet>(
  body: DocumentBody<S>,
  nodeId: NodeId,
  direction: 'backward' | 'forward',
): DocRange | null | undefined {
  const TOP = { cells: [], index: -1 } satisfies CellHit<S>;
  function visit(blocks: readonly BlockNode<S>[], host: CellHit<S>): CellHit<S> | undefined {
    for (const block of blocks) {
      if (block.kind === 'paragraph') {
        if (block.id === nodeId || block.runs.some((r) => r.id === nodeId)) return host;
        continue;
      }
      const cells = block.rows.flatMap((row) => row.cells);
      for (const [index, cell] of cells.entries()) {
        const hit = visit(cell.blocks, { cells, index });
        if (hit) return hit;
      }
    }
    return undefined;
  }
  let hit: CellHit<S> | undefined;
  for (const section of body.sections) {
    hit = visit(section.blocks, TOP);
    if (hit) break;
  }
  if (!hit || hit === TOP) return undefined;
  const step = direction === 'forward' ? 1 : -1;
  for (let i = hit.index + step; i >= 0 && i < hit.cells.length; i += step) {
    const cell = hit.cells[i] as TableCellNode<S>;
    if (cell.vMerge === 'continue') continue;
    const range = rangeOfNode(cell);
    if (range) return range;
  }
  return null;
}
