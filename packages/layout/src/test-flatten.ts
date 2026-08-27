/**
 * 把排完版的一页摊平成「绝对坐标 + 文字 + 字号」的片段 —— **测试专用**，不进公开出口。
 *
 * 抽出来的理由与 `apps/fidelity/src/flatten.ts` 一样：同一套累加规则的抄本已经有好几份了
 * （渲染层的 `paintBlockStack` / `paintCellContent`、语料体检、几个 spike），而
 * 「表格的水平格线也占高度」「拆行的两片各带一份上下边距」这类规则正是抄漏一处就整份
 * 文档往下错的那种。包内的两份真值回归（`table-fixture.test.ts` / `table-split-fixture.test.ts`）
 * 因此共用这一份；`apps/fidelity` 那一份留在那边，它不能反过来依赖 `@uw/layout` 的内部文件。
 *
 * 坐标单位是 **pt**，原点**纸左上角**、y 向下 —— 与 `fixtures/*.truth.json` 同一套。
 */
import type { Twips } from '@uw/core';
import type { PageLayout, PlacedBlock } from './page.ts';
import type { BlockLayout, CellLayout } from './table.ts';
import { contentHeightOf } from './table.ts';

export interface Chunk {
  x: number;
  y: number;
  /** 片段右端（推进宽度算出来的），用来判断相邻两段是不是同一串字被切开的 */
  right: number;
  size: number;
  text: string;
}

const pt = (t: Twips): number => t / 20;

/** 一摞块（格内）从 (x0, y0) 往下排。累加规则与渲染层的 `paintBlockStack` 一致 */
function collectStack(blocks: readonly BlockLayout[], x0: Twips, y0: Twips, out: Chunk[]): void {
  let y = y0;
  for (const b of blocks) {
    if (b.kind === 'table') {
      for (const row of b.layout.rows) {
        for (const cell of row.cells) {
          collectCell(cell, x0 + b.layout.x, y + row.gridAbove, row.height - row.gridAbove, out);
        }
        y += row.height;
      }
      y += b.layout.gridBelow;
      continue;
    }
    y += b.layout.spaceBefore;
    for (const line of b.layout.lines) {
      for (const f of line.fragments) {
        if (f.text === '') continue;
        out.push({
          x: pt(x0 + f.x),
          y: pt(y + line.baseline),
          right: pt(x0 + f.x + f.width),
          size: pt(f.fontSize),
          text: f.text,
        });
      }
      y += line.height;
    }
    y += b.layout.spaceAfter;
  }
}

/** 一格的内容。起始 y 由 `w:vAlign` 决定 —— 与渲染层的 `paintCellContent` 同一套算法 */
export function collectCell(
  cell: CellLayout,
  tableX: Twips,
  rowTop: Twips,
  rowHeight: Twips,
  out: Chunk[],
): void {
  if (cell.vMerge === 'continue') return;
  const inner = contentHeightOf(cell.blocks);
  const avail = rowHeight - cell.paddingTop - cell.paddingBottom;
  let y = rowTop + cell.paddingTop;
  if (cell.verticalAlign === 'center') y += Math.max(0, (avail - inner) / 2);
  else if (cell.verticalAlign === 'bottom') y += Math.max(0, avail - inner);
  collectStack(cell.blocks, tableX + cell.x + cell.paddingLeft, y, out);
}

/** 一页版心里的全部文字片段，坐标相对**纸**左上角，未合并 */
export function chunksOf(page: PageLayout): Chunk[] {
  const out: Chunk[] = [];
  const c = page.geometry.content;
  // `PlacedRow.y` / `PlacedLine.y` 都已经是**相对版心**的绝对 y，别再加块自己的 y
  const walk = (blocks: readonly PlacedBlock[], x0: Twips, y0: Twips): void => {
    for (const b of blocks) {
      if (b.kind === 'paragraph') {
        for (const placed of b.lines) {
          for (const f of placed.line.fragments) {
            if (f.text === '') continue;
            out.push({
              x: pt(x0 + f.x),
              y: pt(y0 + placed.y + placed.line.baseline),
              right: pt(x0 + f.x + f.width),
              size: pt(f.fontSize),
              text: f.text,
            });
          }
        }
        continue;
      }
      for (const placed of b.rows) {
        for (const cell of placed.row.cells) {
          collectCell(
            cell,
            x0 + b.x,
            y0 + placed.y + placed.row.gridAbove,
            placed.height - placed.row.gridAbove,
            out,
          );
        }
      }
    }
  };
  walk(page.blocks, c.x, c.y);
  return out;
}
