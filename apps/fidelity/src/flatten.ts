/**
 * 把排完版的一页摊平成「绝对坐标 + 文字 + 字号」的片段。
 *
 * 抽出来是因为**同一套累加规则现在有三个消费者**：语料体检（`corpus-report.ts`，要按 y
 * 分桶重做分行）、表格穿刺（`spike-table.ts`，要逐格比 x / y / 字号）、以及渲染层自己
 * （`@uw/render-dom` 的 `paintBlockStack` / `paintCellContent`）。前两个抄第三个抄了两遍，
 * 而「表格的水平格线也占高度」这条刚加进来的规则正好是抄漏一处就整份文档往上错的那种 ——
 * 所以这一侧收成一份，渲染层那一份留在包内（它是产物代码，不该被工具脚本牵住）。
 *
 * 坐标单位是 **pt**，原点**纸左上角**、y 向下 —— 与 `fixtures/*.truth.json` 同一套，
 * 比对时不用换算。
 */
import type { Twips } from '@uw/core';
import type { BlockLayout, CellLayout, PageLayout, PlacedBlock } from '@uw/layout';
import { contentHeightOf } from '@uw/layout';

export interface Piece {
  /** 片段左端 x（pt） */
  x: number;
  /** 基线 y（pt） */
  y: number;
  /** 片段推进宽度（pt）—— 用来判断相邻两个片段是不是同一段被切开的 */
  w: number;
  /** 字号（pt）—— 表格条件格式那一路靠它认「这一格命中了哪个条件」 */
  size: number;
  text: string;
}

const pt = (t: Twips): number => t / 20;

/** 一摞块（格内 / 页眉里）从 (x0, y0) 往下排。累加规则与渲染层的 `paintBlockStack` 一致 */
export function collectStack(blocks: readonly BlockLayout[], x0: Twips, y0: Twips, out: Piece[]): void {
  let y = y0;
  for (const b of blocks) {
    if (b.kind === 'table') {
      for (const row of b.layout.rows) {
        // 行的上边那条格线占着 `row.height` 的头一段，内容从格线**以内**起排
        for (const cell of row.cells) {
          collectCell(cell, x0 + b.layout.x, y + row.gridAbove, row.height - row.gridAbove, out);
        }
        y += row.height;
      }
      y += b.layout.gridBelow; // 表底那条线不属于任何一行
      continue;
    }
    y += b.layout.spaceBefore;
    for (const line of b.layout.lines) {
      for (const f of line.fragments) {
        if (f.text !== '') {
          out.push({
            x: pt(x0 + f.x),
            y: pt(y + line.baseline),
            w: pt(f.width),
            size: pt(f.fontSize),
            text: f.text,
          });
        }
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
  out: Piece[],
): void {
  if (cell.vMerge === 'continue') return; // 内容由上面那个 restart 撑着，这一格不画
  const inner = contentHeightOf(cell.blocks);
  const avail = rowHeight - cell.paddingTop - cell.paddingBottom;
  let y = rowTop + cell.paddingTop;
  if (cell.verticalAlign === 'center') y += Math.max(0, (avail - inner) / 2);
  else if (cell.verticalAlign === 'bottom') y += Math.max(0, avail - inner);
  collectStack(cell.blocks, tableX + cell.x + cell.paddingLeft, y, out);
}

/** 一批已分页的块（正文 / 页眉页脚），坐标相对给定的框左上角 */
export function collectBlocks(blocks: readonly PlacedBlock[], x0: Twips, y0: Twips, out: Piece[]): void {
  for (const b of blocks) {
    if (b.kind === 'paragraph') {
      for (const placed of b.lines) {
        for (const f of placed.line.fragments) {
          if (f.text !== '') {
            out.push({
              x: pt(x0 + f.x),
              y: pt(y0 + placed.y + placed.line.baseline),
              w: pt(f.width),
              size: pt(f.fontSize),
              text: f.text,
            });
          }
        }
      }
      continue;
    }
    // `PlacedRow.y` 已经是**相对版心**的绝对 y（与 `PlacedLine.y` 同一套，渲染层
    // 的 `paintPlacedTable` 也是直接用它），再加一次 `b.y` 就把表格整体往下推了
    // 一个块的高度 —— 表格越靠后错得越多。这一处原来抄错过一次，记在这儿。
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
}

/** 一整页（版心 + 页眉 + 页脚）的全部文字片段，坐标相对**纸**左上角 */
export function piecesOf(page: PageLayout): Piece[] {
  const out: Piece[] = [];
  const c = page.geometry.content;
  collectBlocks(page.blocks, c.x, c.y, out);
  for (const hf of [page.header, page.footer]) {
    if (hf !== undefined) collectBlocks(hf.blocks, hf.x, hf.y, out);
  }
  return out;
}
