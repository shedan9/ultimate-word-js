/**
 * 表格**拆行** —— 一行放不下时，把它内部从某一条行间切开，本页留一片、下一页接一片。
 *
 * 在这之前行是原子的（放不下就整行挪走），等价于「全表 `w:cantSplit`」。
 * 代价不是「难看一点」：一行高过整页版心时，原来那条路只能硬塞（`count = 1`），
 * 于是内容直接溢出版心、后面每一页都跟着错位。公文里「附件说明」这种一格塞一整页
 * 的表格并不罕见。
 *
 * ## 为什么切出来的是**真的两份 `RowLayout`**，而不是「一份 + 裁剪窗口」
 *
 * 后者要渲染层加 `clipPath`、要命中测试知道「这一片只露出第几行到第几行」，
 * 还要多一套「行内局部坐标」。而架构第 1 条说阶段之间只传纯数据 ——
 * 把切开这件事在布局里做完，渲染层一个字都不用改（除了接缝那条线，见下），
 * 命中测试拿到的每一片都是自洽的一格内容。代价是切片会复制一层
 * `CellLayout` 外壳（`lines` 数组本身是切片共享的，不复制行）。
 *
 * ## 切在哪
 *
 * 切口只落在**行间**（段落的两行之间、嵌套表格的两行之间），不切开一行文字 ——
 * 与段落跨页同理。每一格各切各的：某一格的内容全放得下、另一格只放得下两行，
 * 这一片的高度按**最高的那一格**算，与不拆行时 `rowHeight()` 取 max 是同一条规则。
 *
 * ## 几处**没有 Word 真值**的判断（改动它们只改画法，不改断行）
 *
 * - **单元格边距**：上边距只算在**头**片、下边距只算在**尾**片（切口两侧不留边距）。
 *   按「每片都补一遍上下边距」实现的话，一行拆成两片会平白长出一整个边距
 * - **`w:vAlign`**：头片一律按 `top` 摆（居中 / 贴底在「还没排完」的片里没有意义），
 *   尾片沿用原值 —— `w:trHeight` 撑出来的富余高度落在尾片上，贴底才有地方贴
 * - **`w:trHeight` 的富余**（`atLeast` / `exact` 比内容高出来的那一截）整个记在**尾**片：
 *   记在头片会让本页多占一截空白，而那正是拆行要省下来的
 * - **接缝处画不画线**：由渲染层决定，见 `@uw/render-dom` 的 `SPLIT_ROW_SEAM_BORDER`
 *
 * ## 没做
 *
 * - **嵌套表格的行不再往下切**：格子里套的表格按行原子处理。真要切得递归调
 *   `splitRow()`，而「表格里套表格且恰好跨页」在公文里没见过，留个洞比留一套没测过的递归好
 * - **格内的孤行寡行**：`w:widowControl` 在格子里不生效，切口就按放得下多少行定
 */
import type { Twips } from '@uw/core';
import type { BlockLayout, CellLayout, RowLayout } from './table.ts';
import { contentHeightOf } from './table.ts';

export interface RowSplit {
  /** 留在本页的那一片 */
  head: RowLayout;
  /** 接到下一页的那一片。它可能还要再切一次（一行高过好几页） */
  tail: RowLayout;
}

/**
 * 把一行切成「本页放得下的一片 + 剩下的一片」。
 *
 * 切不动就返回 `undefined`，调用方退回「整行挪到下一页」那条老路。切不动有两种：
 * ① 一片文字都放不进 `avail`（哪怕第一行）；② 所有内容都放得下（那本来就不该走到这儿，
 * 只可能是下边距差的那一点点 —— 让整行挪走比切出一片空尾巴强）。
 */
export function splitRow(row: RowLayout, avail: Twips): RowSplit | undefined {
  if (avail <= 0) return undefined;

  const heads: CellLayout[] = [];
  const tails: CellLayout[] = [];
  let placed = false; // 头片里至少落下了一行
  let leftover = false; // 尾片里还有内容

  for (const cell of row.cells) {
    // 没有自己内容的格子（`vMerge="continue"` 的、空格子）两片都留一个空壳 ——
    // 渲染层靠它画底纹与格线，少一个格子整条竖线就断了
    if (cell.vMerge === 'continue' || cell.blocks.length === 0) {
      heads.push(sliceCell(cell, [], 'head'));
      tails.push(sliceCell(cell, [], 'tail'));
      continue;
    }
    const cut = splitBlocks(cell.blocks, avail - cell.paddingTop);
    if (cut.head.length > 0) placed = true;
    if (cut.tail.length > 0) leftover = true;
    heads.push(sliceCell(cell, cut.head, 'head'));
    tails.push(sliceCell(cell, cut.tail, 'tail'));
  }

  if (!placed || !leftover) return undefined;

  // `w:trHeight` 比内容高出来的那一截：整个记在尾片（见文件头）
  const surplus = Math.max(0, row.height - maxContentHeight(row.cells));
  return {
    head: { rowId: row.rowId, cells: heads, height: maxContentHeight(heads) },
    tail: { rowId: row.rowId, cells: tails, height: maxContentHeight(tails) + surplus },
  };
}

/** 一行的高度 = 最高那一格的内容高。与 `table.ts` 的 `rowHeight()` 同一条规则 */
function maxContentHeight(cells: readonly CellLayout[]): Twips {
  let h: Twips = 0;
  for (const c of cells) {
    if (c.vMerge === 'continue') continue;
    if (c.contentHeight > h) h = c.contentHeight;
  }
  return h;
}

/**
 * 一格的一片。外壳（x / 宽 / 边框 / 底纹）原样带过去 —— 边框在两片上都要画，
 * 竖边各画各的那一截，横边由 `SPLIT_ROW_SEAM_BORDER` 决定接缝上那条画不画。
 */
function sliceCell(cell: CellLayout, blocks: BlockLayout[], which: 'head' | 'tail'): CellLayout {
  const paddingTop = which === 'head' ? cell.paddingTop : 0;
  const paddingBottom = which === 'head' ? 0 : cell.paddingBottom;
  return {
    ...cell,
    paddingTop,
    paddingBottom,
    verticalAlign: which === 'head' ? 'top' : cell.verticalAlign,
    blocks,
    contentHeight: contentHeightOf(blocks) + paddingTop + paddingBottom,
  };
}

interface BlockCut {
  head: BlockLayout[];
  tail: BlockLayout[];
}

/** 一摞块按 `avail` 切开。累加规则必须与 `contentHeightOf()` 一致，否则切口会漂 */
function splitBlocks(blocks: readonly BlockLayout[], avail: Twips): BlockCut {
  const head: BlockLayout[] = [];
  const tail: BlockLayout[] = [];
  let used: Twips = 0;
  let cutting = true;

  for (const b of blocks) {
    if (!cutting) {
      tail.push(b);
      continue;
    }
    const h = contentHeightOf([b]);
    if (used + h <= avail) {
      head.push(b);
      used += h;
      continue;
    }
    const cut = splitBlock(b, avail - used);
    if (cut.head !== undefined) head.push(cut.head);
    if (cut.tail !== undefined) tail.push(cut.tail);
    cutting = false;
  }
  return { head, tail };
}

/** 一个块切开。两半都可能缺席：整块都放不下（只有 tail）、只有段后间距溢出（只有 head） */
function splitBlock(b: BlockLayout, avail: Twips): { head?: BlockLayout; tail?: BlockLayout } {
  if (b.kind === 'table') {
    // 嵌套表格的行原子处理，不再往下递归（见文件头）
    const t = b.layout;
    let used: Twips = 0;
    let n = 0;
    for (const r of t.rows) {
      if (used + r.height > avail) break;
      used += r.height;
      n += 1;
    }
    if (n === 0) return { tail: b };
    if (n >= t.rows.length) return { head: b };
    return {
      head: { kind: 'table', layout: { ...t, rows: t.rows.slice(0, n) } },
      tail: { kind: 'table', layout: { ...t, rows: t.rows.slice(n) } },
    };
  }

  const p = b.layout;
  let used: Twips = p.spaceBefore;
  let n = 0;
  for (const line of p.lines) {
    if (used + line.height > avail) break;
    used += line.height;
    n += 1;
  }
  // 段前间距就把 avail 吃光了：整段推到下一片，而不是切出一片只有间距的空壳
  if (n === 0) return { tail: b };
  // 行全放得下、只有段后间距溢出：段后间距落在页底本来就不用留（与段落跨页同理）
  if (n >= p.lines.length) return { head: { kind: 'paragraph', layout: { ...p, spaceAfter: 0 } } };
  return {
    head: { kind: 'paragraph', layout: { ...p, lines: p.lines.slice(0, n), spaceAfter: 0 } },
    tail: { kind: 'paragraph', layout: { ...p, lines: p.lines.slice(n), spaceBefore: 0 } },
  };
}
