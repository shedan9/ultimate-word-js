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
 * 把切开这件事在布局里做完，渲染层一个字都不用改，命中测试拿到的每一片都是自洽的一格内容。
 * 代价是切片会复制一层 `CellLayout` 外壳（`lines` 数组本身是切片共享的，不复制行）。
 *
 * ## 切在哪
 *
 * 切口只落在**行间**（段落的两行之间、嵌套表格的两行之间），不切开一行文字 ——
 * 与段落跨页同理。每一格各切各的：某一格的内容全放得下、另一格只放得下两行，
 * 这一片的高度按**最高的那一格**算，与不拆行时 `rowHeight()` 取 max 是同一条规则。
 *
 * ## 四问已经实测（`spike-table-04`，跑 `pnpm --filter @uw/fidelity spike:table-split`）
 *
 * 规则与证据表见下面的 `TABLE_SPLIT_RULES`。四条原来全是「哪种最省地方」猜的，
 * **三条猜反了**：边距两片各补一整份（不是上归头、下归尾）、`w:trHeight` 每一片各要一份
 * （不是整行算完把富余留给尾片）、头片照样认 `w:vAlign`（不是一律 top）。
 * 第四问「接缝上画不画线」的答案是**画**，而且画的是**表级**的上下边框，见 `seamBorders()`。
 *
 * ## 没做
 *
 * - **嵌套表格的行不再往下切**：格子里套的表格按行原子处理。真要切得递归调
 *   `splitRow()`，而「表格里套表格且恰好跨页」在公文里没见过，留个洞比留一套没测过的递归好
 * - **格内的孤行寡行**：`w:widowControl` 在格子里不生效，切口就按放得下多少行定
 */
import type { Twips } from '@uw/core';
import type { Border, TableBorders } from '@uw/model';
import type { BlockLayout, CellLayout, RowLayout } from './table.ts';
import { contentHeightOf } from './table.ts';
import type { BorderSegment } from './table-borders.ts';

/**
 * 拆行的四条规则。**都由 `spike-table-04` 实测**（14 页、七张表，跑
 * `pnpm --filter @uw/fidelity spike:table-split`），实现的这一组是 16 种组合里唯一满分的。
 *
 * | 问 | 原来猜的 | Word | 证据 |
 * |---|---|---|---|
 * | 切在哪一页 | 整行挪到下一页顶上再切 | **就地切**（本页剩下多少用多少） | 表甲：Word 在首页放下了甲 01–甲 10，我们整行挪走、首页只剩表头一行 |
 * | 单元格上下边距 | 上归头片、下归尾片 | **两片各补一整份** | 表甲边距 20pt：头片最后一行离接缝正好 20pt，尾片第一行离接缝也是 20pt；按「只补一边」算头片能多放两行 |
 * | `w:trHeight` 的富余 | 整行算完，富余归尾片 | **每一片各要一份** | 表丁 atLeast 200pt：尾片只有 9 行文字（140pt）却撑到 202pt；表戊 atLeast 100pt、尾片 4 行（62pt）撑到 100.1pt |
 * | 头片的 `w:vAlign` | 一律 top | **照原样** | 表丙头片高 218pt：`bottom` 那一格的字落在片底（基线 335.98），`center` 那一格落在片中（234.79），都不在顶上 |
 *
 * 「每一片各要一份 `w:trHeight`」顺手解释了另外两件本来要单独写规则的事：
 * ① **什么时候不就地切** —— 表乙 atLeast 420pt、本页还剩 266pt，Word 把整行挪到了下一页：
 *    一片都满足不了 420 就没得切。这不是另一条规则，是同一条的推论；
 * ② **续页顶上还重不重复表头** —— 表乙的两片各占满一整页，两页顶上都**没有**重复表头；
 *    表丁的尾片只要 200pt，页顶就照常重复了一遍（见 `page.ts` 的 `placeTable`）。
 *
 * 尾片顶上那条接缝线**像正常格线一样占高度**（`gridAbove`），而它的粗细是**表级上边框**
 * 的粗细，不是这一行自己那条。表丁看起来「比 trHeight 高出 1.9pt」正是量错了这一段：
 * 那张表的外框 2.25pt，接缝线占掉 2.25pt 之后，格线以内正好是 200.1pt = trHeight。
 * 表戊（外框 0.5pt）与表庚（不拆行的对照）都对得上同一条规则。
 *
 * 三处**仍然没有真值**（都不改断行）：
 * - **头片不为接缝那条线预留高度**：表甲头片的内容底 + 20pt 边距 = 接缝线的位置，
 *   线本身画在片外（324.58 + 2.16 仍在版心内）。线粗到跨出版心时 Word 怎么办没量过
 * - 接缝那条线取的是**表级** `w:tblBorders` 的上下边：拆开的那一行自己写了
 *   `w:tcBorders` 时听谁的，样本里没有这个局面
 * - **接缝线与重复表头撞在同一个 y 上时画哪一条**：表丁的尾片顶着重复表头，
 *   两条线都落在 74.54 —— Word 画的是表头那条 0.5pt 的 `insideH`，我们会画 2.25pt 的
 *   接缝线。它**不改几何**（那 2.25pt 的高度 Word 照样留着，只是没画出来）
 */
export interface TableSplitRules {
  /** 放不下的那一行：就地切开，还是先整行挪到下一页顶上再切 */
  place: 'inPlace' | 'nextPage';
  /** 单元格上下边距：两片各补一整份，还是上边距归头片、下边距归尾片 */
  margins: 'both' | 'split';
  /** `w:trHeight`：每一片各要一份，还是按整行算、富余归尾片 */
  trHeight: 'perPiece' | 'wholeRow';
  /** 头片的 `w:vAlign`：照原样，还是一律按 top 摆 */
  headVAlign: 'keep' | 'top';
}

export const TABLE_SPLIT_RULES: TableSplitRules = {
  place: 'inPlace',
  margins: 'both',
  trHeight: 'perPiece',
  headVAlign: 'keep',
};

export interface SplitRowOptions {
  /** 标定用的接缝，正常调用不要传 */
  rules?: TableSplitRules;
  /** 这一行 `w:trHeight` 要的高度（`atLeast` / `exact` 的值，auto 是 0），**格线以内** */
  requested?: Twips;
  /** 表级 `w:tblBorders`：接缝那两条线就是它的上下边（见 `seamBorders()`） */
  tableBorders?: TableBorders;
}

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
export function splitRow(row: RowLayout, avail: Twips, opts: SplitRowOptions = {}): RowSplit | undefined {
  if (avail <= 0) return undefined;
  const rules = opts.rules ?? TABLE_SPLIT_RULES;
  const requested = opts.requested ?? 0;
  // 尾片的顶边是**表级上边框**（实测，见 `seamBorders()`），它像正常格线一样占高度；
  // 头片的顶边仍是这一行自己那条真格线
  const seam = seamBorders(opts.tableBorders);
  const tailGrid = borderThickness(seam.top);

  const heads: CellLayout[] = [];
  const tails: CellLayout[] = [];
  let placed = false; // 头片里至少落下了一行
  let leftover = false; // 尾片里还有内容

  for (const cell of row.cells) {
    // 没有自己内容的格子（`vMerge="continue"` 的、空格子）两片都留一个空壳 ——
    // 渲染层靠它画底纹与格线，少一个格子整条竖线就断了
    if (cell.vMerge === 'continue' || cell.blocks.length === 0) {
      heads.push(sliceCell(cell, [], 'head', rules, seam));
      tails.push(sliceCell(cell, [], 'tail', rules, seam));
      continue;
    }
    // 头片留不留下边距是 `margins` 那一问：留（实测）的话本页少放的正好是一个下边距
    const room =
      avail - row.gridAbove - cell.paddingTop - (rules.margins === 'both' ? cell.paddingBottom : 0);
    const cut = splitBlocks(cell.blocks, room);
    if (cut.head.length > 0) placed = true;
    if (cut.tail.length > 0) leftover = true;
    heads.push(sliceCell(cell, cut.head, 'head', rules, seam));
    tails.push(sliceCell(cell, cut.tail, 'tail', rules, seam));
  }

  if (!placed || !leftover) return undefined;

  // `w:trHeight` 每一片各要一份（实测）。头片被本页剩下的地方封着顶，所以它要的那一份
  // 还得夹在 `avail` 以内 —— 表乙的两片各占满一整页正是这么来的。
  // 「按整行算、富余归尾片」是原来的写法，`wholeRow` 留着给标定脚本排组合。
  const surplus = Math.max(0, row.height - row.gridAbove - maxContentHeight(row.cells));
  const headMin = rules.trHeight === 'perPiece' ? Math.min(requested, avail - row.gridAbove) : 0;
  const tailMin = rules.trHeight === 'perPiece' ? requested : 0;
  const tailExtra = rules.trHeight === 'perPiece' ? 0 : surplus;
  return {
    head: {
      rowId: row.rowId,
      cells: heads,
      gridAbove: row.gridAbove,
      height: row.gridAbove + Math.max(maxContentHeight(heads), headMin),
    },
    tail: {
      rowId: row.rowId,
      cells: tails,
      gridAbove: tailGrid,
      height: tailGrid + Math.max(maxContentHeight(tails) + tailExtra, tailMin),
    },
  };
}

/**
 * 接缝上那两条线：**头片的底取表级下边框、尾片的顶取表级上边框**（`spike-table-04` 实测）。
 *
 * 直觉上应该取「这一行自己的上下边框」，实测不是：表己的第二行后面还跟着一行，
 * 它自己的下边框是 3pt 的绿 `insideH`，而接缝上画出来的是 0.5pt 的黑线 —— 表级的下边框。
 * 尾片那一侧同理（页顶画的是表级上边框，不是 insideH）。也就是说 Word 把每一页上的表格
 * 片段**当成一张自己封口的表**来画，接缝正是这张表的上下口。
 *
 * 竖边不用管：两片各画各的那一截，本来就没有接缝这回事。
 */
function seamBorders(borders: TableBorders | undefined): {
  top: Border | undefined;
  bottom: Border | undefined;
} {
  return { top: borders?.top, bottom: borders?.bottom };
}

/** 边框占的高度。与 `table.ts` 一致：`nil` / `none` / 没写都是 0 */
function borderThickness(b: Border | undefined): Twips {
  if (b === undefined || b.style === 'nil' || b.style === 'none') return 0;
  return b.size;
}

/** 把一条横边整段换成接缝那条线（一格只有一段，跨列的也是一段） */
function seamSegments(cell: CellLayout, border: Border | undefined): BorderSegment[] {
  return [{ col: cell.col, span: cell.span, border }];
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
 * 一格的一片。外壳（x / 宽 / 底纹）原样带过去，只有**接缝那一侧的横边**换成表级的
 * 上 / 下边框（`seamBorders()`）—— 边框在两片上都要画，竖边各画各的那一截。
 */
function sliceCell(
  cell: CellLayout,
  blocks: BlockLayout[],
  which: 'head' | 'tail',
  rules: TableSplitRules,
  seam: { top: Border | undefined; bottom: Border | undefined },
): CellLayout {
  const both = rules.margins === 'both';
  const paddingTop = both || which === 'head' ? cell.paddingTop : 0;
  const paddingBottom = both || which === 'tail' ? cell.paddingBottom : 0;
  // 接缝那一侧的横边换成表级的上 / 下边框：渲染层照旧只认 `cell.borders`，
  // 不必知道「这一片是不是切出来的」（架构第 1 条：阶段之间只传纯数据）
  const borders = {
    ...cell.borders,
    ...(which === 'head'
      ? { bottom: seamSegments(cell, seam.bottom) }
      : { top: seamSegments(cell, seam.top) }),
  };
  return {
    ...cell,
    paddingTop,
    paddingBottom,
    borders,
    verticalAlign: which === 'head' && rules.headVAlign === 'top' ? 'top' : cell.verticalAlign,
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
