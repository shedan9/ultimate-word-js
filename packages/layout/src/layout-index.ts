/**
 * 布局索引：模型位置 ↔ 布局坐标（架构 §4 的 ①↔②）。
 *
 * 命中测试（点了哪个字）、装饰锚定（这个 range 该画在哪几个矩形上）、可选文本层
 * （复制出来的顺序）都由它回答。屏幕坐标那一跳（②↔③）不在这里 —— 那是 `@uw/view`
 * 的事，它要知道缩放与滚动，而这两样每个视图各不相同。
 *
 * **它是消费侧现建的，不是流水线的产物**（与 `@uw/model` 的 `StyleSheet` 同理）：
 * 带方法的对象过不了结构化克隆，而 `DocumentLayout` 必须能过 Worker 边界（原则 1.1）。
 * 所以索引只从 `DocumentLayout` 这一份纯数据里建，主线程拿到什么就能建出什么，
 * 不必把模型也搬过来。
 *
 * **它没有新的标定**：字摆在哪由已经标定完的那几层决定（行高 / 断行 / 分页 / 表格），
 * 索引只保证「画在那儿的字，点它能点中」—— 判据是单测，不是与真值的残差。
 *
 * 两处与架构 §4 原来的写法不一样，都是写实现时才看清的：
 *
 * ① **②→① 不是「一次二分查找」**。表格让同一个 y 上并排坐着好几行（一行一格），
 *    行序与 y 序不再一致，二分的前提就不成立。改成「按页分桶后线性扫，取离点最近的行」——
 *    一页几十行，一次命中测试几十次比较，和二分不在一个量级上值得优化的位置
 * ② **`DocPosition` 是三个字段**（run + 内容片段下标 + 片段内偏移），不是 api.md 原先
 *    写的两个，理由见 `@uw/model` 的 `position.ts`：run 的内容是一列片段，
 *    「run 内的全局偏移」要有模型才算得出来，而索引这一侧没有模型
 */
import type { Twips } from '@uw/core';
import type { DocPosition, DocRange, NodeId } from '@uw/model';
import type { PlacedHeaderFooter } from './header-footer.ts';
import type { DocumentLayout, PlacedBlock } from './page.ts';
import type { BlockLayout, CellLayout, RowLayout } from './table.ts';
import { contentHeightOf } from './table.ts';
import type { LineFragment, LineLayout } from './types.ts';

/** 布局空间的一个点。`page` 是**物理页序**（`PageLayout.index`），x / y 相对**纸左上角** */
export interface LayoutPoint {
  page: number;
  x: Twips;
  y: Twips;
}

/** 布局空间的一个矩形，坐标同样相对纸左上角。一个 range 跨行会给出好几个 */
export interface LayoutRect {
  page: number;
  x: Twips;
  y: Twips;
  width: Twips;
  height: Twips;
}

/**
 * 索引里的一行 —— 把整份 `DocumentLayout` 摊平成的那张表。
 *
 * 行本身（`LineLayout`）的坐标是相对**容器**的（版心 / 单元格内容区 / 页眉框），
 * 摊平这一步做的正是「把容器原点算出来」：`originX + line.x` 才是纸上的位置。
 * 摊平的顺序 = 画的顺序 = 文档顺序（`page.ts` 与渲染层就是这么摞的），
 * 所以下标本身就是文档序，比较位置先后靠的就是它。
 */
export interface IndexedLine {
  page: number;
  /** 容器原点（纸坐标）：`line.x` 与 `LineFragment.x` 都相对它 */
  originX: Twips;
  /** 行顶（纸坐标）。基线 = `top + line.baseline` */
  top: Twips;
  line: LineLayout;
  /**
   * 跨页重复出来的表头行里的行（`PlacedRow.repeated`）。
   *
   * **收进索引但标出来**：Word 里点重复表头是能把光标放进去的（放进的是文档里那一份），
   * 命中测试照着做才对；但**可选文本层必须跳过它**，否则复制出来每页都多一遍表头。
   * 前向映射（`rectsOf`）照样给它的矩形 —— 选中表头行时每页都高亮，也是 Word 的样子。
   */
  repeated: boolean;
  /** 页眉 / 页脚里的行。正文的行没有这个字段 */
  frame?: 'header' | 'footer';
}

export interface LayoutIndex {
  /** 摊平后的全部行，**文档序**。可选文本层直接顺着它走（记得跳过 `repeated`） */
  readonly lines: readonly IndexedLine[];
  /**
   * 命中测试：布局坐标 → 模型位置。取离点最近的行，再取行内离点最近的**字缝**
   * （光标落点），所以点在行外、页边距上、两栏之间都答得出来。
   *
   * 给不出位置时返回 `undefined`：整页没有可定位的文字（空页、只有编号或域结果的行）。
   * 编号与域结果**不可定位** —— 编号不在 document.xml 里、域结果不是文件里那串字符，
   * 拿它们的位置去反查只会指到别处（见 `CharItem.numbering` / `CharItem.field`）。
   */
  positionAt(point: LayoutPoint): DocPosition | undefined;
  /** 一个 range 覆盖的矩形，一行一个（一行内的相邻片段已经并成一段） */
  rectsOf(range: DocRange): LayoutRect[];
  /** 光标矩形：宽 0、高一行。`scrollTo` 与 overlay 锚定用它 */
  caretRect(pos: DocPosition): LayoutRect | undefined;
  /**
   * 文档序比较。任一端的 run 没排出来时返回 `undefined` ——
   * 空 run、`vMerge="continue"` 的格子、域的指令 run 都排不出任何一个字形，
   * 硬给一个 0 会让调用方以为两点相等
   */
  compare(a: DocPosition, b: DocPosition): -1 | 0 | 1 | undefined;
}

/** 片段在索引里的位置：第几行、行里第几个片段 */
interface FragRef {
  li: number;
  fi: number;
}

export function buildLayoutIndex(doc: DocumentLayout): LayoutIndex {
  const lines: IndexedLine[] = [];
  for (const page of doc.pages) {
    const g = page.geometry;
    if (page.header !== undefined) collectFrame(page.header, page.index, lines);
    for (const block of page.blocks) {
      collectBlock(block, g.content.x, g.content.y, page.index, false, undefined, lines);
    }
    if (page.footer !== undefined) collectFrame(page.footer, page.index, lines);
  }

  // run 的文档序 + 每个 run 的片段落在哪几行。两张表一趟建完：
  // 顺序取**第一次出现**，重复表头与跨页拆开的段落因此不会把顺序搅乱
  const order = new Map<NodeId, number>();
  const byRun = new Map<NodeId, FragRef[]>();
  const byPage = new Map<number, number[]>();
  for (let li = 0; li < lines.length; li++) {
    const l = lines[li] as IndexedLine;
    const bucket = byPage.get(l.page);
    if (bucket === undefined) byPage.set(l.page, [li]);
    else bucket.push(li);
    for (let fi = 0; fi < l.line.fragments.length; fi++) {
      const frag = l.line.fragments[fi] as LineFragment;
      if (!order.has(frag.runId)) order.set(frag.runId, order.size);
      const refs = byRun.get(frag.runId);
      if (refs === undefined) byRun.set(frag.runId, [{ li, fi }]);
      else refs.push({ li, fi });
    }
  }

  const fragAt = (ref: FragRef): { line: IndexedLine; frag: LineFragment } => {
    const line = lines[ref.li] as IndexedLine;
    return { line, frag: line.line.fragments[ref.fi] as LineFragment };
  };

  const compare = (a: DocPosition, b: DocPosition): -1 | 0 | 1 | undefined => {
    const oa = order.get(a.nodeId);
    const ob = order.get(b.nodeId);
    if (oa === undefined || ob === undefined) return undefined;
    return sign(cmpKey(oa, a.contentIndex, a.offset, ob, b.contentIndex, b.offset));
  };

  return {
    lines,

    positionAt(point: LayoutPoint): DocPosition | undefined {
      const bucket = byPage.get(point.page);
      if (bucket === undefined) return undefined;
      let best: IndexedLine | undefined;
      let bestDy = Number.POSITIVE_INFINITY;
      let bestDx = Number.POSITIVE_INFINITY;
      for (const li of bucket) {
        const l = lines[li] as IndexedLine;
        const dy = gap(point.y, l.top, l.top + l.line.height);
        const x0 = l.originX + l.line.x;
        const dx = gap(point.x, x0, x0 + l.line.width);
        // 先比纵向再比横向：同一个 y 上并排的几行（表格的一行几格）由横向分胜负，
        // 上下两行则永远由纵向分 —— 反过来会让「点在字的正上方」落到隔壁格里
        if (dy < bestDy || (dy === bestDy && dx < bestDx)) {
          best = l;
          bestDy = dy;
          bestDx = dx;
        }
      }
      if (best === undefined) return undefined;
      return caretIn(best, point.x - best.originX);
    },

    rectsOf(range: DocRange): LayoutRect[] {
      const rel = compare(range.start, range.end);
      if (rel === undefined) return [];
      const [s, e] = rel === 1 ? [range.end, range.start] : [range.start, range.end];
      const so = order.get(s.nodeId) as number;
      const eo = order.get(e.nodeId) as number;

      const out: LayoutRect[] = [];
      let cur: LayoutRect | undefined;
      let curLine = -1;
      for (let li = 0; li < lines.length; li++) {
        const l = lines[li] as IndexedLine;
        for (const frag of l.line.fragments) {
          const span = coveredSpan(frag, order, so, s, eo, e);
          if (span === undefined) continue;
          const [x1, x2] = fragXRange(frag, span[0], span[1]);
          const x = l.originX + x1;
          // 同一行里相邻的片段并成一个矩形：一行一格是装饰层最省事的形状，
          // 也免得「同一款字体切成两段」这种纯渲染的分片漏到调用方那里
          if (cur !== undefined && curLine === li && x - (cur.x + cur.width) <= 0) {
            cur.width = x + (x2 - x1) - cur.x;
            continue;
          }
          cur = { page: l.page, x, y: l.top, width: x2 - x1, height: l.line.height };
          curLine = li;
          out.push(cur);
        }
      }
      return out;
    },

    caretRect(pos: DocPosition): LayoutRect | undefined {
      const refs = byRun.get(pos.nodeId);
      if (refs === undefined) return undefined;
      let tail: LayoutRect | undefined;
      for (const ref of refs) {
        const { line, frag } = fragAt(ref);
        if (frag.offset < 0 || frag.contentIndex !== pos.contentIndex) continue;
        const end = frag.offset + frag.text.length;
        if (pos.offset < frag.offset || pos.offset > end) continue;
        const [x1] = fragXRange(frag, pos.offset, pos.offset);
        const rect = {
          page: line.page,
          x: line.originX + x1,
          y: line.top,
          width: 0,
          height: line.line.height,
        };
        // 位置正好落在片段末尾时，它既是「这一段的末尾」也是「下一段的开头」——
        // 优先给下一段的开头（断行处的光标该出现在下一行行首，与 Word 一致），
        // 所以末尾这一个先攒着，扫完没有更好的再用
        if (pos.offset === end) tail ??= rect;
        else return rect;
      }
      return tail;
    },

    compare,
  };
}

// ── 摊平 ──────────────────────────────────────────────────────────────────────
// 这一段的几何必须与 `@uw/render-dom` 的 paint.ts 一字不差：画在哪儿、点在哪儿，
// 说的是同一件事。两边各写一遍是因为依赖方向单向（渲染层依赖布局，反过来不行），
// 而摊平这件事布局层自己就该会做 —— 命中测试不该逼着调用方先渲染一遍。

function collectFrame(frame: PlacedHeaderFooter, page: number, out: IndexedLine[]): void {
  for (const block of frame.blocks) {
    collectBlock(block, frame.x, frame.y, page, false, frame.kind, out);
  }
}

function collectBlock(
  block: PlacedBlock,
  ox: Twips,
  oy: Twips,
  page: number,
  repeated: boolean,
  frame: 'header' | 'footer' | undefined,
  out: IndexedLine[],
): void {
  if (block.kind === 'paragraph') {
    for (const placed of block.lines) {
      push(out, page, ox, oy + placed.y, placed.line, repeated, frame);
    }
    return;
  }
  for (const placed of block.rows) {
    // 行的上边那条格线占着高度，格内内容从格线**以内**起排（`RowLayout.gridAbove`，实测）
    const inner = oy + placed.y + placed.row.gridAbove;
    const innerHeight = placed.height - placed.row.gridAbove;
    const rowRepeated = repeated || placed.repeated === true;
    for (const cell of placed.row.cells) {
      // `vMerge="continue"` 的格子不画内容（上面那个 restart 撑着它），也就没有行
      if (cell.vMerge === 'continue') continue;
      collectCell(cell, ox + block.x, inner, innerHeight, page, rowRepeated, frame, out);
    }
  }
}

function collectCell(
  cell: CellLayout,
  tableX: Twips,
  rowTop: Twips,
  rowHeight: Twips,
  page: number,
  repeated: boolean,
  frame: 'header' | 'footer' | undefined,
  out: IndexedLine[],
): void {
  const inner = contentHeightOf(cell.blocks);
  const avail = rowHeight - cell.paddingTop - cell.paddingBottom;
  let y = rowTop + cell.paddingTop;
  if (cell.verticalAlign === 'center') y += Math.max(0, (avail - inner) / 2);
  else if (cell.verticalAlign === 'bottom') y += Math.max(0, avail - inner);
  collectStack(cell.blocks, tableX + cell.x + cell.paddingLeft, y, page, repeated, frame, out);
}

/** 一摞块（段落 / 嵌套表格）从 `y0` 往下排。累加规则与 `contentHeightOf` 必须一致 */
function collectStack(
  blocks: readonly BlockLayout[],
  x0: Twips,
  y0: Twips,
  page: number,
  repeated: boolean,
  frame: 'header' | 'footer' | undefined,
  out: IndexedLine[],
): void {
  let y = y0;
  for (const b of blocks) {
    if (b.kind === 'table') {
      for (const row of b.layout.rows) {
        collectNestedRow(row, x0 + b.layout.x, y, page, repeated, frame, out);
        y += row.height;
      }
      // 表底那条格线不属于任何一行（`TableLayout.gridBelow`），漏了它后面的块会整体上移
      y += b.layout.gridBelow;
      continue;
    }
    y += b.layout.spaceBefore;
    for (const line of b.layout.lines) {
      push(out, page, x0, y, line, repeated, frame);
      y += line.height;
    }
    y += b.layout.spaceAfter;
  }
}

function collectNestedRow(
  row: RowLayout,
  x0: Twips,
  y: Twips,
  page: number,
  repeated: boolean,
  frame: 'header' | 'footer' | undefined,
  out: IndexedLine[],
): void {
  for (const cell of row.cells) {
    if (cell.vMerge === 'continue') continue;
    collectCell(cell, x0, y + row.gridAbove, row.height - row.gridAbove, page, repeated, frame, out);
  }
}

function push(
  out: IndexedLine[],
  page: number,
  originX: Twips,
  top: Twips,
  line: LineLayout,
  repeated: boolean,
  frame: 'header' | 'footer' | undefined,
): void {
  out.push({ page, originX, top, line, repeated, ...(frame === undefined ? {} : { frame }) });
}

// ── 行内 ──────────────────────────────────────────────────────────────────────

/**
 * 行内的命中：取离 `x` 最近的**字缝**。
 *
 * 逐字缝比而不是「先找片段再找字」，是因为片段之间可能有空隙（制表位、悬挂出去的标点、
 * 两端对齐拉开的空白），先定片段会把空隙里的点判给错的那一边。一行几十个字，
 * 全扫一遍的代价可以忽略。
 */
function caretIn(l: IndexedLine, x: Twips): DocPosition | undefined {
  let best: DocPosition | undefined;
  let bestD = Number.POSITIVE_INFINITY;
  for (const frag of l.line.fragments) {
    // 编号与域结果没有源位置，落在它们身上的点归旁边真实的字（见 `LayoutIndex.positionAt`）
    if (frag.offset < 0) continue;
    let u16 = 0;
    let g = 0;
    for (const ch of frag.text) {
      const left = frag.glyphX[g] as Twips;
      const d = Math.abs(x - left);
      // 打平取**后**一个：一个字的右边与下一个字的左边是同一条缝，
      // 报成「下一个字之前」才与 `caretRect` 一致（断行处的光标在下一行行首）
      if (d <= bestD) {
        bestD = d;
        best = { nodeId: frag.runId, contentIndex: frag.contentIndex, offset: frag.offset + u16 };
      }
      u16 += ch.length;
      g++;
    }
    // 片段末尾那一缝：`glyphX` 只有每个字的**左**边，最后一个字的右边要用片段宽度补
    const right = frag.x + frag.width;
    const d = Math.abs(x - right);
    if (d < bestD) {
      bestD = d;
      best = { nodeId: frag.runId, contentIndex: frag.contentIndex, offset: frag.offset + u16 };
    }
  }
  return best;
}

/**
 * 片段与 range 的交集，返回**片段自己那套** UTF-16 偏移的区间 `[from, to)`
 * （与 `LineFragment.offset` 同一套，直接喂给 `fragXRange`）；不相交给 `undefined`。
 *
 * 编号与域结果（`offset < 0`）**只有被整个包住时才算**：它们的字符在文件里没有位置，
 * 说不出「第 3 个字在不在范围里」。整段包住时把它整个收进来 —— 选一整段再复制，
 * 域算出来的页码该跟着走（编号则由它自己的合成 runId 挡在范围外，见 items.ts）。
 */
function coveredSpan(
  frag: LineFragment,
  order: ReadonlyMap<NodeId, number>,
  so: number,
  s: DocPosition,
  eo: number,
  e: DocPosition,
): [number, number] | undefined {
  const fo = order.get(frag.runId);
  if (fo === undefined) return undefined;
  const len = frag.text.length;
  if (frag.offset < 0) {
    const after = cmpKey(fo, frag.contentIndex, frag.offset, so, s.contentIndex, s.offset) >= 0;
    const before = cmpKey(fo, frag.contentIndex, frag.offset, eo, e.contentIndex, e.offset) < 0;
    return after && before ? [frag.offset, frag.offset + len] : undefined;
  }
  const start = frag.offset;
  const end = frag.offset + len;
  // 片段整个在范围之前 / 之后
  if (cmpKey(fo, frag.contentIndex, end, so, s.contentIndex, s.offset) <= 0) return undefined;
  if (cmpKey(fo, frag.contentIndex, start, eo, e.contentIndex, e.offset) >= 0) return undefined;
  const from = fo === so && frag.contentIndex === s.contentIndex ? Math.max(start, s.offset) : start;
  const to = fo === eo && frag.contentIndex === e.contentIndex ? Math.min(end, e.offset) : end;
  return to > from ? [from, to] : undefined;
}

/**
 * 片段内 UTF-16 区间 `[from, to)` 的横向范围（相对容器左边）。
 *
 * 右边界取的是**下一个字的左边**而不是这个字的墨迹右边：装饰要连成一片，
 * 中间不能因为字距漏出一条白缝。`to === from` 时给的是一条缝（光标）。
 */
function fragXRange(frag: LineFragment, from: number, to: number): [Twips, Twips] {
  const right = frag.x + frag.width;
  let u16 = 0;
  let g = 0;
  let x1: Twips | undefined;
  let x2: Twips = right;
  for (const ch of frag.text) {
    const abs = frag.offset + u16;
    if (x1 === undefined && abs >= from) x1 = frag.glyphX[g] as Twips;
    if (abs >= to) {
      x2 = frag.glyphX[g] as Twips;
      break;
    }
    u16 += ch.length;
    g++;
  }
  return [x1 ?? right, Math.max(x1 ?? right, x2)];
}

// ── 小工具 ────────────────────────────────────────────────────────────────────

/** 点到区间的距离，区间内是 0 */
function gap(v: number, lo: number, hi: number): number {
  if (v < lo) return lo - v;
  if (v > hi) return v - hi;
  return 0;
}

function cmpKey(ao: number, aci: number, aoff: number, bo: number, bci: number, boff: number): number {
  if (ao !== bo) return ao - bo;
  if (aci !== bci) return aci - bci;
  return aoff - boff;
}

function sign(n: number): -1 | 0 | 1 {
  return n < 0 ? -1 : n > 0 ? 1 : 0;
}
