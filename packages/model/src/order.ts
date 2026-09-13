/**
 * 模型空间的**文档序** —— `DocPosition` 之间怎么比先后、一个节点占哪一段。
 *
 * `LayoutIndex.compare()` 也答文档序，但它只认**排出来**的 run：空 run、隐藏 run、
 * `vMerge="continue"` 格子里的 run 都排不出一个字形，它一律答 `undefined`。
 * 查找与批注却常常指着这些位置（搜到的字在隐藏 run 里、书签压在空段落上），
 * 所以模型这一侧要有一份**完整**的顺序 —— 树里有的 run 都算。
 *
 * 顺序按**一次遍历**建成 `Map`（run id → 序号），而不是每次比较都走一遍树：
 * `find()` 的每一条结果、`contains()` 的每一次判断都要比，逐次遍历是 O(n) × 结果数。
 * 它**不是流水线的产物**（带 Map 的对象过不了结构化克隆，原则 1.1），
 * 与 `LayoutIndex` 同理，在消费侧现建；编辑期节点增删之后要重建。
 */
import type {
  BlockNode,
  DocumentBody,
  NodeId,
  PropSet,
  RunNode,
  TableCellNode,
  TableRowNode,
} from './nodes.ts';
import { walkParagraphs } from './nodes.ts';
import type { DocPosition, DocRange } from './position.ts';

/** run id → 文档序号。run 与无 run 的空段落均有序号 */
export type RunOrder = ReadonlyMap<NodeId, number>;

export function buildRunOrder<S extends PropSet>(body: DocumentBody<S>): RunOrder {
  const order = new Map<NodeId, number>();
  for (const p of walkParagraphs(body)) {
    if (!p.runs.length) order.set(p.id, order.size);
    for (const run of p.runs) order.set(run.id, order.size);
  }
  return order;
}

/**
 * 文档序比较。任一端的 run 不在这份 body 里时返回 `undefined` ——
 * 页眉页脚的 run（id 带 `rId7:` 前缀）就不在正文的顺序里，硬答一个数会把两份内容搅在一起。
 */
export function compareDocPositions(order: RunOrder, a: DocPosition, b: DocPosition): -1 | 0 | 1 | undefined {
  const oa = order.get(a.nodeId);
  const ob = order.get(b.nodeId);
  if (oa === undefined || ob === undefined) return undefined;
  const d = oa - ob || a.contentIndex - b.contentIndex || a.offset - b.offset;
  return d < 0 ? -1 : d > 0 ? 1 : 0;
}

/** `[start, end)` 包不包住一个位置或另一个 range。任一方不在顺序里答 `false` */
export function rangeContains(order: RunOrder, range: DocRange, other: DocPosition | DocRange): boolean {
  if ('start' in other) {
    return rangeContains(order, range, other.start) && rangeContainsEnd(order, range, other.end);
  }
  const a = compareDocPositions(order, range.start, other);
  const b = compareDocPositions(order, other, range.end);
  return a !== undefined && b !== undefined && a <= 0 && b < 0;
}

/** 末端允许与 range 的末端重合（半开区间的右端本身不算在内） */
function rangeContainsEnd(order: RunOrder, range: DocRange, end: DocPosition): boolean {
  const a = compareDocPositions(order, range.start, end);
  const b = compareDocPositions(order, end, range.end);
  return a !== undefined && b !== undefined && a <= 0 && b <= 0;
}

/** 能问「占哪一段」的节点：块、run、表格的行与格。节（section）不算 —— 它是分组不是内容 */
export type RangeableNode<S extends PropSet> = BlockNode<S> | RunNode<S> | TableRowNode<S> | TableCellNode<S>;

/**
 * 一个节点覆盖的 range：首 run 的开头到末 run 的结尾。
 *
 * 空段落以段落 id 的折叠范围表示，单元格 / 表格的范围也包含首末空段落。
 * 原先只有 run 可定位时这里返回 undefined；输入层接入后必须保留空段落插入点。
 */
export function rangeOfNode<S extends PropSet>(node: RangeableNode<S>): DocRange | undefined {
  if (node.kind === 'paragraph' && !node.runs.length) {
    const position = { nodeId: node.id, contentIndex: 0, offset: 0 };
    return { start: { ...position }, end: { ...position } };
  }
  if (node.kind === 'run') return { start: runStart(node), end: runEnd(node) };
  if (node.kind === 'paragraph')
    return { start: runStart(node.runs[0] as RunNode<S>), end: runEnd(node.runs.at(-1) as RunNode<S>) };
  const children = node.kind === 'table' ? node.rows : node.kind === 'row' ? node.cells : node.blocks;
  const ranges = children.flatMap((child) => {
    const range = rangeOfNode(child);
    return range ? [range] : [];
  });
  const first = ranges[0];
  const last = ranges.at(-1);
  return first && last ? { start: first.start, end: last.end } : undefined;
}

export function runStart<S extends PropSet>(run: RunNode<S>): DocPosition {
  return { nodeId: run.id, contentIndex: 0, offset: 0 };
}

/** run 的结尾：最后一个片段之后。没有片段的 run 落在 `{0, 0}`，与开头重合 */
export function runEnd<S extends PropSet>(run: RunNode<S>): DocPosition {
  const ci = run.content.length - 1;
  if (ci < 0) return { nodeId: run.id, contentIndex: 0, offset: 0 };
  return {
    nodeId: run.id,
    contentIndex: ci,
    offset: contentLength(run.content[ci] as RunNode<S>['content'][number]),
  };
}

/**
 * 片段的「长度」—— 位置能落在 `0..length` 之间。与布局层 `LineFragment` 的口径一致：
 * 文字按 UTF-16 单元数，符号 / 连字符 / 制表位 / 换行 / 对象各占 1，
 * 域界桩与指令占 0（它们不显示、也不可定位）。
 */
export function contentLength(c: RunNode<PropSet>['content'][number]): number {
  switch (c.kind) {
    case 'text':
      return c.text.length;
    case 'fieldChar':
    case 'fieldInstruction':
      return 0;
    // 布局只画第一个码点，但那个码点可能是两个 UTF-16 单元（items.ts 的 symbol 分支）
    case 'symbol':
      return String.fromCodePoint(c.char.codePointAt(0) ?? 0xfffd).length;
    default:
      return 1;
  }
}

/** 节点下的全部 run，文档序（表格按行 → 格 → 块下钻） */
export function* runsOf<S extends PropSet>(node: RangeableNode<S>): Generator<RunNode<S>> {
  switch (node.kind) {
    case 'run':
      yield node;
      return;
    case 'paragraph':
      yield* node.runs;
      return;
    case 'table':
      for (const row of node.rows) yield* runsOf(row);
      return;
    case 'row':
      for (const cell of node.cells) yield* runsOf(cell);
      return;
    case 'cell':
      for (const b of node.blocks) yield* runsOf(b);
      return;
  }
}
