/**
 * 文字修改的可克隆记录。片段删空后保留空 text，run 也不合并，
 * 因此本轮编辑不会挪动 contentIndex 或丢掉样式 / 超链接所属的节点。
 */
import type { NodeId } from './nodes.ts';
import type { DocPosition, DocRange } from './position.ts';

export interface PositionMove {
  from: DocPosition;
  to: DocPosition;
  length: number;
  /** 拆分边界的 before 位置留在前段。 */
  afterOnly?: boolean;
  /**
   * 源区间整个消失（删行）：区间内的位置一律落到 `to`，不按偏移平移 ——
   * 平移会把第 5 个字的位置映射成目标段落里并不存在的第 5 个字。
   */
  collapse?: boolean;
}

export interface TextChange {
  /** 结构操作的片段迁移；空文字字段不代表文字插入。 */
  moves?: PositionMove[];
  inverseMoves?: PositionMove[];
  affectedParagraphIds?: NodeId[];
  paragraphId: NodeId;
  nodeId: NodeId;
  contentIndex: number;
  offset: number;
  deletedText: string;
  insertedText: string;
}

export interface TextChangeSet {
  /** 每项的位置相对上一项应用后的树；undo 返回逆序、反向的修改。 */
  changes: TextChange[];
  paragraphIds: NodeId[];
}

/** 插入点上的位置跟随哪一侧。选区端点与光标可按各自语义选择。 */
export type PositionAffinity = 'before' | 'after';

export function mapTextPosition(
  position: DocPosition,
  changeSet: TextChangeSet,
  affinity: PositionAffinity = 'after',
): DocPosition {
  let out = { ...position };
  for (const c of changeSet.changes) {
    if (c.moves !== undefined) {
      const move = c.moves.find(
        (m) =>
          out.nodeId === m.from.nodeId &&
          out.contentIndex === m.from.contentIndex &&
          out.offset >= m.from.offset &&
          out.offset <= m.from.offset + m.length &&
          (affinity === 'after' || out.offset > m.from.offset || !m.afterOnly),
      );
      if (move !== undefined)
        out = move.collapse
          ? { ...move.to }
          : { ...move.to, offset: move.to.offset + out.offset - move.from.offset };
      continue;
    }
    if (out.nodeId !== c.nodeId || out.contentIndex !== c.contentIndex || out.offset < c.offset) continue;
    const end = c.offset + c.deletedText.length;
    if (out.offset > end) out.offset += c.insertedText.length - c.deletedText.length;
    else out.offset = c.offset + (affinity === 'after' ? c.insertedText.length : 0);
  }
  return out;
}

/** 两端默认排除边界上新插入的文字；折叠范围保持折叠并跟随输入。 */
export function mapTextRange(range: DocRange, changeSet: TextChangeSet): DocRange {
  let start = { ...range.start };
  let end = { ...range.end };
  for (const change of changeSet.changes) {
    const collapsed =
      start.nodeId === end.nodeId && start.contentIndex === end.contentIndex && start.offset === end.offset;
    const step = { changes: [change], paragraphIds: [] };
    start = mapTextPosition(start, step, 'after');
    end = mapTextPosition(end, step, collapsed ? 'after' : 'before');
    // 替换把整个范围吞掉时，两端必须一起落到新文字之后，不能形成反向选区。
    if (start.nodeId === end.nodeId && start.contentIndex === end.contentIndex && start.offset > end.offset) {
      end = { ...start };
    }
  }
  return { start, end };
}

export function invertTextChanges(changes: readonly TextChange[]): TextChange[] {
  return [...changes].reverse().map((c) => ({
    ...c,
    deletedText: c.insertedText,
    insertedText: c.deletedText,
    ...(c.moves === undefined ? {} : { moves: c.inverseMoves ?? [], inverseMoves: c.moves }),
  }));
}
