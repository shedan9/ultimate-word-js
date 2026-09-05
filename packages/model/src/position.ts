/**
 * 模型空间的位置（架构 §4 的第 ① 个坐标空间）。
 *
 * 它**与排版无关**：重排、缩放、换一份页面设置之后依然指着同一个字符，
 * 所以批注、书签、选区都该存这个而不是「第几页第几行」。
 * 布局空间（`LayoutPoint`）与它之间的两次转换归 `@uw/layout` 的 `LayoutIndex`。
 *
 * 放在 model 而不是 layout，是因为它说的是**模型树里的一点**：将来 `doc.rangeOf(node)`
 * / 事务系统在编辑后平移位置，都在这一层。layout 只是恰好也要说这件事。
 */
import type { NodeId } from './nodes.ts';

/**
 * 「哪个 run 的第几个内容片段的第几个 UTF-16 单元」。
 *
 * **三个字段而不是 api.md 原先写的两个**：一个 run 的内容是一列片段
 * （`w:t` / `w:tab` / `w:br` / `w:drawing` …，见 `RunContent`），片段没有自己的 id，
 * 而「run 内的全局字符偏移」要把前面每个片段的长度加起来才算得出 —— 那是模型才有的数据。
 * 布局层必须能**独立**说出位置（Worker 化之后主线程手上只有 `DocumentLayout` 这一份数据，
 * 模型根本不在这一侧），所以位置里带着片段下标。
 *
 * `contentIndex` 是数组下标、会被编辑挪动，这一点与 `nodeId` 的「稳定标识」不同 ——
 * 但它只在**本 run 内**挪，平移由事务系统负责，代价远小于全局字符偏移。
 */
export interface DocPosition {
  nodeId: NodeId;
  /** `RunNode.content` 里的下标 */
  contentIndex: number;
  /** 该片段内的 UTF-16 偏移。指向片段末尾（= 片段长度）表示「这一段之后」 */
  offset: number;
}

/** 半开区间 `[start, end)`。两端的先后由 `LayoutIndex.compare()` 判 */
export interface DocRange {
  start: DocPosition;
  end: DocPosition;
}
