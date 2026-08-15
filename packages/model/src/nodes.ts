/**
 * 正文节点树的类型。
 *
 * 这棵树是**布局的输入**（架构 §3.1 的 `LayoutInput`），所以整棵树必须可结构化克隆
 * （原则 1.1）：没有方法、没有闭包、没有父指针、没有 `XmlElement` 引用。
 * 父指针尤其要忍住 —— 它是 Worker 边界上最常见的破坏者，而且会让 JSON 快照测试变成环。
 *
 * 树按属性类型分两次实例化，靠 `<P, R>` 两个类型参数复用同一套形状：
 * - `Block` / `Body` —— 属性是**直接格式**（`ParaProps` / `RunProps`，可缺席）。解析产出这个，
 *   编辑也改这个：级联结果是派生量，存进树里迟早会过期
 * - `ResolvedBlock` / `ResolvedBody` —— 属性已级联完（字段全有值）。这才是交给 `@uw/layout` 的
 *
 * 之所以要两套而不是「树里存直接格式、布局时现算」：`StyleSheet` 带方法（`chainOf`），
 * 不可结构化克隆，因此级联必须发生在 Worker 边界**之前**。见 resolve-body.ts。
 */
import type { Twips } from '@uw/core';
import type { ParaProps, ResolvedParaProps, ResolvedRunProps, RunProps } from './props.ts';

/**
 * 节点标识。`DocPosition{nodeId, offset}` 靠它在重排后依然有效（架构 §5）。
 *
 * 现在是按解析顺序生成的自增 id（`p0` / `r3` / `t1c2`），因此**同一份文件重复解析结果稳定**，
 * 但**重新解析后不保证与编辑前一致** —— 编辑期（Phase 6）的做法是只增不改：新节点取新号，
 * 老节点的号跟着节点走。所以别把 id 当成「第几段」来用。
 */
export type NodeId = string;

// ── run 内部的内容片段 ────────────────────────────────────────────────────────

/**
 * 一个 run 里的内容片段。
 *
 * 拆成片段而不是「一个 run 一段字符串」，是因为 `w:tab` / `w:br` / `w:drawing` 这些
 * 与文字**同层**出现且**顺序有意义**；合成字符串会丢掉顺序，制表位就没法算了。
 */
export type RunContent =
  | { kind: 'text'; text: string }
  | { kind: 'tab' }
  /** `w:br`。`page` / `column` 会中断行盒，`line` 只换行不结束段落 */
  | { kind: 'break'; breakType: 'line' | 'page' | 'column' }
  /** `w:sym`：指定字体 + 字符码位的符号。字体是**片段自己的**，覆盖 run 的字体桶 */
  | { kind: 'symbol'; font: string; char: string }
  /** `w:noBreakHyphen`：显示为连字符，但**不允许**在此断行 */
  | { kind: 'noBreakHyphen' }
  /** `w:softHyphen`：平时不显示，只有在此处断行时才显示出连字符 */
  | { kind: 'softHyphen' }
  /**
   * `w:drawing` / `w:pict` / `w:object`。
   * 只取外框尺寸 —— 内嵌行的行高要它，图形内容本身是 Phase 7 的事。
   * 尺寸缺失（`w:pict` 常见）时为 0，布局层按「零尺寸占位」处理，不要当成 bug。
   */
  | { kind: 'object'; objectKind: 'drawing' | 'picture' | 'object'; width: Twips; height: Twips }
  /** `w:fldChar`：域的三个界桩。域的**求值**是后续阶段的事，这里只保留位置 */
  | { kind: 'fieldChar'; charType: 'begin' | 'separate' | 'end' }
  /** `w:instrText`：域代码正文（如 `PAGE`）。不参与排版，但求值要靠它 */
  | { kind: 'fieldInstruction'; text: string };

// ── 节点 ──────────────────────────────────────────────────────────────────────

export interface RunNode<R> {
  kind: 'run';
  id: NodeId;
  props: R;
  content: RunContent[];
  /**
   * 外层 `w:hyperlink` 的信息。
   *
   * XML 里超链接是**包着 run 的容器**，这里把它压平成 run 上的一个标记 ——
   * 段落的子节点因此是一列扁平的 run，断行算法不必递归下钻。
   * 排版上超链接与普通文字毫无区别，只有渲染层需要这两个字段。
   */
  hyperlink?: { relId?: string; anchor?: string };
}

export interface ParagraphNode<P, R> {
  kind: 'paragraph';
  id: NodeId;
  props: P;
  runs: RunNode<R>[];
}

/**
 * 表格。Phase 1 只建**结构**，不解析表格属性（宽度、边框、单元格边距都在 Phase 4）。
 *
 * 为什么不干脆跳过：跳过等于把表格里的文字**静默丢掉**，而 `w:tbl` 在公文的版记里很常见。
 * 建了结构至少内容都在、`Body` 是完整的，Phase 4 补属性时不动这棵树的形状。
 */
export interface TableNode<P, R> {
  kind: 'table';
  id: NodeId;
  rows: TableRowNode<P, R>[];
}

export interface TableRowNode<P, R> {
  kind: 'row';
  id: NodeId;
  cells: TableCellNode<P, R>[];
}

export interface TableCellNode<P, R> {
  kind: 'cell';
  id: NodeId;
  /** 单元格里又是完整的块级内容 —— 表格可以嵌套表格 */
  blocks: BlockNode<P, R>[];
  /** `w:gridSpan`：横向合并占几列，默认 1 */
  gridSpan: number;
  /** `w:vMerge`：`restart` 是合并区的第一格，`continue` 的内容不显示（由上格占位） */
  vMerge: 'none' | 'restart' | 'continue';
}

export type BlockNode<P, R> = ParagraphNode<P, R> | TableNode<P, R>;

// ── 分节 ──────────────────────────────────────────────────────────────────────

/**
 * `w:docGrid`：行网格 —— 中文排版的命门。
 *
 * 中文版 Word 的 Normal 模板**默认开着**它，基线会被吸附到 `linePitch` 的整数倍上，
 * 把字体度量的差异整个盖掉（Phase 0 穿刺踩过这个坑）。公文「每页 22 行」就是靠它实现的。
 *
 * `charSpace` 的刻度规范说得含糊，Phase 2 拿真值定下来之前**原样保留不解释**，
 * 免得拿一个猜出来的换算去污染坐标。
 */
export interface DocGrid {
  type: 'default' | 'lines' | 'linesAndChars' | 'snapToChars';
  /** 行间距，twips。`type` 为 `default` 时无意义 */
  linePitch: Twips;
  /** 原样保留，未换算 */
  charSpace: number;
}

/** `w:headerReference` / `w:footerReference`：指向 header/footer 部件的关系 id */
export interface HeaderFooterRef {
  type: 'default' | 'first' | 'even';
  relId: string;
}

export interface SectionProps {
  page: { width: Twips; height: Twips; orientation: 'portrait' | 'landscape' };
  /** `header` / `footer` 是页眉页脚**到纸边**的距离，不是到版心的 */
  margin: {
    top: Twips;
    right: Twips;
    bottom: Twips;
    left: Twips;
    header: Twips;
    footer: Twips;
    gutter: Twips;
  };
  docGrid: DocGrid;
  /** 分栏数。多栏排版是非目标，这里只记着，>1 时布局层应发诊断 */
  columns: number;
  /** `w:titlePg`：首页用不同的页眉页脚 */
  titlePage: boolean;
  headers: HeaderFooterRef[];
  footers: HeaderFooterRef[];
  /** `w:pgNumType w:start`：本节页码起始值 */
  pageNumStart?: number;
}

/**
 * 一节。
 *
 * 注意 OOXML 的分节写法很别扭：节属性挂在这一节**最后一个段落**的 `w:pPr/w:sectPr` 上，
 * 最后一节的挂在 `w:body` 末尾。这里已经归一化成「一节 = 属性 + 它管辖的块」，
 * 布局层不必再懂这套约定。
 */
export interface SectionNode<P, R> {
  id: NodeId;
  props: SectionProps;
  blocks: BlockNode<P, R>[];
}

export interface DocumentBody<P, R> {
  sections: SectionNode<P, R>[];
}

// ── 两次实例化的别名 ──────────────────────────────────────────────────────────

export type Run = RunNode<RunProps>;
export type Paragraph = ParagraphNode<ParaProps, RunProps>;
export type Block = BlockNode<ParaProps, RunProps>;
export type Section = SectionNode<ParaProps, RunProps>;
export type Body = DocumentBody<ParaProps, RunProps>;

export type ResolvedRun = RunNode<ResolvedRunProps>;
export type ResolvedParagraph = ParagraphNode<ResolvedParaProps, ResolvedRunProps>;
export type ResolvedBlock = BlockNode<ResolvedParaProps, ResolvedRunProps>;
export type ResolvedSection = SectionNode<ResolvedParaProps, ResolvedRunProps>;
export type ResolvedBody = DocumentBody<ResolvedParaProps, ResolvedRunProps>;

// ── 遍历与取文本 ──────────────────────────────────────────────────────────────

/** 深度优先遍历所有块（会下钻进表格单元格） */
export function* walkBlocks<P, R>(blocks: readonly BlockNode<P, R>[]): Generator<BlockNode<P, R>> {
  for (const b of blocks) {
    yield b;
    if (b.kind === 'table') {
      for (const row of b.rows) for (const cell of row.cells) yield* walkBlocks(cell.blocks);
    }
  }
}

/** 深度优先遍历所有段落 */
export function* walkParagraphs<P, R>(body: DocumentBody<P, R>): Generator<ParagraphNode<P, R>> {
  for (const section of body.sections) {
    for (const b of walkBlocks(section.blocks)) if (b.kind === 'paragraph') yield b;
  }
}

/**
 * 段落的纯文本。
 *
 * 只给调试与测试用 —— **不要拿它去排版**：它把制表位当成 `\t`、把符号和图形当成空，
 * 这些在真正排版时的宽度完全不是这么算的。
 */
export function paragraphText<P, R>(p: ParagraphNode<P, R>): string {
  let out = '';
  for (const run of p.runs) {
    for (const c of run.content) {
      if (c.kind === 'text') out += c.text;
      else if (c.kind === 'tab') out += '\t';
      else if (c.kind === 'break') out += '\n';
      else if (c.kind === 'symbol') out += c.char;
      else if (c.kind === 'noBreakHyphen') out += '-';
    }
  }
  return out;
}
