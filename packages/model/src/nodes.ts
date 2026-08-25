/**
 * 正文节点树的类型。
 *
 * 这棵树是**布局的输入**（架构 §3.1 的 `LayoutInput`），所以整棵树必须可结构化克隆
 * （原则 1.1）：没有方法、没有闭包、没有父指针、没有 `XmlElement` 引用。
 * 父指针尤其要忍住 —— 它是 Worker 边界上最常见的破坏者，而且会让 JSON 快照测试变成环。
 *
 * 树按属性类型分两次实例化，靠**一个** `PropSet` 类型参数复用同一套形状：
 * - `Block` / `Body` —— 属性是**直接格式**（`ParaProps` / `RunProps` / …，可缺席）。解析产出这个，
 *   编辑也改这个：级联结果是派生量，存进树里迟早会过期
 * - `ResolvedBlock` / `ResolvedBody` —— 属性已级联完（字段全有值）。这才是交给 `@uw/layout` 的
 *
 * 之所以要两套而不是「树里存直接格式、布局时现算」：`StyleSheet` 带方法（`chainOf`），
 * 不可结构化克隆，因此级联必须发生在 Worker 边界**之前**。见 resolve-body.ts。
 *
 * 打包成一个 `PropSet` 而不是排开成 `<P, R, T, TR, TC>`：属性的**种类**还会增加
 * （图片、脚注…），每加一种就改遍所有签名的设计撑不住。
 */
import type { Twips } from '@uw/core';
import type { ParaProps, ResolvedParaProps, ResolvedRunProps, RunProps } from './props.ts';
import type {
  CellProps,
  ResolvedCellProps,
  ResolvedRowProps,
  ResolvedTableProps,
  RowProps,
  TableProps,
} from './table-props.ts';

/**
 * 一棵树用的一套属性类型。两个实例见下面的 `DirectProps` / `ResolvedProps`。
 *
 * 字段用 `unknown` 而不是具体类型：这个接口只负责「有这么几个槽」，
 * 谁填进去由实例说了算。节点里写 `S['para']` 取槽。
 */
export interface PropSet {
  para: unknown;
  run: unknown;
  table: unknown;
  row: unknown;
  cell: unknown;
}

/** 直接格式那棵树（解析产出、编辑改的那棵） */
export interface DirectProps extends PropSet {
  para: ParaProps;
  run: RunProps;
  table: TableProps;
  row: RowProps;
  cell: CellProps;
}

/** 级联完那棵树（交给布局的那棵） */
export interface ResolvedProps extends PropSet {
  para: ResolvedParaProps;
  run: ResolvedRunProps;
  table: ResolvedTableProps;
  row: ResolvedRowProps;
  cell: ResolvedCellProps;
}

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
  /** `w:drawing` / `w:pict` / `w:object`，见 `ObjectContent` */
  | ObjectContent
  /** `w:fldChar`：域的三个界桩。域的**求值**是后续阶段的事，这里只保留位置 */
  | { kind: 'fieldChar'; charType: 'begin' | 'separate' | 'end' }
  /** `w:instrText`：域代码正文（如 `PAGE`）。不参与排版，但求值要靠它 */
  | { kind: 'fieldInstruction'; text: string };

// ── 内嵌 / 浮动对象 ───────────────────────────────────────────────────────────

/**
 * `w:drawing` / `w:pict` / `w:object` —— 图片与图形。
 *
 * **外框尺寸（`width` / `height`）是必需的**：内嵌图参与行高，缺了整行的基线都会错。
 * 尺寸取的是 `wp:extent`（**显示**尺寸，不是图片自己的像素尺寸）—— 用户在 Word 里
 * 拖小的图，extent 就是拖完的那个数，照着画才与 Word 一致。
 *
 * 图片**字节不在这里**：`image.id` 是把手，去 `LoadedDocument.images` 查。
 * 理由与页眉页脚一样（load.ts）—— 同一张图常被引用多次（每页一个页眉 logo），
 * 挂在节点上要么复制几份字节，要么变成跨节点引用（原则 1.1 挡的正是后者）。
 *
 * 没有 `image` 的对象是**画不出来的图形**（图表 / SmartArt / 形状 / OLE），
 * 它们是非目标（开发计划 §5），渲染层画占位框 —— 但**占位框的尺寸是对的**，
 * 所以周围的文字不会跟着错位。
 */
export interface ObjectContent {
  kind: 'object';
  objectKind: 'drawing' | 'picture' | 'object';
  width: Twips;
  height: Twips;
  /** 图片资源引用。非图片图形（图表 / SmartArt / 纯形状）没有它 */
  image?: ImageRef;
  /** 可选文本（`wp:docPr` 的 `descr` / `title`，退到 `name`）。占位框上要显示它 */
  alt?: string;
  /** `a:graphicData@uri` 的短名（`picture` / `chart` / `diagram` / `wordprocessingShape`…）。
   *  画不出来时靠它说清楚「这是张图表」，而不是笼统的「有个对象」 */
  graphic?: string;
  /** 浮动对象（`wp:anchor`）的定位与环绕。内嵌（`wp:inline`）时缺席 */
  anchor?: DrawingAnchor;
}

/**
 * 图片资源引用。
 *
 * `id` 是**部件前缀 + 关系 id**，与节点 id 同一套前缀（见 parse-body.ts 的 `Ctx.idPrefix`）：
 * 页眉部件里的 `rId1` 与正文里的 `rId1` 是两张不同的图，不带前缀会互相顶掉 ——
 * 与页脚页码画进正文那个坑是同一个成因。
 */
export interface ImageRef {
  /** 在 `LoadedDocument.images` 里查字节用的 key */
  id: string;
  /** 关系 id 原文，回写 docx 要它 */
  relId: string;
  /** `r:link`（外链图片）：字节不在包里，`ImageResource.url` 才是地址 */
  linked?: true;
  /**
   * `a:srcRect` 裁剪：四边**各裁掉的比例**（0–1，已从千分之一百分点换算）。
   * 注意它裁的是图片，不是外框 —— 外框尺寸不变，所以裁剪只改「画哪一块」。
   */
  crop?: { left: number; top: number; right: number; bottom: number };
  /** `a:xfrm@rot`，度（已从 1/60000 度换算）。正值顺时针 */
  rotation?: number;
  flipH?: true;
  flipV?: true;
}

/**
 * 浮动对象的定位与环绕（`wp:anchor`）。
 *
 * **有 anchor 就不参与文字流**，位置一律按下面的 `h` / `v` 算 —— 环绕方式回答的是
 * 「文字怎么让开」，不是「它在不在文字流里」。没做的是「让开」那一半：方形与上下型
 * 环绕的对象位置对、大小对，但文字不绕着它走（紧密型 / 穿越型是开发计划 §5 写死的非目标）。
 */
export interface DrawingAnchor {
  wrap: 'none' | 'square' | 'tight' | 'through' | 'topAndBottom';
  /** 衬于文字下方（`behindDoc="1"`）。false = 浮于文字上方 */
  behindDoc: boolean;
  /** z 序（`relativeHeight`）。同一页上几个浮动对象的叠放顺序 */
  z: number;
  h: AnchorPos;
  v: AnchorPos;
  /** 与正文的间距（EMU 已换算成 twips）。`wrap: 'none'` 用不上，方形环绕做进来时要 */
  dist: { top: Twips; bottom: Twips; left: Twips; right: Twips };
}

/**
 * 一个方向上的定位。`offset` 与 `align` **二选一**（规范里是 choice）：
 * 前者是「距参照物 x twips」，后者是「靠参照物的左/中/右」。
 */
export interface AnchorPos {
  /** 参照物：`page` / `margin` / `column` / `character` / `paragraph` / `line` … */
  relativeFrom: string;
  offset?: Twips;
  align?: string;
}

// ── 节点 ──────────────────────────────────────────────────────────────────────

export interface RunNode<S extends PropSet> {
  kind: 'run';
  id: NodeId;
  props: S['run'];
  content: RunContent[];
  /**
   * 超链接。
   *
   * XML 里超链接是**包着 run 的容器**（`w:hyperlink`），这里把它压平成 run 上的一个标记 ——
   * 段落的子节点因此是一列扁平的 run，断行算法不必递归下钻。
   * 排版上超链接与普通文字毫无区别，只有渲染层需要这几个字段。
   *
   * 三个字段不会同时出现：容器那条路给 `relId`（地址要查 rels），HYPERLINK **域**那条路
   * 把地址字面写在指令里、给 `url`（见 fields.ts）。域那条是**派生**的，只出现在
   * `ResolvedBody` 上 —— 可编辑的那棵树里它仍然只是一串界桩，回写时才对得上原样。
   */
  hyperlink?: { relId?: string; url?: string; anchor?: string };
  /**
   * 外层 `w:fldSimple`（**简单域**）的域代码。
   *
   * 简单域是「begin + instrText + separate + 结果 + end」那一串的压缩写法：整个域一个元素，
   * 域代码在属性里、结果就是子 run。压平的道理与超链接相同，但**必须带 id** ——
   * 挨着的两个 `w:instr="PAGE"` 是两个域，只比指令文字会把它们并成一个。
   */
  fieldSimple?: { id: NodeId; instr: string };
}

export interface ParagraphNode<S extends PropSet> {
  kind: 'paragraph';
  id: NodeId;
  props: S['para'];
  runs: RunNode<S>[];
}

export interface TableNode<S extends PropSet> {
  kind: 'table';
  id: NodeId;
  props: S['table'];
  /**
   * `w:tblGrid`：列宽的基准网格，一列一个值（twips）。
   *
   * 放在**节点**上而不是属性里，因为它不参与样式级联 —— 表格样式定义不了 `w:tblGrid`。
   * 与 `gridSpan` 同类：这是结构，不是格式。空数组表示文件里没写，
   * 那就只能靠 `w:tcW` 反推（见布局层的列宽算法）。
   */
  grid: Twips[];
  rows: TableRowNode<S>[];
}

export interface TableRowNode<S extends PropSet> {
  kind: 'row';
  id: NodeId;
  props: S['row'];
  /**
   * `w:tblPrEx`：**本行专有的表级属性**（边框 / 单元格边距 / 底纹 / `w:tblLook`）。
   * Word 在「从另一张表粘过来一行」时写它，真实公文里很常见。
   *
   * 它是**直接格式**，级联时就吃掉了（`cascade-table.ts` 的 `applyRowExceptions`），
   * 所以只出现在直接格式那棵树上，`ResolvedTable` 上永远是 undefined ——
   * 类型写成部分属性而不是 `S['table']` 正是这个意思。
   * 漏解析的后果是这一行的格线与格内边距沿用整表的，与 Word 差一整条边框宽。
   */
  propsEx?: TableProps;
  cells: TableCellNode<S>[];
}

export interface TableCellNode<S extends PropSet> {
  kind: 'cell';
  id: NodeId;
  props: S['cell'];
  /** 单元格里又是完整的块级内容 —— 表格可以嵌套表格 */
  blocks: BlockNode<S>[];
  /**
   * `w:gridSpan`：横向合并占几列，默认 1。
   * 和 `vMerge` 一样是**结构**不是格式：放进 `props` 的话，表格样式的条件格式
   * 就能把「这格合并了几列」给覆盖掉。
   */
  gridSpan: number;
  /** `w:vMerge`：`restart` 是合并区的第一格，`continue` 的内容不显示（由上格占位） */
  vMerge: 'none' | 'restart' | 'continue';
}

export type BlockNode<S extends PropSet> = ParagraphNode<S> | TableNode<S>;

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

/**
 * `w:sectPr/w:type`：这一节**自己**从哪儿开始（不是「下一节怎么开始」）。
 *
 * 容易搞反的地方：`sectPr` 存在一节的**末尾**，但 `w:type` 说的是**本节相对前一节**
 * 怎么落位（§17.6.22）。按「它管下一节」实现的话，整份文档的分页会整体错开一节。
 * 缺席时按 `nextPage`。`nextColumn` 在单栏文档里等同于 `nextPage`（多栏是非目标）。
 */
export type SectionStart = 'nextPage' | 'continuous' | 'nextColumn' | 'evenPage' | 'oddPage';

export interface SectionProps {
  page: { width: Twips; height: Twips; orientation: 'portrait' | 'landscape' };
  /** 本节从哪儿开始，见 `SectionStart` */
  type: SectionStart;
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
  /**
   * `w:pgNumType w:fmt`：本节页码用哪种数字格式（取值就是 `w:numFmt` 那一套，
   * `upperRoman` / `chineseCounting` …）。缺席按 `decimal`。
   *
   * 它只管**页码**：PAGE 域没写 `\*` 开关时跟着它走（前言用罗马数字、正文用阿拉伯数字
   * 就是靠分节 + 这个字段实现的），页眉里的页码同理。域自己写了 `\*` 的以域为准。
   */
  pageNumFormat?: string;
}

/**
 * 一节。
 *
 * 注意 OOXML 的分节写法很别扭：节属性挂在这一节**最后一个段落**的 `w:pPr/w:sectPr` 上，
 * 最后一节的挂在 `w:body` 末尾。这里已经归一化成「一节 = 属性 + 它管辖的块」，
 * 布局层不必再懂这套约定。
 */
export interface SectionNode<S extends PropSet> {
  id: NodeId;
  props: SectionProps;
  blocks: BlockNode<S>[];
}

export interface DocumentBody<S extends PropSet> {
  sections: SectionNode<S>[];
}

// ── 两次实例化的别名 ──────────────────────────────────────────────────────────

export type Run = RunNode<DirectProps>;
export type Paragraph = ParagraphNode<DirectProps>;
export type Table = TableNode<DirectProps>;
export type TableRow = TableRowNode<DirectProps>;
export type TableCell = TableCellNode<DirectProps>;
export type Block = BlockNode<DirectProps>;
export type Section = SectionNode<DirectProps>;
export type Body = DocumentBody<DirectProps>;

export type ResolvedRun = RunNode<ResolvedProps>;
export type ResolvedParagraph = ParagraphNode<ResolvedProps>;
export type ResolvedTable = TableNode<ResolvedProps>;
export type ResolvedTableRow = TableRowNode<ResolvedProps>;
export type ResolvedTableCell = TableCellNode<ResolvedProps>;
export type ResolvedBlock = BlockNode<ResolvedProps>;
export type ResolvedSection = SectionNode<ResolvedProps>;
export type ResolvedBody = DocumentBody<ResolvedProps>;

// ── 遍历与取文本 ──────────────────────────────────────────────────────────────

/** 深度优先遍历所有块（会下钻进表格单元格） */
export function* walkBlocks<S extends PropSet>(blocks: readonly BlockNode<S>[]): Generator<BlockNode<S>> {
  for (const b of blocks) {
    yield b;
    if (b.kind === 'table') {
      for (const row of b.rows) for (const cell of row.cells) yield* walkBlocks(cell.blocks);
    }
  }
}

/** 深度优先遍历所有段落 */
export function* walkParagraphs<S extends PropSet>(body: DocumentBody<S>): Generator<ParagraphNode<S>> {
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
export function paragraphText<S extends PropSet>(p: ParagraphNode<S>): string {
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
