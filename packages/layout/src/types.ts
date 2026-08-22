/**
 * 布局的数据形状。
 *
 * 全部**可结构化克隆**（原则 1.1）：没有类实例、没有闭包、没有反向指针 ——
 * 这一条同时买到 Worker 化、golden file 回归、以及将来把断行换成 Rust/WASM 的能力。
 * 单位一律 twips（原则 1.3），出现 px 视为 bug。
 *
 * 行盒里已经有 `baseline`（基线穿刺标定完了，见 `@uw/fonts` 的 `baselineOffset`），
 * 但**仍然没有 y、没有页** —— 那是分页的产物，段落自己不知道它排在第几页的哪个高度上。
 * 段落的坐标原点是**它自己的左上角**：`LineLayout.x` 相对版心左边，行的 y 靠把前面所有行的
 * `height` 累加起来，分页时再整段平移。这样一段的布局结果可以缓存、可以复用，改动第一段
 * 不会让第五十段的坐标全部失效。
 */
import type { Twips } from '@uw/core';
import type { ScriptKind } from '@uw/fonts';
import type { NodeId, VerticalAlign } from '@uw/model';

/**
 * 一段文字**位置以外**的一切 —— 渲染层照着画，不必回头去查 `ResolvedRunProps`。
 *
 * 为什么要复制一份而不是让渲染层拿 `runId` 回模型里查：`render-*` 只依赖 `@uw/layout`
 * （包依赖方向严格单向，见架构 §3），让它反查模型等于把 model 提到渲染层的依赖里；
 * 而且 Worker 化之后主线程手上只有 `DocumentLayout` 这一份可结构化克隆的数据，
 * 模型根本不在这一侧。字段少、每个 run 共用同一个对象，复制的代价可以忽略。
 *
 * **只收「画」要用的**：影响宽度的那些（`w:spacing` 字间距、`w:sz` 字号、小型大写的缩放）
 * 早就折进了 `CharItem.width` 与 `fontSize`，再带一遍就有两个真相。
 */
export interface FragmentStyle {
  bold: boolean;
  italic: boolean;
  /** 六位十六进制或 `auto`（= 由渲染层挑，通常是黑）。与 model 一样不解析成 RGB */
  color: string;
  /** `w:u/@w:val` 原样带着（`single` / `double` / `wave` …）；`none` = 不画 */
  underline: string;
  strike: boolean;
  doubleStrike: boolean;
  /**
   * `w:vertAlign`。字号已经按 `VERT_ALIGN_SCALE` 缩过了（宽度要用），
   * **升降量还没有** —— 那是渲染层的事，且至今没有真值，见 render-dom 的 uncalibrated.ts
   */
  vertAlign: VerticalAlign;
  /** `w:position`：基线抬高多少 twips（负数是压低）。不影响宽度，所以只有画的时候用得上 */
  position: Twips;
  /**
   * `w:w` 横向缩放，百分数（100 = 不缩）。宽度早已折进 `CharItem.width`，
   * 但**字形本身**也要跟着扁 —— 那件事只有渲染层做得了，所以这个数得带过去
   */
  scale: number;
}

/** 断行与度量的最小单位。一个码点一个 item —— 逐字 x 是中文排版的硬需求 */
export interface CharItem {
  kind: 'char';
  runId: NodeId;
  /**
   * 在 `run.content` 里的下标 + 该片段内的 UTF-16 偏移，命中测试与 `DocPosition` 反查靠它。
   * 编号文字（`numbering`）没有这个位置，两项都是 -1。
   */
  contentIndex: number;
  offset: number;
  /** 实际参与度量与渲染的码点（`w:caps` 的大写化已经做掉） */
  cp: number;
  font: string;
  /** 已含上下标 / 小型大写的缩放，直接拿去量 */
  fontSize: Twips;
  script: ScriptKind;
  /** 画这个字要用的视觉属性。同一个 run 里的所有 item **共用同一个对象** */
  style: FragmentStyle;
  /** 推进宽度，已含 `w:w` 横向缩放与 `w:spacing` 字间距 */
  width: Twips;
  /**
   * 与**前一个** item 之间的间距。行首的那一个不生效。
   *
   * 可以是**负数**：两个全角标点相邻时固定挤掉半个字（实测，见 `PUNCT_PAIR_COMPRESS_EM`），
   * 记法就是给后一个标点一个负间距 —— Word 导出的 PDF 里也是用负偏移把它往左挪的。
   * 负间距同时意味着「这个标点的空半边已经交出去了」，断行时不该再挤它。
   */
  gapBefore: Twips;
  /** 空白字符：断点在它之后，且行尾的它不计入行宽 */
  space: boolean;
  /** 禁则归属：`noStart` 不能出现在行首，`noEnd` 不能出现在行尾 */
  kinsoku: 'none' | 'noStart' | 'noEnd';
  /** 全角标点，塞不下时可以压掉空着的半边 */
  compressible: boolean;
  /** `w:softHyphen`：平时宽度为 0，只有在此处断行时才显出连字符 */
  softHyphen?: true;
  /** `w:noBreakHyphen`：画成连字符但**不许**在此断行 */
  noBreak?: true;
  /**
   * 列表编号的文字。**它不在 document.xml 里** —— 选不中、删不掉、复制不出来，
   * 命中测试与 `DocPosition` 反查必须跳过它，两端对齐也不许在它内部张开。
   */
  numbering?: true;
  /**
   * 域求值算出来的文字（PAGE / NUMPAGES …，见 `fields.ts`）。
   *
   * 与 `numbering` **不是一回事**：编号连复制都不该带上，域结果是要复制、要能被
   * Ctrl+F 搜到的正文；但它同样**不是文件里那串字符**（文件里存的是上次算出来的旧值），
   * 所以 `contentIndex` / `offset` 一律 -1，拿去反查 `DocPosition` 只会指到别处。
   * 排版行为与普通文字完全相同 —— 两端对齐照样在它内部张开，Word 就是这么排的。
   */
  field?: true;
}

export interface TabItem {
  kind: 'tab';
  runId: NodeId;
  contentIndex: number;
  /** 前导符由命中的制表位决定，断行时才知道 */
  fontSize: Twips;
  /**
   * 编号与正文之间那个制表位（`w:suff="tab"`）。它比普通制表位多一个**隐含停靠点**：
   * 段落左缩进（悬挂缩进落脚的地方），见 linebreak.ts 的 `tabAdvance`。
   */
  numbering?: true;
}

export interface BreakItem {
  kind: 'break';
  runId: NodeId;
  contentIndex: number;
  breakType: 'line' | 'page' | 'column';
}

/** 内嵌对象（图片等）。只占位，内容是渲染层的事 */
export interface ObjectItem {
  kind: 'object';
  runId: NodeId;
  contentIndex: number;
  width: Twips;
  height: Twips;
  gapBefore: Twips;
}

export type LayoutItem = CharItem | TabItem | BreakItem | ObjectItem;

/** 制表位的前导符：从 x1 画到 x2（目录里那排点） */
export interface TabLeader {
  x1: Twips;
  x2: Twips;
  leader: 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
}

/**
 * 渲染的最小单位：一行里「同一个 run、同一款字体、同一个字号」的连续字符。
 *
 * **不是一字一元素**（DOM 会爆），也不是一行一元素（逐字微调就没处放了）。
 * `glyphX` 直接喂给 SVG `<text x="x1 x2 x3 …">`，两端对齐 / 标点挤压 / 中西文间距
 * 造成的逐字偏移全在里面。
 */
export interface LineFragment {
  runId: NodeId;
  font: string;
  fontSize: Twips;
  script: ScriptKind;
  style: FragmentStyle;
  text: string;
  /** 相对**版心左边**（不是行首），twips */
  x: Twips;
  width: Twips;
  /** 每个**码点**一个 x，与 `text` 的码点序一一对应 */
  glyphX: Twips[];
  /**
   * 列表编号。渲染层照画，但**可选文本层不要收它**（复制出来会多出一串「一、」），
   * `runId` 也不指向任何真实节点，见 `CharItem.numbering`。
   */
  numbering?: true;
  /** 域求值的结果文字。可选文本层**要**收它，但它不可编辑，见 `CharItem.field` */
  field?: true;
}

export interface LineLayout {
  /** 对应 `LayoutItem[]` 的区间 `[start, end)` */
  start: number;
  end: number;
  /** 行首 x，相对版心左边：缩进 + 对齐产生的偏移都已算进去 */
  x: Twips;
  /** 行内容宽度，**不含**行尾空格与悬挂出边界的标点 */
  width: Twips;
  /** 行高总量：行距规则与网格吸附之后的值 */
  height: Twips;
  /**
   * 行顶到基线，`0 <= baseline <= height`。渲染时文字的基线 = 行顶 y + 这个值。
   * 「多出来的空间上下均分」那条规则就体现在它与 `natural` 的差上，见 `@uw/fonts`
   * 的 `baselineOffset`。
   */
  baseline: Twips;
  /**
   * 未经行距规则与网格调整的自然行高。留着是为了回归比对能分辨
   * 「行高不对是度量的锅还是规则的锅」—— 两者的修法完全不同。
   */
  natural: Twips;
  fragments: LineFragment[];
  leaders: TabLeader[];
  /** 段落的最后一行 —— 两端对齐不拉伸它 */
  isLast: boolean;
  /** 行尾的硬换行（`w:br`）。分页要看 page / column */
  breakAfter?: 'line' | 'page' | 'column';
}

export interface ParagraphLayout {
  paragraphId: NodeId;
  lines: LineLayout[];
  /** 段前 / 段后间距，`w:*Lines`（1/100 行）已经换算成 twips */
  spaceBefore: Twips;
  spaceAfter: Twips;
  /** 版心可用宽度，回归比对时要知道行是在多宽的框里排的 */
  contentWidth: Twips;
}
