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
import type { DrawingAnchor, ImageRef, NodeId, ObjectContent, VerticalAlign } from '@uw/model';

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
  /**
   * 所在 run 的字体（按 ASCII 桶取，制表符本身是 U+0009）。
   * 制表位不占字形，但它**参与行高** —— 只有一个制表位的那一行（目录、签发人栏）
   * 行高全靠它。缺了这个字段那里会拿空字体名去问度量器，退到等宽近似、行高整行错，
   * 顺带每份文档都报一条 `font-missing 字体「」`。
   */
  font: string;
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

/**
 * 内嵌对象（图片 / 图表 / 形状）。断行只关心它占多宽多高，画什么是渲染层的事。
 *
 * 视觉信息（`image` / `alt` / `graphic`）从 model 一路带到这里，与 `LineFragment.style`
 * 同一个道理：**渲染层不回头查模型**（包依赖单向，Worker 化之后主线程手上只有
 * `DocumentLayout` 这一份数据）。字节仍然不在这里 —— `image.id` 是把手，
 * 去 `LoadedDocument.images` 查。
 */
export interface ObjectItem {
  kind: 'object';
  runId: NodeId;
  contentIndex: number;
  width: Twips;
  height: Twips;
  gapBefore: Twips;
  objectKind: ObjectContent['objectKind'];
  /**
   * `w:position`（基线升降）—— 对象底边**高于基线**多少 twips，负数是压到基线以下。
   *
   * 实测它对图片照样起作用（`spike-image-01` 的末两行：`w:position` ±6pt 让图整个跟着升降，
   * 且行盒跟着长高 / 长出下伸），所以它不是「只有文字才有」的属性。缺省 0 = 底边坐在基线上。
   */
  raise?: Twips;
  image?: ImageRef;
  alt?: string;
  graphic?: string;
  /**
   * **浮动且不参与文字流**（有 `wp:anchor` 就算：印章、水印、衬于文字下方的红头、
   * 页脚里的文本框）。
   *
   * 在场时外层的 `width` / `height` 是 **0** —— 断行、行宽、行高、两端对齐全都当它不存在，
   * 这正是 Word 的行为；真实外框在这里面。**尺寸分两处不是冗余**：一处回答「占多少地方」，
   * 另一处回答「画多大」，浮动对象的这两个答案不一样。
   * 位置要页面几何，所以算在分页那一步（page.ts 的 `placeFloats`）。
   *
   * **环绕方式不参与这个判断**：它回答的是「文字怎么让开」，不是「它在不在文字流里」。
   * 没做的正是「让开」那一半 —— 方形 / 上下型环绕的对象位置与大小都对，只是文字
   * 不绕着它走（可能压在一起）；紧密型与穿越型是写死的非目标（开发计划 §5）。
   * 原来按 `wrap === 'none'` 判断、其余退化成**内嵌**，那是错的：内嵌意味着它
   * 撑高所在的行，而 Word 里它根本不在那一行上（真实语料的证据见 items.ts）。
   */
  float?: { anchor: DrawingAnchor; width: Twips; height: Twips };
}

/** 一个对象落在行里的位置。`x` 与 `LineFragment.x` 同一套坐标（相对版心左边） */
export interface LineObject {
  runId: NodeId;
  contentIndex: number;
  x: Twips;
  width: Twips;
  height: Twips;
  /**
   * 对象**底边高于基线**多少 twips。渲染时 y = 基线 − 高 − raise。
   *
   * 两样东西加在一起：`w:position`（基线升降），以及「盒高 − 图高」——
   * 坐在基线上的是**盒**，盒高按 1.5pt 四舍五入，图在盒里靠上放（实测，见
   * `line-height.ts` 的 `objectBoxHeight`）。渲染层只该拿到最终那一个数：
   * 它没有行盒，也不该知道量化这回事。
   */
  raise?: Twips;
  objectKind: ObjectContent['objectKind'];
  image?: ImageRef;
  alt?: string;
  graphic?: string;
}

/**
 * 一个浮动对象**在段落里的锚点**。绝对位置要等分页（page.ts 的 `hoistFloats`）——
 * 段落自己不知道它排在第几页的哪个高度上，而 `page` / `margin` 这些参照物都是页面级的。
 *
 * 留在行上而不是就地算掉，是因为 `ParagraphLayout` 要能缓存复用（types.ts 文件头）：
 * 同一段排在第 3 页还是第 4 页，锚点数据一个字都不变。
 */
export interface LineFloat extends LineObject {
  anchor: DrawingAnchor;
  /**
   * `relativeFrom="character"` 参照的那个字的左边缘（相对版心左边，与 `x` 同一套坐标）。
   *
   * 实测参照的是锚点**前一个**字而不是锚点自己（`spike-image-02` 的三级阶梯：锚在第
   * 1 / 5 / 9 个字之后，x 分别落在第 0 / 4 / 8 个字的左边缘上）。所以它必须在**行内**算 ——
   * 到了分页那一层只剩下对象自己的 x，前一个字是谁已经看不见了。
   */
  anchorX?: Twips;
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
  /**
   * 片段第一个字在 run 里的位置：`RunNode.content` 的下标 + 该片段内的 UTF-16 偏移。
   * 编号与域结果没有源位置，两项都是 -1（与 `CharItem` 同一套约定）。
   *
   * **片段因此不跨内容片，也不跨断开的偏移**（`fragmentsOf` 会在那里切开）：
   * 有了这条保证，「片段里第 k 个码点」的位置就是 `offset + 前 k 个码点的 UTF-16 长度`，
   * 一个数就够，不必给每个字形再存一份位置。命中测试与 `DocPosition` 反查靠的正是它 ——
   * 没有它，`LineFragment` 只说得出「这几个字属于哪个 run」，说不出是**哪几个**字。
   */
  contentIndex: number;
  offset: number;
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
  /** 行内的图片 / 图形。绝大多数行没有，所以是可选的（与 `breakAfter` 同理） */
  objects?: LineObject[];
  /**
   * 锚在这一行上、**不参与文字流**的浮动对象。
   *
   * 渲染层**不画它** —— 画的是 `PageLayout.floats`（分页时按页面几何算出绝对坐标的那一份）。
   * 两处都留着不是冗余：这一份是**输入**（可缓存、与页无关），那一份是**结果**。
   */
  floats?: LineFloat[];
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
