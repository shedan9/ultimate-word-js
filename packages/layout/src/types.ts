/**
 * 布局的数据形状。
 *
 * 全部**可结构化克隆**（原则 1.1）：没有类实例、没有闭包、没有反向指针 ——
 * 这一条同时买到 Worker 化、golden file 回归、以及将来把断行换成 Rust/WASM 的能力。
 * 单位一律 twips（原则 1.3），出现 px 视为 bug。
 *
 * ⚠️ 这里**没有 y、没有基线、没有页**：东亚行高里那 30% 额外行距在基线上下如何分配
 * 还没标定（见 `@uw/fonts` 的 metrics.ts），行盒装配与分页因此全部停工。
 * 现在做完的是流水线里「行盒之前」的那一段：分桶 → 度量 → 断行 → 行内水平几何。
 * 补完穿刺后，`LineLayout` 加 `baseline` / `y` 即可，其余字段不用动。
 */
import type { Twips } from '@uw/core';
import type { LineMetrics, ScriptKind } from '@uw/fonts';
import type { NodeId } from '@uw/model';

/** 断行与度量的最小单位。一个码点一个 item —— 逐字 x 是中文排版的硬需求 */
export interface CharItem {
  kind: 'char';
  runId: NodeId;
  /** 在 `run.content` 里的下标 + 该片段内的 UTF-16 偏移，命中测试与 `DocPosition` 反查靠它 */
  contentIndex: number;
  offset: number;
  /** 实际参与度量与渲染的码点（`w:caps` 的大写化已经做掉） */
  cp: number;
  font: string;
  /** 已含上下标 / 小型大写的缩放，直接拿去量 */
  fontSize: Twips;
  script: ScriptKind;
  /** 推进宽度，已含 `w:w` 横向缩放与 `w:spacing` 字间距 */
  width: Twips;
  /** 与**前一个** item 之间的中西文自动间距。行首的那一个不生效 */
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
}

export interface TabItem {
  kind: 'tab';
  runId: NodeId;
  contentIndex: number;
  /** 前导符由命中的制表位决定，断行时才知道 */
  fontSize: Twips;
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
  text: string;
  /** 相对**版心左边**（不是行首），twips */
  x: Twips;
  width: Twips;
  /** 每个**码点**一个 x，与 `text` 的码点序一一对应 */
  glyphX: Twips[];
}

export interface LineLayout {
  /** 对应 `LayoutItem[]` 的区间 `[start, end)` */
  start: number;
  end: number;
  /** 行首 x，相对版心左边：缩进 + 对齐产生的偏移都已算进去 */
  x: Twips;
  /** 行内容宽度，**不含**行尾空格与悬挂出边界的标点 */
  width: Twips;
  /** 行高总量。基线在行内哪个位置**尚未标定**，见文件头 */
  height: Twips;
  /** 各字体逐项取 max 之后的行度量，`lineHeight` 是未经行距规则调整的自然值 */
  metrics: LineMetrics;
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
