/**
 * 段落属性与字符属性的类型。
 *
 * 分两套：
 * - `ParaProps` / `RunProps` —— **部分**属性，每一项都可能缺席。级联的每一层
 *   （docDefaults、样式、直接格式）各产出一份，缺席表示「这层没意见」
 * - `ResolvedParaProps` / `ResolvedRunProps` —— 级联的**结果**，字段全部有值。
 *   布局层只吃这个，不必再判断 undefined，也不必再认识 XML
 *
 * 两个纪律：
 * 1. **单位在解析处就转成 twips**（原则 1.3），别把半磅、1/240 行这些原始刻度
 *    带进布局层。唯一的例外见 `LineSpacing`，那里的刻度取决于另一个字段
 * 2. 这些类型必须可结构化克隆 —— 它们是 `LayoutInput` 的组成部分（原则 1.1）
 */
import type { Twips } from '@uw/core';

// ── 字符属性 ──────────────────────────────────────────────────────────────────

/**
 * `w:rFonts` 的四个脚本桶。
 *
 * 这**不是**「给这个 run 指定一款字体」，而是同时挂四款：引擎要逐字符判断
 * 该字符属于哪个桶，再取对应字体的度量。「汉字用宋体、数字英文用 Times New Roman」
 * 就是这么实现的 —— 也因此**一个 run 内可能横跨多款字体**。
 *
 * 分桶发生在 `@uw/fonts`（架构 §3.1），model 只负责把四个名字解析出来（含主题字体）。
 */
export interface RunFonts {
  ascii?: string;
  hAnsi?: string;
  eastAsia?: string;
  cs?: string;
  /**
   * 歧义字符（全角标点、部分符号）该归哪个桶的决断依据。
   * `eastAsia` 表示「按东亚处理」—— 中文文档里 Word 几乎总是写这个。
   */
  hint?: 'default' | 'eastAsia' | 'cs';
}

/** 主题字体引用（`w:asciiTheme="minorHAnsi"` 之类），解析后会被换成真实字体名 */
export interface RunFontThemes {
  ascii?: string;
  hAnsi?: string;
  eastAsia?: string;
  cs?: string;
}

export type VerticalAlign = 'baseline' | 'superscript' | 'subscript';

export interface RunProps {
  /** 字符样式 id（`w:rStyle`） */
  styleId?: string;
  fonts?: RunFonts;
  /** 主题字体引用，级联完再解析成真实字体名 */
  fontThemes?: RunFontThemes;
  bold?: boolean;
  boldCs?: boolean;
  italic?: boolean;
  italicCs?: boolean;
  caps?: boolean;
  smallCaps?: boolean;
  strike?: boolean;
  doubleStrike?: boolean;
  /** `w:vanish`：隐藏文字。**不参与排版**，不是「画成透明」 */
  hidden?: boolean;
  /** 字号。`w:sz` 是半磅，这里已转 twips */
  size?: Twips;
  sizeCs?: Twips;
  /** `w:u w:val`，原样保留（渲染层的事，不影响排版） */
  underline?: string;
  /** 六位十六进制或 `auto`。主题色未解析，见文件末尾说明 */
  color?: string;
  themeColor?: string;
  vertAlign?: VerticalAlign;
  /** `w:spacing`：字符间距，本来就是 twips */
  charSpacing?: Twips;
  /** `w:w`：字符横向缩放，百分比（100 = 不缩放）。**直接改字符宽度**，排版必须认 */
  scale?: number;
  /** `w:position`：基线升降，`w:sz` 同为半磅，已转 twips。正数升 */
  position?: Twips;
  /** `w:kern`：字号小于此值时不做字距调整。半磅 → twips */
  kerning?: Twips;
  /** `w:snapToGrid`：这个 run 是否吸附行网格。公文里关掉它会破坏网格对齐 */
  snapToGrid?: boolean;
  /** `w:lang w:eastAsia`，决定主题字体里东亚脚本的回退目标 */
  langEastAsia?: string;
}

/** 级联结果：字段全有值，布局层不必再判 undefined */
export interface ResolvedRunProps {
  /** 四个桶都已解析成真实字体名（主题引用已展开）。空串表示这个桶没有指定 */
  fonts: Required<Omit<RunFonts, 'hint'>> & { hint: NonNullable<RunFonts['hint']> };
  bold: boolean;
  boldCs: boolean;
  italic: boolean;
  italicCs: boolean;
  caps: boolean;
  smallCaps: boolean;
  strike: boolean;
  doubleStrike: boolean;
  hidden: boolean;
  size: Twips;
  sizeCs: Twips;
  underline: string;
  color: string;
  vertAlign: VerticalAlign;
  charSpacing: Twips;
  scale: number;
  position: Twips;
  kerning: Twips;
  snapToGrid: boolean;
  langEastAsia: string;
}

// ── 段落属性 ──────────────────────────────────────────────────────────────────

export type Justification = 'left' | 'center' | 'right' | 'both' | 'distribute';

/**
 * 缩进。
 *
 * `*Chars` 系列的单位是 **1/100 字符**，且「一个字符」= 当前字号的全角宽 ——
 * 所以它**不能**在这里换成 twips，字号是级联之后才知道的。
 * 「首行缩进 2 字符」= `firstLineChars: 200`，公文里几乎每段都有。
 *
 * 两者同时存在时（Word 为兼容旧版会同时写 `w:firstLineChars` 和 `w:firstLine`），
 * **字符单位优先** —— 这是 Word 的实际行为，选错会让每段首行差几十 twips。
 */
export interface Indent {
  left?: Twips;
  right?: Twips;
  firstLine?: Twips;
  hanging?: Twips;
  leftChars?: number;
  rightChars?: number;
  firstLineChars?: number;
  hangingChars?: number;
}

/**
 * `w:line` 的刻度取决于 `w:lineRule`，这是 OOXML 里最容易搞错的一处：
 * - `auto`（多倍行距）：`line` 的单位是 **1/240 行**，240 = 单倍
 * - `exact`（固定值）/ `atLeast`（最小值）：`line` 的单位是 **twips**
 *
 * 所以这里**故意不转单位**，原样带着 rule 一起交给布局层，
 * 由它按 rule 分支解释。在这里转会丢信息。
 */
export type LineRule = 'auto' | 'exact' | 'atLeast';

export interface ParagraphSpacing {
  before?: Twips;
  after?: Twips;
  /** 段前/段后的「行」单位版本，1/100 行 */
  beforeLines?: number;
  afterLines?: number;
  line?: number;
  lineRule?: LineRule;
  /** `w:beforeAutospacing`：段前自动间距（HTML 段落语义），开着时 before 被忽略 */
  beforeAutospacing?: boolean;
  afterAutospacing?: boolean;
}

/**
 * 一个制表位（`w:tabs/w:tab`）。
 *
 * `pos` 的原点是**版心左边**（不含左缩进），这一点和 `w:ind` 不同 —— 记错的话
 * 公文里靠制表位对齐的「签发人」一栏会整体偏掉一个缩进量。
 *
 * `clear` 不是一种对齐方式，而是「把继承来的、位于 `pos` 的那个制表位删掉」，
 * 只在级联时有意义；级联结果里不会再出现（见 cascade.ts 的 applyTabs）。
 */
export interface TabStop {
  pos: Twips;
  alignment: 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'clear';
  /** 前导符（目录里那排点就是它）。渲染要用，排版上不影响坐标 */
  leader: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
}

/** `w:numPr`：列表编号引用 */
export interface NumberingRef {
  numId?: number;
  level?: number;
}

/**
 * 一个段落最终显示出来的编号，计数器跑完才有。
 *
 * 它**不是**段落的一个 run：编号文字不在 `document.xml` 里，改不了、选不中、
 * 复制不出来，字符属性也另有来源（`w:lvl/w:rPr`，不是正文那份）。放在段落属性上
 * 而不是伪造一个 run，正是为了让编辑期不会误把它当成可编辑内容。
 */
export interface NumberLabel {
  /** 展开后的编号文字。`numFmt=none` 时是空串（**仍然占位**，不等于没有编号） */
  text: string;
  /** 本段在本级的计数值。交叉引用要的是这个数 */
  value: number;
  /** 编号与正文之间的分隔。缺省是制表位，不是空格 */
  suffix: 'tab' | 'space' | 'nothing';
  /** 编号自己在编号区里的对齐（`w:lvlJc`），右对齐用于「 9.」「10.」对齐个位 */
  justification: Justification;
  /** 编号文字自己的字符属性：项目符号的字体就在这儿（Symbol / Wingdings） */
  runProps: ResolvedRunProps;
}

export interface ParaProps {
  /** 段落样式 id（`w:pStyle`） */
  styleId?: string;
  justification?: Justification;
  indent?: Indent;
  spacing?: ParagraphSpacing;
  /** 本层声明的制表位。级联是**逐个合并**的，不是整块替换，见 cascade.ts */
  tabs?: TabStop[];
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  /** 孤行寡行控制，Word 默认开 */
  widowControl?: boolean;
  outlineLevel?: number;
  numbering?: NumberingRef;
  /** 段落是否吸附行网格。公文的「每页 22 行」靠它 */
  snapToGrid?: boolean;
  /** 中西文之间自动加 1/4 em 间距（实测，见 @uw/layout 的 WIDTH_RULES），默认开 */
  autoSpaceDE?: boolean;
  /** 中文与数字之间的自动间距，默认开 */
  autoSpaceDN?: boolean;
  /** 允许标点溢出版心，默认开 —— 行尾句号可以吐出边界 */
  overflowPunct?: boolean;
  /**
   * 段落标记（¶）自己的字符属性。
   * 不是摆设：空段落的行高、以及段末换行符的度量都取它。
   */
  markRunProps?: RunProps;
}

export interface ResolvedParaProps {
  styleId: string;
  justification: Justification;
  indent: Required<Indent>;
  spacing: Required<Omit<ParagraphSpacing, 'lineRule'>> & { lineRule: LineRule };
  /** 显式制表位，按 `pos` 升序、已去掉 `clear`。空数组表示只走 `defaultTabStop` */
  tabs: TabStop[];
  keepNext: boolean;
  keepLines: boolean;
  pageBreakBefore: boolean;
  widowControl: boolean;
  outlineLevel: number;
  /**
   * 编号引用与算好的编号。`label` 只在**按文档顺序**级联（`resolveBody`）时才有 ——
   * 单独调 `resolveParaProps` 拿不到计数器，那时只有 numId / level。
   */
  numbering: { numId: number; level: number; label?: NumberLabel };
  snapToGrid: boolean;
  autoSpaceDE: boolean;
  autoSpaceDN: boolean;
  overflowPunct: boolean;
  markRunProps: ResolvedRunProps;
}

// ── 关于主题色 ────────────────────────────────────────────────────────────────
// `w:themeColor` + `w:themeTint` / `w:themeShade` 没有解析成最终 RGB，只是原样带着。
// 理由：颜色对**排版坐标零影响**，而 tint/shade 的混色算法要连 clrScheme 一起实现。
// Phase 2 的判据是「与真值差多少 pt」，颜色一分都不差 —— 等真要渲染好看了再补，
// 位置就在 theme.ts 旁边。
