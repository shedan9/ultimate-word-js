/**
 * 样式级联 —— Phase 1 的正题。
 *
 * 顺序（ECMA-376 §17.7.2，后面的覆盖前面的）：
 *
 * ```
 * 段落属性： docDefaults.pPr → 表格样式层 → 段落样式链(祖先→自己) → 编号 pPr → 直接 pPr
 * 字符属性： docDefaults.rPr → 表格样式层 → 段落样式链.rPr → 字符样式链.rPr → 直接 rPr
 * 编号文字： docDefaults.rPr → 表格样式层 → 段落样式链.rPr → 编号 rPr → 段落标记的 rPr
 * ```
 *
 * 「表格样式层」只在单元格里的段落上有（见 `CascadeContext.tableStyleLayers`），
 * 位置在段落样式**之前** —— 表头行的加粗要能被段落自己的样式盖掉。
 *
 * 三处容易错的地方：
 * 1. **段落样式也带字符属性**，且它排在字符样式**前面**。漏了这一层，
 *    「标题 1 是 24 号字」就没了
 * 2. **合并是逐属性的，不是整块替换**。样式设了 `w:ind w:left`，直接格式设了
 *    `w:ind w:firstLine`，结果两个都在。整块替换会把左缩进吃掉
 * 3. **编号的 pPr 与 rPr 作用域不同**：`w:lvl/w:pPr`（缩进、制表位）作用于**整个段落**，
 *    而 `w:lvl/w:rPr`（§17.9.24）只作用于**编号文字本身**。把编号的 rPr 也铺到正文上，
 *    是「项目符号用 Symbol 字体 → 整段正文都变成 Symbol」这类错误的来源
 */
import { halfPtToTwips } from '@uw/core';
import type { Numbering, NumberingLevel } from './numbering.ts';
import { numberingLevel } from './numbering.ts';
import type { NumberedParagraph, NumberingCounters } from './numbering-counter.ts';
import type {
  Indent,
  Justification,
  NumberingRef,
  NumberLabel,
  ParagraphSpacing,
  ParaProps,
  ResolvedParaProps,
  ResolvedRunProps,
  RunProps,
  TabStop,
} from './props.ts';
import type { DocumentSettings } from './settings.ts';
import type { StyleSheet } from './styles.ts';
import type { TableStyleLayer } from './table-props.ts';
import type { Theme } from './theme.ts';
import { resolveThemeFont } from './theme.ts';

export interface CascadeContext {
  styles: StyleSheet;
  theme: Theme;
  /**
   * `settings.xml`。级联只用到 `themeFontLang.eastAsia` 一项，但**这一项非有不可**：
   * 主题里 `a:ea` 是空串时，东亚字体要按语言回退到 `a:font script="..."`，
   * 而语言的来源就是这里。缺了它只能硬编码 zh-CN，非中文文档会拿到简体中文字体。
   */
  settings: DocumentSettings;
  /**
   * `numbering.xml`。编号那一级自带 pPr（几乎总有缩进）要插进段落级联，
   * 所以它是级联的**输入**，不是解析出来放着看的旁支。缺席时给 `EMPTY_NUMBERING`。
   */
  numbering: Numbering;
  /**
   * 「本段落在某个表格单元格里」时，该单元格命中的表格样式层（见 cascade-table.ts）。
   *
   * 放在上下文里而不是当参数传，是因为它**就是**环境的一部分：进单元格时派生一个带层的
   * ctx，出来就没了。嵌套表格因此天然是「内层的层覆盖外层」—— 内层派生时直接换掉这个
   * 字段，不做叠加（没有真值支持叠加，且嵌套表格在公文里罕见）。
   */
  tableStyleLayers?: readonly TableStyleLayer[];
}

/**
 * 一切属性都没写时的兜底值 —— 来自 Word 在 `docDefaults` 缺失时的行为。
 *
 * 实际文档几乎总有 docDefaults，所以这些值很少真正生效；但它们必须存在，
 * 否则 `Resolved*` 就没法保证「字段全有值」这个对布局层的承诺。
 */
const DEFAULT_SIZE = halfPtToTwips(20); // 10pt
const DEFAULT_LINE = 240; // lineRule=auto 下 240 = 单倍行距

const ZERO_INDENT: Required<Indent> = {
  left: 0,
  right: 0,
  firstLine: 0,
  hanging: 0,
  leftChars: 0,
  rightChars: 0,
  firstLineChars: 0,
  hangingChars: 0,
};

// ── 字符属性 ──────────────────────────────────────────────────────────────────

/** 字体桶的来源：显式字体名，还是主题引用。谁在后面谁说了算，所以要一起记 */
type FontSlot = { kind: 'name'; value: string } | { kind: 'theme'; value: string };
type FontSlots = { ascii?: FontSlot; hAnsi?: FontSlot; eastAsia?: FontSlot; cs?: FontSlot };
const BUCKETS = ['ascii', 'hAnsi', 'eastAsia', 'cs'] as const;

interface RunAccum {
  props: RunProps;
  slots: FontSlots;
  hint: NonNullable<RunProps['fonts']>['hint'] | undefined;
}

function applyRunLevel(acc: RunAccum, level: RunProps): void {
  const p = acc.props;
  // 标量：本层写了就覆盖，没写就保持 —— 「没写」不等于 false，见 xml-values.ts 的 onOff
  if (level.bold !== undefined) p.bold = level.bold;
  if (level.boldCs !== undefined) p.boldCs = level.boldCs;
  if (level.italic !== undefined) p.italic = level.italic;
  if (level.italicCs !== undefined) p.italicCs = level.italicCs;
  if (level.caps !== undefined) p.caps = level.caps;
  if (level.smallCaps !== undefined) p.smallCaps = level.smallCaps;
  if (level.strike !== undefined) p.strike = level.strike;
  if (level.doubleStrike !== undefined) p.doubleStrike = level.doubleStrike;
  if (level.hidden !== undefined) p.hidden = level.hidden;
  if (level.snapToGrid !== undefined) p.snapToGrid = level.snapToGrid;
  if (level.size !== undefined) p.size = level.size;
  if (level.sizeCs !== undefined) p.sizeCs = level.sizeCs;
  if (level.underline !== undefined) p.underline = level.underline;
  if (level.color !== undefined) p.color = level.color;
  if (level.vertAlign !== undefined) p.vertAlign = level.vertAlign;
  if (level.charSpacing !== undefined) p.charSpacing = level.charSpacing;
  if (level.scale !== undefined) p.scale = level.scale;
  if (level.position !== undefined) p.position = level.position;
  if (level.kerning !== undefined) p.kerning = level.kerning;
  if (level.langEastAsia !== undefined) p.langEastAsia = level.langEastAsia;

  // 字体：逐桶覆盖。同一层里显式字体名压过主题引用
  for (const b of BUCKETS) {
    const name = level.fonts?.[b];
    const themeRef = level.fontThemes?.[b];
    if (name !== undefined && name !== '') acc.slots[b] = { kind: 'name', value: name };
    else if (themeRef !== undefined && themeRef !== '') acc.slots[b] = { kind: 'theme', value: themeRef };
  }
  if (level.fonts?.hint !== undefined) acc.hint = level.fonts.hint;
}

function finishRun(acc: RunAccum, theme: Theme, settings: DocumentSettings): ResolvedRunProps {
  const p = acc.props;
  // 优先级：run 自己的 w:lang w:eastAsia → 文档的 themeFontLang → 兜底 zh-CN。
  // 兜底值只在文件两处都没写时生效，此时按这个库的定位（中文公文）猜简体中文
  const lang = p.langEastAsia ?? emptyToUndefined(settings.themeFontLang.eastAsia) ?? 'zh-CN';
  const fontOf = (slot: FontSlot | undefined): string => {
    if (slot === undefined) return '';
    return slot.kind === 'name' ? slot.value : resolveThemeFont(theme, slot.value, lang);
  };
  const size = p.size ?? DEFAULT_SIZE;
  return {
    fonts: {
      ascii: fontOf(acc.slots.ascii),
      hAnsi: fontOf(acc.slots.hAnsi),
      eastAsia: fontOf(acc.slots.eastAsia),
      cs: fontOf(acc.slots.cs),
      hint: acc.hint ?? 'default',
    },
    bold: p.bold ?? false,
    boldCs: p.boldCs ?? false,
    italic: p.italic ?? false,
    italicCs: p.italicCs ?? false,
    caps: p.caps ?? false,
    smallCaps: p.smallCaps ?? false,
    strike: p.strike ?? false,
    doubleStrike: p.doubleStrike ?? false,
    hidden: p.hidden ?? false,
    size,
    // szCs 没写时跟随 sz —— Word 就是这么退的，不是退到 10pt
    sizeCs: p.sizeCs ?? size,
    underline: p.underline ?? 'none',
    color: p.color ?? 'auto',
    vertAlign: p.vertAlign ?? 'baseline',
    charSpacing: p.charSpacing ?? 0,
    scale: p.scale ?? 100,
    position: p.position ?? 0,
    kerning: p.kerning ?? 0,
    snapToGrid: p.snapToGrid ?? true,
    langEastAsia: lang,
  };
}

/**
 * 解析一个 run 的最终字符属性。
 *
 * `paraProps` 传的是**段落的直接属性**（用来找段落样式），不是解析后的 —— 因为
 * 段落样式链上的 rPr 要参与字符级联，而那是解析后的 ParaProps 里已经丢掉的信息。
 *
 * `numberingRunProps` **只有算编号文字时才传**：`w:lvl/w:rPr` 的作用域是编号本身，
 * 不是段落正文（见文件头第 3 条）。给正文 run 传它等于让整段跟着项目符号变字体。
 */
export function resolveRunProps(
  ctx: CascadeContext,
  paraProps: ParaProps | undefined,
  direct: RunProps | undefined,
  numberingRunProps?: RunProps,
): ResolvedRunProps {
  const acc: RunAccum = { props: {}, slots: {}, hint: undefined };

  applyRunLevel(acc, ctx.styles.defaults.runProps);
  for (const l of ctx.tableStyleLayers ?? []) applyRunLevel(acc, l.runProps);
  for (const s of ctx.styles.chainOf(paragraphStyleId(ctx, paraProps))) applyRunLevel(acc, s.runProps);
  if (numberingRunProps !== undefined) applyRunLevel(acc, numberingRunProps);
  // 字符样式链排在段落样式之后 —— 字符样式是「更局部」的那一个
  for (const s of ctx.styles.chainOf(direct?.styleId)) applyRunLevel(acc, s.runProps);
  if (direct !== undefined) applyRunLevel(acc, direct);

  return finishRun(acc, ctx.theme, ctx.settings);
}

function emptyToUndefined(s: string): string | undefined {
  return s === '' ? undefined : s;
}

// ── 段落属性 ──────────────────────────────────────────────────────────────────

interface ParaAccum {
  props: ParaProps;
  indent: Indent;
  spacing: ParagraphSpacing;
  numbering: NumberingRef;
  /** 制表位按 `pos` 索引：同一个位置只能有一个，后来的层覆盖前面的 */
  tabs: Map<number, TabStop>;
}

/**
 * 制表位的合并 —— 它不像别的属性那样「整块替换」，而是**逐个位置**合并：
 * 样式定了 1440 处一个右对齐制表位，段落又定了 720 处一个左对齐的，两个都在。
 *
 * `w:val="clear"` 是唯一的删除手段：它把继承来的、位于同一 `pos` 的那个删掉，
 * 自己不留下任何东西。漏了这一条，Word 里被清掉的制表位会在我们这儿复活，
 * 那一行的文字会停在错误的横坐标上。
 */
function applyTabs(acc: ParaAccum, tabs: readonly TabStop[]): void {
  for (const t of tabs) {
    if (t.alignment === 'clear') acc.tabs.delete(t.pos);
    else acc.tabs.set(t.pos, { ...t });
  }
}

function applyParaLevel(acc: ParaAccum, level: ParaProps): void {
  const p = acc.props;
  if (level.justification !== undefined) p.justification = level.justification;
  if (level.keepNext !== undefined) p.keepNext = level.keepNext;
  if (level.keepLines !== undefined) p.keepLines = level.keepLines;
  if (level.pageBreakBefore !== undefined) p.pageBreakBefore = level.pageBreakBefore;
  if (level.widowControl !== undefined) p.widowControl = level.widowControl;
  if (level.snapToGrid !== undefined) p.snapToGrid = level.snapToGrid;
  if (level.autoSpaceDE !== undefined) p.autoSpaceDE = level.autoSpaceDE;
  if (level.autoSpaceDN !== undefined) p.autoSpaceDN = level.autoSpaceDN;
  if (level.overflowPunct !== undefined) p.overflowPunct = level.overflowPunct;
  if (level.outlineLevel !== undefined) p.outlineLevel = level.outlineLevel;

  // 嵌套的三块逐属性合并，不整块替换
  if (level.indent !== undefined) Object.assign(acc.indent, definedOnly(level.indent));
  if (level.spacing !== undefined) Object.assign(acc.spacing, definedOnly(level.spacing));
  if (level.numbering !== undefined) Object.assign(acc.numbering, definedOnly(level.numbering));
  if (level.tabs !== undefined) applyTabs(acc, level.tabs);
}

/** `Object.assign` 会把显式的 undefined 也拷过去，先滤掉 */
function definedOnly<T extends object>(o: T): Partial<T> {
  const out: Partial<T> = {};
  for (const [k, v] of Object.entries(o)) {
    if (v !== undefined) out[k as keyof T] = v as T[keyof T];
  }
  return out;
}

/** 没写 `w:pStyle` 时，用样式表里标了 `w:default="1"` 的那个段落样式 */
function paragraphStyleId(ctx: CascadeContext, direct: ParaProps | undefined): string {
  return direct?.styleId ?? ctx.styles.defaultParagraphStyleId();
}

/**
 * 编号级自带的 pPr 要**去掉三样**再往段落上铺。
 *
 * - `w:pStyle`：`w:lvl/w:pStyle` 是**反向**关系（「用了这个样式的段落自动获得本编号」），
 *   不是「本段改用这个样式」。照搬会让所有带编号的段落被换成列表样式
 * - `w:numPr`：编号级里再指一个编号就成了自指，级联会绕回来
 * - 段落标记的 rPr：编号文字的字符属性走 `w:lvl/w:rPr` 那条路，不是这条
 */
function numberingParaLayer(p: ParaProps): ParaProps {
  const layer: ParaProps = { ...p };
  delete layer.styleId;
  delete layer.numbering;
  delete layer.markRunProps;
  return layer;
}

/**
 * 段落属性级联。
 *
 * `counters` **有副作用**：传了它就会推进编号计数器，因此每个段落只能调一次、
 * 且必须按文档顺序调 —— 实际只有 `resolveBody()` 该传。不传时编号级的 pPr
 * （缩进那些）照样生效，只是没有 `label`：编号是「第几」要靠前文才知道。
 */
export function resolveParaProps(
  ctx: CascadeContext,
  direct: ParaProps | undefined,
  counters?: NumberingCounters,
): ResolvedParaProps {
  const styleId = paragraphStyleId(ctx, direct);
  const chain = ctx.styles.chainOf(styleId);
  // 先只把 numPr 那一项级联出来：编号级的 pPr 要插进下面这条链，而它插在哪一级
  // 又取决于这条链算出来的 numId —— 鸡生蛋只能拆成两步。这一步很便宜（只读一个字段）
  const ref = numberingRefOf(ctx, chain, direct);
  const numbered = numberingOf(ctx, ref, counters);

  const acc: ParaAccum = { props: {}, indent: {}, spacing: {}, numbering: {}, tabs: new Map() };
  applyParaLevel(acc, ctx.styles.defaults.paraProps);
  for (const l of ctx.tableStyleLayers ?? []) applyParaLevel(acc, l.paraProps);
  for (const s of chain) applyParaLevel(acc, s.paraProps);
  if (numbered !== undefined) applyParaLevel(acc, numberingParaLayer(numbered.level.paraProps));
  if (direct !== undefined) applyParaLevel(acc, direct);

  const p = acc.props;
  return {
    styleId,
    justification: p.justification ?? ('left' as Justification),
    indent: { ...ZERO_INDENT, ...definedOnly(acc.indent) },
    // 布局层要按 x 递增找「下一个制表位」，在这里排好省得每行再排一次
    tabs: [...acc.tabs.values()].sort((a, b) => a.pos - b.pos),
    spacing: {
      before: acc.spacing.before ?? 0,
      after: acc.spacing.after ?? 0,
      beforeLines: acc.spacing.beforeLines ?? 0,
      afterLines: acc.spacing.afterLines ?? 0,
      line: acc.spacing.line ?? DEFAULT_LINE,
      lineRule: acc.spacing.lineRule ?? 'auto',
      beforeAutospacing: acc.spacing.beforeAutospacing ?? false,
      afterAutospacing: acc.spacing.afterAutospacing ?? false,
    },
    keepNext: p.keepNext ?? false,
    keepLines: p.keepLines ?? false,
    pageBreakBefore: p.pageBreakBefore ?? false,
    // 以下四项 Word 默认是**开**的，不是关的
    widowControl: p.widowControl ?? true,
    snapToGrid: p.snapToGrid ?? true,
    autoSpaceDE: p.autoSpaceDE ?? true,
    autoSpaceDN: p.autoSpaceDN ?? true,
    overflowPunct: p.overflowPunct ?? true,
    outlineLevel: p.outlineLevel ?? 9, // 9 = 正文，不进目录
    numbering: {
      ...ref,
      // 编号级的 pPr 里如果又写了 numPr（不合法但见过），上面已经剥掉，所以这里
      // 用的是 ref 而不是 acc.numbering —— 两者只在那种畸形文件上才不同
      // 没有计数器时只有级定义没有「第几」，那就不给 label —— 与其填一个
      // 空文字的假编号，不如让下游一眼看出「这次级联没跑编号」
      ...(numbered !== undefined && 'text' in numbered ? { label: labelOf(ctx, direct, numbered) } : {}),
    },
    // 段落标记的字符属性走同一条字符级联，最后再叠自己那份
    markRunProps: resolveRunProps(ctx, direct, direct?.markRunProps),
  };
}

/** 只级联 `w:numPr` 这一项。缺席时 numId=0，也就是「没有编号」 */
function numberingRefOf(
  ctx: CascadeContext,
  chain: readonly { paraProps: ParaProps }[],
  direct: ParaProps | undefined,
): { numId: number; level: number } {
  const ref: NumberingRef = {};
  const apply = (p: ParaProps): void => {
    if (p.numbering !== undefined) Object.assign(ref, definedOnly(p.numbering));
  };
  apply(ctx.styles.defaults.paraProps);
  for (const s of chain) apply(s.paraProps);
  if (direct !== undefined) apply(direct);
  return { numId: ref.numId ?? 0, level: ref.level ?? 0 };
}

/**
 * 取本段的编号级定义（有计数器时顺便把「第几」算出来）。
 *
 * 两条路都必须能拿到 `level`：缩进是编号级给的，与「第几」无关 ——
 * 只有计数器在场时才有 `text`，没有计数器的段落也不该丢掉悬挂缩进。
 */
function numberingOf(
  ctx: CascadeContext,
  ref: { numId: number; level: number },
  counters: NumberingCounters | undefined,
): NumberedParagraph | { level: NumberingLevel } | undefined {
  if (ref.numId === 0) return undefined;
  if (counters !== undefined) return counters.advance(ref.numId, ref.level);
  const level = numberingLevel(ctx.numbering, ref.numId, ref.level, ctx.styles);
  return level === undefined ? undefined : { level };
}

/**
 * 编号文字自己的字符属性。
 *
 * 层序：docDefaults → 段落样式链 → `w:lvl/w:rPr` → 段落标记的 rPr。
 * 段落标记排在最后是因为它是**用户直接改的那个**（在 Word 里选中 ¶ 调字号，
 * 编号跟着变大，这是所有人都用过的行为），而 `w:lvl/w:rPr` 来自模板。
 */
function labelOf(
  ctx: CascadeContext,
  direct: ParaProps | undefined,
  numbered: NumberedParagraph,
): NumberLabel {
  const level = numbered.level;
  return {
    text: numbered.text,
    value: numbered.value,
    suffix: level.suffix,
    justification: level.justification,
    runProps: resolveRunProps(ctx, direct, direct?.markRunProps, level.runProps),
  };
}

// ── 仍然存在的洞（写下来免得以为已经做了）──────────────────────────────────
//
// 1. **toggle 属性按「后者覆盖」处理，没做 XOR**。规范（§17.7.3）说 b / i / caps /
//    strike / vanish 这类开关属性在**样式层之间**是异或的：祖先开、后代再开 = 关。
//    这里一律按覆盖处理。差异只在「两级以上样式重复设同一个开关」时才显现，
//    真实公文里罕见；更重要的是**没有 Word 真值样本能验证 XOR 的确切边界**，
//    照规范硬写一个测不了的实现，比留一个写明的洞更危险（原则 1.5）。
//    补的办法：造一份多级样式重复设 b 的 docx，跑 fidelity 拿真值，再回来改。
//    表格样式的条件格式（Phase 4）同理，也还没有位置。
