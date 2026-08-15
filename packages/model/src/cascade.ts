/**
 * 样式级联 —— Phase 1 的正题。
 *
 * 顺序（ECMA-376 §17.7.2，后面的覆盖前面的）：
 *
 * ```
 * 段落属性： docDefaults.pPr → 段落样式链(祖先→自己) → [编号 pPr] → 直接 pPr
 * 字符属性： docDefaults.rPr → 段落样式链.rPr      → [编号 rPr] → 字符样式链.rPr → 直接 rPr
 * ```
 *
 * 两处容易错的地方：
 * 1. **段落样式也带字符属性**，且它排在字符样式**前面**。漏了这一层，
 *    「标题 1 是 24 号字」就没了
 * 2. **合并是逐属性的，不是整块替换**。样式设了 `w:ind w:left`，直接格式设了
 *    `w:ind w:firstLine`，结果两个都在。整块替换会把左缩进吃掉
 *
 * 方括号里的编号那层是 Phase 5 的洞，见文件末尾。
 */
import { halfPtToTwips } from '@uw/core';
import type {
  Indent,
  Justification,
  NumberingRef,
  ParagraphSpacing,
  ParaProps,
  ResolvedParaProps,
  ResolvedRunProps,
  RunProps,
} from './props.ts';
import type { DocumentSettings } from './settings.ts';
import type { StyleSheet } from './styles.ts';
import type { Theme } from './theme.ts';
import { resolveThemeFont } from './theme.ts';

export interface CascadeContext {
  styles: StyleSheet;
  theme: Theme;
  /**
   * `settings.xml`。级联只用到 `themeFontLang.eastAsia` 一项，但**这一项非有不可**：
   * 主题里 `a:ea` 是空串时，东亚字体要按语言回退到 `a:font script="..."`，
   * 而语言的来源就是这里。缺了它只能硬编码 zh-CN，日文文档会拿到简体中文字体。
   */
  settings: DocumentSettings;
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
 */
export function resolveRunProps(
  ctx: CascadeContext,
  paraProps: ParaProps | undefined,
  direct: RunProps | undefined,
): ResolvedRunProps {
  const acc: RunAccum = { props: {}, slots: {}, hint: undefined };

  applyRunLevel(acc, ctx.styles.defaults.runProps);
  for (const s of ctx.styles.chainOf(paragraphStyleId(ctx, paraProps))) applyRunLevel(acc, s.runProps);
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

export function resolveParaProps(ctx: CascadeContext, direct: ParaProps | undefined): ResolvedParaProps {
  const styleId = paragraphStyleId(ctx, direct);
  const acc: ParaAccum = { props: {}, indent: {}, spacing: {}, numbering: {} };

  applyParaLevel(acc, ctx.styles.defaults.paraProps);
  for (const s of ctx.styles.chainOf(styleId)) applyParaLevel(acc, s.paraProps);
  if (direct !== undefined) applyParaLevel(acc, direct);

  const p = acc.props;
  return {
    styleId,
    justification: p.justification ?? ('left' as Justification),
    indent: { ...ZERO_INDENT, ...definedOnly(acc.indent) },
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
    numbering: { numId: acc.numbering.numId ?? 0, level: acc.numbering.level ?? 0 },
    // 段落标记的字符属性走同一条字符级联，最后再叠自己那份
    markRunProps: resolveRunProps(ctx, direct, direct?.markRunProps),
  };
}

// ── 两个已知的洞（写下来免得以为已经做了）────────────────────────────────────
//
// 1. **编号那一层没接**。`w:numPr` 指向的 `numbering.xml` 里，每个级别自带 pPr / rPr
//    （常见的是缩进与项目符号字体），它排在段落样式之后、直接格式之前。
//    Phase 5 做编号时补，接口位置就是 applyParaLevel / applyRunLevel 之间。
//
// 2. **toggle 属性按「后者覆盖」处理，没做 XOR**。规范（§17.7.3）说 b / i / caps /
//    strike / vanish 这类开关属性在**样式层之间**是异或的：祖先开、后代再开 = 关。
//    这里一律按覆盖处理。差异只在「两级以上样式重复设同一个开关」时才显现，
//    真实公文里罕见；更重要的是**没有 Word 真值样本能验证 XOR 的确切边界**，
//    照规范硬写一个测不了的实现，比留一个写明的洞更危险（原则 1.5）。
//    补的办法：造一份多级样式重复设 b 的 docx，跑 fidelity 拿真值，再回来改。
//    表格样式的条件格式（Phase 4）同理，也还没有位置。
