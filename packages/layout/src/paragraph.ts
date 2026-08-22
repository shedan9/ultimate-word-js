/**
 * 段落装配：缩进 → 断行 → 对齐 → 行内片段。
 *
 * 这是 Phase 2 在**基线穿刺之前**能做完的最后一步。产出 `ParagraphLayout`：
 * 每一行的行内 x 全部定死、行高总量算好，唯独没有 y —— 补完穿刺后给 `LineLayout`
 * 加上 `baseline` 并把行按顺序摞起来，就是行盒装配与分页。
 */
import type { Twips } from '@uw/core';
import type { TextMeasurer } from '@uw/fonts';
import type { DocGrid, DocumentSettings, NodeId, ResolvedParagraph, ResolvedParaProps } from '@uw/model';
import type { KinsokuSets } from './break-class.ts';
import { isNumberingItem, kinsokuFrom } from './break-class.ts';
import { buildItems } from './items.ts';
import { lineHeight } from './line-height.ts';
import type { BrokenLine, LineBreakContext } from './linebreak.ts';
import { breakLines } from './linebreak.ts';
import type {
  LayoutItem,
  LineFloat,
  LineFragment,
  LineLayout,
  LineObject,
  ParagraphLayout,
  TabLeader,
} from './types.ts';
import { CHAR_UNIT_EM } from './uncalibrated.ts';

export interface LayoutParagraphOptions {
  measurer: TextMeasurer;
  /** 版心宽度（页宽减左右页边距）。表格单元格里就是单元格的可用宽 */
  contentWidth: Twips;
  settings: DocumentSettings;
  docGrid: DocGrid;
  /** 四个字体桶全空时用哪款字体 */
  defaultFont?: string;
  /** 域求值的结果（run id → 显示的文字），见 `BuildItemsOptions.fieldValues` */
  fieldValues?: ReadonlyMap<NodeId, string>;
}

export function layoutParagraph(p: ResolvedParagraph, opts: LayoutParagraphOptions): ParagraphLayout {
  const kinsoku: KinsokuSets = kinsokuFrom(opts.settings);
  const itemOpts = {
    measurer: opts.measurer,
    kinsoku,
    // 相邻标点挤压是**常态排版**，所以在 item 流那一步就做掉，不等断行
    compressPunctuation: opts.settings.characterSpacingControl !== 'doNotCompress',
    ...(opts.defaultFont === undefined ? {} : { defaultFont: opts.defaultFont }),
    ...(opts.fieldValues === undefined ? {} : { fieldValues: opts.fieldValues }),
  };
  const items = buildItems(p, itemOpts);

  const geom = indentGeometry(p, items, opts.contentWidth);
  const ctx: LineBreakContext = {
    availWidth: (n) => (n === 0 ? geom.firstAvail : geom.restAvail),
    lineLeft: (n) => (n === 0 ? geom.firstLeft : geom.left),
    tabs: p.props.tabs,
    defaultTabStop: opts.settings.defaultTabStop,
    compressPunctuation: opts.settings.characterSpacingControl !== 'doNotCompress',
    // 临时挤压是两端对齐才有的行为（实测，见 LineBreakContext.justified）
    justified: p.props.justification === 'both' || p.props.justification === 'distribute',
    overflowPunct: p.props.overflowPunct,
    // 编号后的制表位停在正文的左边缘（也就是悬挂缩进落脚处），见 linebreak.ts
    numberingTabStop: geom.left,
  };

  const broken = breakLines(items, ctx);
  const lines = broken.map((line, n) =>
    assemble(line, items, p.props, {
      isLast: n === broken.length - 1,
      left: ctx.lineLeft(n),
      avail: ctx.availWidth(n),
      measurer: opts.measurer,
      docGrid: opts.docGrid,
      ...(opts.defaultFont === undefined ? {} : { defaultFont: opts.defaultFont }),
    }),
  );

  const firstHeight = lines[0]?.height ?? 0;
  return {
    paragraphId: p.id,
    lines,
    spaceBefore: spacing(p.props.spacing.before, p.props.spacing.beforeLines, firstHeight),
    spaceAfter: spacing(p.props.spacing.after, p.props.spacing.afterLines, firstHeight),
    contentWidth: opts.contentWidth,
  };
}

/**
 * 段前 / 段后间距：`w:beforeLines` 是 1/100 **行**，`w:before` 是 twips。
 * 两者同时存在时以行单位为准（与缩进的字符单位同理，Word 为兼容旧版会两个都写）。
 */
function spacing(twips: Twips, lines: number, lineHeightTwips: Twips): Twips {
  return lines > 0 ? (lines / 100) * lineHeightTwips : twips;
}

interface IndentGeometry {
  /** 首行左边缘（相对版心左边） */
  firstLeft: Twips;
  /** 其余行左边缘 */
  left: Twips;
  firstAvail: Twips;
  restAvail: Twips;
}

/**
 * 缩进 —— 公文里每段都有的「首行缩进 2 字符」就在这儿。
 *
 * 两条容易错的：
 * 1. **字符单位优先**：`w:firstLineChars` 与 `w:firstLine` 同时存在时（Word 为兼容旧版
 *    常常两个都写），按字符单位算。选错会让每段首行差几十 twips
 * 2. **悬挂缩进是负的首行缩进**：`w:hanging` 让首行往左伸出去，其余行按 `w:left` 排；
 *    它与 `w:firstLine` 互斥，同时出现时 hanging 赢（Word 的行为）
 *
 * 「一个字符」的宽度取段落里第一个字符的字号（没有字符就取段落标记的字号），
 * 按 1 em 折算，理由见 uncalibrated.ts 的 CHAR_UNIT_EM。
 */
function indentGeometry(
  p: ResolvedParagraph,
  items: readonly LayoutItem[],
  contentWidth: Twips,
): IndentGeometry {
  const unit = charUnit(p, items);
  const ind = p.props.indent;
  const left = ind.leftChars !== 0 ? (ind.leftChars / 100) * unit : ind.left;
  const right = ind.rightChars !== 0 ? (ind.rightChars / 100) * unit : ind.right;
  const hanging = ind.hangingChars !== 0 ? (ind.hangingChars / 100) * unit : ind.hanging;
  const firstLine = ind.firstLineChars !== 0 ? (ind.firstLineChars / 100) * unit : ind.firstLine;
  const firstLeft = hanging !== 0 ? left - hanging : left + firstLine;

  return {
    firstLeft,
    left,
    // 负缩进（首行伸到版心外）在公文里是合法排法，可用宽度跟着变大，不夹到 0
    firstAvail: contentWidth - firstLeft - right,
    restAvail: contentWidth - left - right,
  };
}

/**
 * 「一个字符」有多宽 —— 取段落里**第一个正文字符**的字号。
 *
 * 跳过编号文字：项目符号常常是 Symbol 字体的另一个字号，拿它当尺子会让整段的
 * 「首行缩进 2 字符」按错误的字号折算。段落里一个正文字符都没有时取段落标记的字号。
 */
function charUnit(p: ResolvedParagraph, items: readonly LayoutItem[]): Twips {
  for (const item of items) {
    if (item.kind === 'char' && !isNumberingItem(item)) return item.fontSize * CHAR_UNIT_EM;
  }
  return p.props.markRunProps.size * CHAR_UNIT_EM;
}

interface AssembleContext {
  isLast: boolean;
  left: Twips;
  avail: Twips;
  measurer: TextMeasurer;
  docGrid: DocGrid;
  defaultFont?: string;
}

function assemble(
  line: BrokenLine,
  items: readonly LayoutItem[],
  props: ResolvedParaProps,
  ctx: AssembleContext,
): LineLayout {
  const height = lineHeight(items, line, props, {
    measurer: ctx.measurer,
    docGrid: ctx.docGrid,
    ...(ctx.defaultFont === undefined ? {} : { defaultFont: ctx.defaultFont }),
  });

  const xs = line.xs.slice();
  const { offset, stretched } = justify(xs, line, items, props, ctx);

  const out: LineLayout = {
    start: line.start,
    end: line.end,
    x: ctx.left + offset,
    // 拉伸过的行正好占满可用宽度 —— 回归比对时这一项要能直接和真值的行宽对上
    width: stretched ? ctx.avail : line.width,
    height: height.height,
    baseline: height.baseline,
    natural: height.natural,
    fragments: fragmentsOf(line, items, xs, ctx.left + offset),
    leaders: leadersOf(line, xs, ctx.left + offset),
    isLast: ctx.isLast,
  };
  const { objects, floats } = objectsOf(line, items, xs, ctx.left + offset);
  if (objects.length > 0) out.objects = objects;
  if (floats.length > 0) out.floats = floats;
  if (line.breakAfter !== undefined) out.breakAfter = line.breakAfter;
  return out;
}

/**
 * 对齐。左对齐 / 居中 / 右对齐只是整行平移，两端对齐与分散对齐要把多余的宽度**摊到行内**。
 *
 * 摊法：行里有空格就摊在空格上（拉丁文的做法），没有空格就摊在**字与字之间**
 * （中文的做法）。中文行里加空格是没有的事，所以这个分支就等价于「中文摊字距、
 * 英文摊词距」，也是 Word 的实际观感。
 *
 * ⚠️ 中英混排的行两种摊法各占多少，Word 有自己的权重，这里没有真值：现在只要有空格
 * 就全摊在空格上。混排行的逐字 x 会有偏差（**不影响断行点**，断行早于对齐）。
 *
 * 三个不拉伸的例外：段落最后一行、以硬换行（`w:br`）结束的行、行宽已经不小于可用宽度。
 * 分散对齐（`distribute`）连最后一行也拉，这正是它与两端对齐的区别。
 */
function justify(
  xs: Twips[],
  line: BrokenLine,
  items: readonly LayoutItem[],
  props: ResolvedParaProps,
  ctx: AssembleContext,
): { offset: Twips; stretched: boolean } {
  const slack = ctx.avail - line.width;
  const jc = props.justification;

  if (jc === 'center') return { offset: slack / 2, stretched: false };
  if (jc === 'right') return { offset: slack, stretched: false };
  if (jc !== 'both' && jc !== 'distribute') return { offset: 0, stretched: false };

  const stretchable = jc === 'distribute' || (!ctx.isLast && line.breakAfter === undefined);
  if (!stretchable || slack <= 0) return { offset: 0, stretched: false };

  // 可以张开的位置：优先空格，其次字与字之间。行尾那一个不算 —— 在它后面加空隙等于没加
  const spaces: number[] = [];
  const gaps: number[] = [];
  for (let k = 0; k < xs.length - 1; k++) {
    const item = items[line.start + k];
    if (item === undefined || item.kind === 'break') continue;
    // 编号内部不张开：Word 两端对齐拉的是正文，编号与它后面那个制表位纹丝不动
    if (isNumberingItem(item)) continue;
    if (item.kind === 'char' && item.space) spaces.push(k);
    gaps.push(k);
  }
  const points = spaces.length > 0 ? spaces : gaps;
  if (points.length === 0) return { offset: 0, stretched: false };

  const step = slack / points.length;
  let shift = 0;
  let next = 0;
  for (let k = 0; k < xs.length; k++) {
    xs[k] = (xs[k] as Twips) + shift;
    if (points[next] === k) {
      shift += step;
      next++;
    }
  }
  return { offset: 0, stretched: true };
}

/**
 * 合成渲染片段：一行里「同 run、同字体、同字号」的连续字符合成一段。
 *
 * 不一字一元素（DOM 会爆），也不整行一元素（逐字微调没处放）—— `glyphX` 承载
 * 挤压 / 间距 / 两端对齐造成的所有逐字偏移，直接喂给 SVG `<text x="…">`。
 */
function fragmentsOf(
  line: BrokenLine,
  items: readonly LayoutItem[],
  xs: readonly Twips[],
  lineX: Twips,
): LineFragment[] {
  const out: LineFragment[] = [];
  let current: LineFragment | undefined;
  let currentEnd = 0;

  for (let k = 0; k < xs.length; k++) {
    const item = items[line.start + k];
    if (item === undefined || item.kind !== 'char') {
      current = undefined;
      continue;
    }
    // 软连字符只有在行尾断开时才显形，行中间的一律不画
    if (item.softHyphen === true && line.start + k !== line.end - 1) {
      current = undefined;
      continue;
    }
    const x = lineX + (xs[k] as Twips);
    const ch = String.fromCodePoint(item.cp);
    if (
      current === undefined ||
      current.runId !== item.runId ||
      current.font !== item.font ||
      current.fontSize !== item.fontSize ||
      current.script !== item.script
    ) {
      current = {
        runId: item.runId,
        font: item.font,
        fontSize: item.fontSize,
        script: item.script,
        style: item.style,
        text: ch,
        x,
        width: item.width,
        glyphX: [x],
        // 编号自成片段（runId 不同，天然分开），标出来让渲染层别把它算进可选文本
        ...(item.numbering === true ? { numbering: true as const } : {}),
        // 域结果同理自成片段（整个 run 都被换掉了），标记的用途见 CharItem.field
        ...(item.field === true ? { field: true as const } : {}),
      };
      out.push(current);
    } else {
      current.text += ch;
      current.glyphX.push(x);
    }
    currentEnd = x + (line.ws[k] as Twips);
    current.width = currentEnd - current.x;
  }
  return out;
}

/**
 * 行内的图片 / 图形。
 *
 * 与文字分开收，是因为它们在渲染层是完全不同的元素（`<image>` vs `<text>`），
 * 而且**没有逐字 x** —— 一个对象就是一个矩形，两端对齐拉开的空隙落在它两侧，不落在内部。
 *
 * 垂直方向这里不给 y：**对象的底边坐在基线上**（`line-height.ts` 让它整个高度都算进
 * 基线以上），渲染时 y = 基线 − 高度。⚠️ 这条与「核心盒在行高里居中」一样还没有 Word 真值，
 * 钉死它的样本写在 `uncalibrated.ts` 的 `OBJECT_SITS_ON_BASELINE`。
 */
function objectsOf(
  line: BrokenLine,
  items: readonly LayoutItem[],
  xs: readonly Twips[],
  lineX: Twips,
): { objects: LineObject[]; floats: LineFloat[] } {
  const objects: LineObject[] = [];
  const floats: LineFloat[] = [];
  for (let k = 0; k < xs.length; k++) {
    const item = items[line.start + k];
    if (item === undefined || item.kind !== 'object') continue;
    const x = lineX + (xs[k] as Twips);
    const common = {
      runId: item.runId,
      contentIndex: item.contentIndex,
      x,
      objectKind: item.objectKind,
      ...(item.image === undefined ? {} : { image: item.image }),
      ...(item.alt === undefined ? {} : { alt: item.alt }),
      ...(item.graphic === undefined ? {} : { graphic: item.graphic }),
    };
    if (item.float !== undefined) {
      floats.push({
        ...common,
        width: item.float.width,
        height: item.float.height,
        anchor: item.float.anchor,
      });
    } else {
      objects.push({ ...common, width: item.width, height: item.height });
    }
  }
  return { objects, floats };
}

function leadersOf(line: BrokenLine, xs: readonly Twips[], lineX: Twips): TabLeader[] {
  const out: TabLeader[] = [];
  for (const hit of line.tabs) {
    if (hit.leader === 'none') continue;
    const k = hit.index - line.start;
    const x1 = lineX + (xs[k] as Twips);
    out.push({ x1, x2: x1 + (line.ws[k] as Twips), leader: hit.leader });
  }
  return out;
}
