/**
 * 断行 —— 中文公文保真的第一个胜负手（架构 §5 里标橙色的那一步）。
 *
 * 策略是**贪心 + 压缩优先**，不是 Knuth-Plass 的全局最优：Word 自己就是逐行贪心的，
 * 追求全局最优反而会系统性地偏离真值。塞不下时按三条补救措施依次尝试，
 * 全都不行才回退到上一个允许的断点：
 *
 * 1. **挤压**：全角标点压掉空着的半边（`w:characterSpacingControl`）
 * 2. **悬挂**：后置标点允许溢出版心（`w:overflowPunct`，默认开）—— 行尾那个句号吐出边界
 * 3. **回退**：往回找最近的合法断点，把前面的字一起推到下一行
 *
 * ⚠️ 这个顺序是照开发计划 §2.2 的「压缩优先」写的，**没有真值验证**：
 * 挤压与悬挂在 Word 里的确切触发条件（尤其是两者同时可用时谁先）要靠
 * 「一行末尾恰好差半个标点」的样本钉死。顺序错了会差一个字的断行位置。
 *
 * 输出只有**行内水平几何**：没有 y、没有基线。理由见 types.ts 的文件头。
 */
import type { Twips } from '@uw/core';
import type { TabStop } from '@uw/model';
import { canBreakBetween } from './break-class.ts';
import type { LayoutItem } from './types.ts';
import { PUNCT_COMPRESS_RATIO } from './uncalibrated.ts';

export interface LineBreakContext {
  /** 第 `lineIndex` 行（0 起）的可用宽度：版心宽减掉左右缩进与首行缩进 */
  availWidth(lineIndex: number): Twips;
  /** 第 `lineIndex` 行左边缘相对**版心左边**的偏移 —— 制表位是按版心算的，不是按行首 */
  lineLeft(lineIndex: number): Twips;
  tabs: readonly TabStop[];
  defaultTabStop: Twips;
  /** `w:characterSpacingControl` 不是 doNotCompress 时为 true */
  compressPunctuation: boolean;
  /** `w:overflowPunct`，默认开 */
  overflowPunct: boolean;
  /**
   * 编号后那个制表位（`w:suff="tab"`）的**隐含停靠点**：段落左缩进的绝对位置。
   *
   * 带编号的段落几乎都是「悬挂缩进」：编号从 `left - hanging` 开始，正文从 `left` 开始，
   * 中间靠这个制表位跨过去 —— 而 `left` 通常并不在 `w:tabs` 里，也不是
   * `defaultTabStop` 的整数倍。不认这个隐含停靠点，正文会停到 720 twips 的整数倍上，
   * 整段左缩进全错。编号已经超过 `left` 时照常往后找下一个停靠点。
   *
   * ⚠️ 「超过之后落到哪」没有真值：这里按「继续走正常的制表位规则」处理，
   * Word 也可能是「紧跟一个空格」。样本：把编号写长到超过悬挂缩进，看正文起点。
   */
  numberingTabStop?: Twips;
}

/** 一个被命中的制表位，前导符与对齐后处理都要它 */
export interface TabHit {
  /** 在 `items` 里的下标 */
  index: number;
  alignment: TabStop['alignment'];
  leader: TabStop['leader'];
  /** 制表位的绝对位置（相对版心左边）。默认制表位也折算成绝对位置 */
  pos: Twips;
}

export interface BrokenLine {
  /** 对应 `items` 的区间 `[start, end)` */
  start: number;
  end: number;
  /** 行内容宽度：行尾空格与悬挂出去的标点**不算** */
  width: Twips;
  /** 每个 item 的行内 x（相对行首），下标与 `[start, end)` 一一对应 */
  xs: Twips[];
  /** 每个 item 的实际占宽：制表位已解成宽度、挤压已生效 */
  ws: Twips[];
  /** 溢出版心的那些 item（悬挂标点、行尾空格） */
  hanging: boolean[];
  tabs: TabHit[];
  breakAfter?: 'line' | 'page' | 'column';
}

export function breakLines(items: readonly LayoutItem[], ctx: LineBreakContext): BrokenLine[] {
  const lines: BrokenLine[] = [];
  let lineIndex = 0;
  let start = 0;
  let x = 0;
  let xs: Twips[] = [];
  let ws: Twips[] = [];
  let hanging: boolean[] = [];
  let tabs: TabHit[] = [];
  let avail = ctx.availWidth(0);

  function close(end: number, breakAfter?: BrokenLine['breakAfter']): number {
    const line: BrokenLine = { start, end, width: 0, xs, ws, hanging, tabs };
    line.width = contentWidth(line, items);
    if (breakAfter !== undefined) line.breakAfter = breakAfter;
    alignTabTargets(line, items, ctx.lineLeft(lineIndex));
    lines.push(line);
    lineIndex++;
    start = end;
    x = 0;
    xs = [];
    ws = [];
    hanging = [];
    tabs = [];
    avail = ctx.availWidth(lineIndex);
    return end;
  }

  function accept(width: Twips, gap: Twips, hang: boolean): void {
    xs.push(x + gap);
    ws.push(width);
    hanging.push(hang);
    x += gap + width;
  }

  let i = 0;
  while (i < items.length) {
    const item = items[i] as LayoutItem;

    if (item.kind === 'break') {
      // 硬换行本身不占宽度，但它属于这一行 —— 分页要靠它知道「这里断过」
      accept(0, 0, false);
      i = close(i + 1, item.breakType);
      continue;
    }

    const gap = i === start ? 0 : gapOf(item);
    let width = widthOf(item, x + gap, ctx, lineIndex);
    if (item.kind === 'tab') recordTab(tabs, i, x + gap, ctx, lineIndex, item.numbering === true);

    if (x + gap + width <= avail || i === start) {
      // 一个 item 比整行还宽时也必须收下（否则死循环），让它溢出去
      accept(width, gap, false);
      i++;
      continue;
    }

    // ① 挤压
    if (item.kind === 'char' && item.compressible && ctx.compressPunctuation) {
      const compressed = width * (1 - PUNCT_COMPRESS_RATIO);
      if (x + gap + compressed <= avail) {
        accept(compressed, gap, false);
        i++;
        continue;
      }
      width = compressed; // 压过之后仍塞不下，后面按压缩宽度悬挂
    }

    // ② 悬挂：后置标点与行尾空格可以吐出版心
    if (item.kind === 'char' && (item.space || (item.kinsoku === 'noStart' && ctx.overflowPunct))) {
      accept(width, gap, true);
      i++;
      continue;
    }

    // ③ 回退到最近的合法断点。禁则已经写进 canBreakBetween，这里不必再修一遍
    let b = i;
    while (b > start && !canBreakBetween(items[b - 1], items[b] as LayoutItem)) b--;
    if (b <= start) b = i; // 整行没有合法断点：只能硬断在这里
    // 断点之后的 item 要重排，把它们从当前行摘掉
    xs.length = b - start;
    ws.length = b - start;
    hanging.length = b - start;
    tabs = tabs.filter((t) => t.index < b);
    x = xs.length === 0 ? 0 : (xs[xs.length - 1] as Twips) + (ws[ws.length - 1] as Twips);
    i = close(b);
  }

  if (start < items.length || lines.length === 0) close(items.length);
  return lines;
}

function gapOf(item: LayoutItem): Twips {
  return item.kind === 'char' || item.kind === 'object' ? item.gapBefore : 0;
}

function widthOf(item: LayoutItem, x: Twips, ctx: LineBreakContext, lineIndex: number): Twips {
  if (item.kind === 'tab') {
    return tabAdvance(ctx.lineLeft(lineIndex) + x, ctx, item.numbering === true).width;
  }
  if (item.kind === 'break') return 0;
  return item.width;
}

function recordTab(
  out: TabHit[],
  index: number,
  x: Twips,
  ctx: LineBreakContext,
  lineIndex: number,
  numbering: boolean,
): void {
  const hit = tabAdvance(ctx.lineLeft(lineIndex) + x, ctx, numbering);
  out.push({ index, alignment: hit.alignment, leader: hit.leader, pos: hit.pos });
}

/**
 * 制表位推进：先找显式制表位，没有再落到 `defaultTabStop` 的整数倍。
 *
 * `bar` 型不参与推进 —— 它是「在这个位置画一条竖线」的装饰，不是停靠点。
 * 把它当停靠点会让那一行的文字整体右移。
 */
function tabAdvance(
  absolute: Twips,
  ctx: LineBreakContext,
  numbering = false,
): { width: Twips; alignment: TabStop['alignment']; leader: TabStop['leader']; pos: Twips } {
  // 编号后的制表位先看那个隐含停靠点（段落左缩进），它优先于显式制表位 ——
  // 悬挂缩进的正文就落在那儿，见 LineBreakContext.numberingTabStop
  const implicit = ctx.numberingTabStop;
  if (numbering && implicit !== undefined && implicit > absolute) {
    return { width: implicit - absolute, alignment: 'left', leader: 'none', pos: implicit };
  }
  for (const t of ctx.tabs) {
    if (t.alignment === 'bar' || t.alignment === 'clear') continue;
    if (t.pos > absolute)
      return { width: t.pos - absolute, alignment: t.alignment, leader: t.leader, pos: t.pos };
  }
  const step = ctx.defaultTabStop > 0 ? ctx.defaultTabStop : 720;
  const pos = (Math.floor(absolute / step) + 1) * step;
  return { width: pos - absolute, alignment: 'left', leader: 'none', pos };
}

/**
 * 右对齐 / 居中 / 小数点对齐的制表位：**断完行之后**再把它后面那段拉到位。
 *
 * 断行时一律按「推进到停靠点」估宽（也就是当成左对齐），因为那时还不知道这一段
 * 后面会跟多少字。估宽偏大而不偏小，所以只会让行断得略早，不会让文字溢出版心 ——
 * 公文里靠右对齐制表位排的是「签发人」「页码」这种短内容，差异到不了一个字。
 * 真要消除，得像 Word 那样对含非左对齐制表位的行多跑一轮，等有真值样本了再说。
 */
function alignTabTargets(line: BrokenLine, items: readonly LayoutItem[], lineLeft: Twips): void {
  for (const hit of line.tabs) {
    if (hit.alignment === 'left' || hit.alignment === 'bar' || hit.alignment === 'clear') continue;
    const k = hit.index - line.start;
    const segStart = k + 1;
    // 这一段一直延伸到下一个制表位或行尾
    const next = line.tabs.find((t) => t.index > hit.index);
    const segEnd = next === undefined ? line.xs.length : next.index - line.start;
    if (segStart >= segEnd) continue;

    const segX = line.xs[segStart] as Twips;
    const segWidth = (line.xs[segEnd - 1] as Twips) + (line.ws[segEnd - 1] as Twips) - segX;
    const target =
      hit.alignment === 'center'
        ? hit.pos - lineLeft - segWidth / 2
        : hit.alignment === 'right'
          ? hit.pos - lineLeft - segWidth
          : decimalTarget(line, items, segStart, segEnd, hit.pos - lineLeft);
    // 挤不下时停靠点失效，文字紧跟着制表位排 —— Word 也是这么退的
    const delta = Math.max(target - segX, -(line.ws[k] as Twips));
    line.ws[k] = (line.ws[k] as Twips) + delta;
    for (let j = segStart; j < line.xs.length; j++) line.xs[j] = (line.xs[j] as Twips) + delta;
    line.width = contentWidth(line, items);
  }
}

/** 小数点对齐：小数点落在停靠点上；这一段里没有小数点时按右对齐处理（Word 的行为） */
function decimalTarget(
  line: BrokenLine,
  items: readonly LayoutItem[],
  segStart: number,
  segEnd: number,
  stop: Twips,
): Twips {
  for (let j = segStart; j < segEnd; j++) {
    const item = items[line.start + j] as LayoutItem;
    if (item.kind === 'char' && (item.cp === 0x2e || item.cp === 0xff0e)) {
      return stop - ((line.xs[j] as Twips) - (line.xs[segStart] as Twips));
    }
  }
  const segX = line.xs[segStart] as Twips;
  return stop - ((line.xs[segEnd - 1] as Twips) + (line.ws[segEnd - 1] as Twips) - segX);
}

/**
 * 行宽**不含行尾**的空格与悬挂标点。
 *
 * 行尾空格必须扣掉，否则「居中」的中文标题只要末尾多打了一个空格就会整体左移半个字宽 ——
 * 这是肉眼可见的错位，而公文标题末尾带空格相当常见。行中间的空格照常计入。
 */
function contentWidth(line: BrokenLine, items: readonly LayoutItem[]): Twips {
  for (let k = line.xs.length - 1; k >= 0; k--) {
    const item = items[line.start + k];
    if (line.hanging[k] === true) continue;
    if (item !== undefined && item.kind === 'char' && item.space) continue;
    return (line.xs[k] as Twips) + (line.ws[k] as Twips);
  }
  return 0;
}
