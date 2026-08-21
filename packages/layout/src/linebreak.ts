/**
 * 断行 —— 中文公文保真的第一个胜负手（架构 §5 里标橙色的那一步）。
 *
 * 策略是**逐行贪心**，不是 Knuth-Plass 的全局最优：Word 自己就是逐行贪心的，
 * 追求全局最优反而会系统性地偏离真值。塞不下时按三条补救措施依次尝试，
 * 全都不行才回退到上一个允许的断点：
 *
 * 1. **悬挂**：后置标点与行尾空格允许溢出版心（`w:overflowPunct`，默认开）。
 *    吐出去的**只是空的那半边**，标点的墨要留在版心内（`HANG_INSIDE_RATIO`，实测）——
 *    所以连半宽都塞不下时，得先挤压腾出地方再挂
 * 2. **挤压**：挂不出去（溢出的是汉字或拉丁字）才挤，把**整行**的全角标点挤到刚好够 ——
 *    见 `compress()`
 * 3. **回退**：往回找最近的合法断点，把前面的字一起推到下一行
 *
 * **先挂后挤**是实测的，与开发计划 §2.2 写的「压缩优先」相反。判据是 gongwen-01 的两行
 * （行号按真值的顺序，0 起）：第 4 行溢出的是「，」，Word 把它挂出版心 8.05pt 就收行了，
 * **没有**再去多收一个字；第 3 行溢出的是「自」（挂不了），Word 才挤出 6pt 把它留在行内。
 * 顺序反了会差一个字，而且错会顺着往后每一行传下去。
 *
 * 另有一条**与断行无关**的常态挤压：相邻两个全角标点固定挤掉半个字，
 * 那个在 `items.ts` 的 `applyPunctPairs` 里、`buildItems` 阶段就做完了。
 *
 * 输出只有**行内水平几何**：没有 y、没有基线。理由见 types.ts 的文件头。
 */
import type { Twips } from '@uw/core';
import type { TabStop } from '@uw/model';
import { canBreakBetween, HANG_INSIDE_RATIO } from './break-class.ts';
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
  /** 行内容宽度：行尾空格不算，悬挂的标点只算留在版心内的半个（见 `contentWidth`） */
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
  /** 每个 item 前面的间隙（中西文自动间距）。回头挤压之后要靠它重排 xs */
  let gaps: Twips[] = [];
  /** 每个 item 还能挤掉多少宽度。只有全角标点非零，压过之后相应减少 */
  let room: Twips[] = [];
  let tabs: TabHit[] = [];
  /** 这一行已经有东西吐出版心了 —— 见循环里那段注释，它决定行到此为止 */
  let hung = false;
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
    gaps = [];
    room = [];
    tabs = [];
    hung = false;
    avail = ctx.availWidth(lineIndex);
    return end;
  }

  function accept(width: Twips, gap: Twips, hang: boolean, compressRoom = 0): void {
    xs.push(x + gap);
    ws.push(width);
    hanging.push(hang);
    gaps.push(gap);
    room.push(compressRoom);
    x += gap + width;
  }

  /**
   * 这个 item 还能挤掉多少 —— 只有全角标点能挤，且要文档没关掉挤压。
   *
   * `gapBefore < 0` 的标点排除在外：它已经因为「紧跟着上一个标点」交出了空半边
   * （`items.ts` 的 `applyPunctPairs`），再挤就是把墨压掉了。
   */
  function roomOf(item: LayoutItem, width: Twips): Twips {
    if (item.kind !== 'char' || !item.compressible || !ctx.compressPunctuation) return 0;
    return item.gapBefore < 0 ? 0 : width * PUNCT_COMPRESS_RATIO;
  }

  /**
   * 回头把这一行**已经放下**的标点挤掉一点，给正要放进来的字腾地方。
   *
   * 这是 gongwen-01 的真值逼出来的：`为落实……通知如下。自` 那一行（真值第 3 行，0 起），
   * 26 个三号字自然宽 416pt、可用宽只有 410.25pt，Word 却把 26 个字全塞进去了，
   * 行宽 410.00pt —— 它把行内的「，」「。」各挤掉一点（合计 6pt），
   * 而不是把「自」推到下一行。只挤「正在溢出的那个字」
   * （下面 ① 的老做法）解释不了这一行，因为溢出的是「自」，它不是标点。
   *
   * 挤多少：**刚好够**，不是一律 50%。同样是那一行的证据（Word 只挤了 6pt，
   * 而两个标点合起来能挤 16pt）。多个标点之间按等额分摊，先到 50% 上限的不再摊 ——
   * 这就是下面按 `room` 升序做的注水。
   *
   * ⚠️ 未标定两处，都在 `uncalibrated.ts` 的 `PUNCT_COMPRESS_RATIO` 里：
   * ① **挤到多少就该放弃**（Word 接受过 9.30pt、拒绝过 13.75pt）—— 这是 L2 剩下 7 行的唯一原因；
   * ② 等额分摊是判断，Word 也可能优先挤行尾那个。分摊方式只影响行内逐字 x（L4），不影响断点。
   *
   * @param deficit 还差多少宽度
   * @param incoming 正要放进来的那个 item 自己能挤掉多少
   * @returns 真挤出来的宽度与其中属于 incoming 的部分；挤不够时调用方走下一条补救措施
   */
  function compress(deficit: Twips, incoming: Twips): { got: Twips; fromIncoming: Twips } {
    // 制表位之前的 item 不能动：挤了它们，制表位就不再停在它该停的位置上。
    // 之后的可以动 —— 挤它们不改变制表位自身的宽度与位置
    const lastTab = tabs.length === 0 ? -1 : (tabs[tabs.length - 1] as TabHit).index - start;
    const ks = room
      .map((_, k) => k)
      .filter((k) => k > lastTab && (room[k] as Twips) > 0)
      .sort((a, b) => (room[a] as Twips) - (room[b] as Twips));

    // 注水：按剩余额度升序摊，额度小的先摊满，省下的自动流给额度大的
    const caps = [...ks.map((k) => room[k] as Twips), incoming];
    const order = caps.map((_, n) => n).sort((a, b) => (caps[a] as Twips) - (caps[b] as Twips));
    const take = new Array<Twips>(caps.length).fill(0);
    let need = deficit;
    let left = caps.length;
    for (const n of order) {
      const t = Math.min(need / left, caps[n] as Twips);
      take[n] = t;
      need -= t;
      left--;
    }
    const got = deficit - need;
    // 差一点点就够时是浮点误差，别把整行推到下一行
    if (got < deficit - 1e-9) return { got: 0, fromIncoming: 0 };

    ks.forEach((k, n) => {
      const t = take[n] as Twips;
      ws[k] = (ws[k] as Twips) - t;
      room[k] = (room[k] as Twips) - t;
    });
    // 挤过之后 xs 全都要往左挪 —— 靠 gaps 重排，不要试图增量修补
    let cur = 0;
    for (let k = 0; k < ws.length; k++) {
      cur += gaps[k] as Twips;
      xs[k] = cur;
      cur += ws[k] as Twips;
    }
    x = cur;
    return { got, fromIncoming: take[caps.length - 1] as Twips };
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
    const width = widthOf(item, x + gap, ctx, lineIndex);
    if (item.kind === 'tab') recordTab(tabs, i, x + gap, ctx, lineIndex, item.numbering === true);

    if (x + gap + width <= avail || i === start) {
      // 一个 item 比整行还宽时也必须收下（否则死循环），让它溢出去
      accept(width, gap, false, roomOf(item, width));
      i++;
      continue;
    }

    // 悬挂过之后这一行就结束了：吐出版心的标点按定义在行尾，后面不可能再有字。
    // 唯一的例外是空格 —— 行尾的一串空格要一起吐出去，不能把第二个空格推到下一行的行首。
    // 少了这一条，① 的挤压会在悬挂之后又腾出地方来，把下一个字硬拉回这一行
    // （gongwen-01 真值第 4 行就是这么错的：Word 让「，」悬挂出去 8.05pt 然后收行）。
    if (hung) {
      if (item.kind === 'char' && item.space) {
        accept(width, gap, true);
        i++;
        continue;
      }
      i = close(i);
      continue;
    }

    // ① 悬挂：后置标点与行尾空格可以吐出版心。**吐出去的只是空的那半边** ——
    // 标点的墨要留在版心内（HANG_INSIDE_RATIO，实测），所以半宽也塞不下时得先挤压。
    // 行尾空格没有墨，整个吐出去
    if (item.kind === 'char' && (item.space || (item.kinsoku === 'noStart' && ctx.overflowPunct))) {
      const inside = item.space ? 0 : width * HANG_INSIDE_RATIO;
      const over = x + gap + inside - avail;
      if (over <= 0 || compress(over, 0).got > 0) {
        accept(width, gap, true);
        hung = true;
        i++;
        continue;
      }
      // 墨塞不进去、也挤不出地方 —— 落到 ③ 把它整个推到下一行
    }

    // ② 挤压：挂不出去（溢出的是汉字或拉丁字）才挤，把这一行的全角标点挤到刚好够
    const incomingRoom = roomOf(item, width);
    const deficit = x + gap + width - avail;
    const { got, fromIncoming } = compress(deficit, incomingRoom);
    if (got > 0) {
      accept(width - fromIncoming, gap, false, incomingRoom - fromIncoming);
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
    gaps.length = b - start;
    room.length = b - start;
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
 * 行宽**不含行尾空格**，悬挂标点只算**留在版心内的那半个**。
 *
 * 行尾空格必须整个扣掉，否则「居中」的中文标题只要末尾多打了一个空格就会整体左移半个字宽 ——
 * 这是肉眼可见的错位，而公文标题末尾带空格相当常见。行中间的空格照常计入。
 *
 * 悬挂的标点不一样：它的墨还在版心里（`HANG_INSIDE_RATIO`，实测），只有空半边吐了出去。
 * 整个不计入的话，两端对齐会拿这半个字当成空隙再拉一次，行内每个字都跟着往右挪 ——
 * 真值第 4 行的行尾「，」左边缘正好落在「版心宽 − 半个字」上。
 */
function contentWidth(line: BrokenLine, items: readonly LayoutItem[]): Twips {
  for (let k = line.xs.length - 1; k >= 0; k--) {
    const item = items[line.start + k];
    const space = item !== undefined && item.kind === 'char' && item.space;
    if (space) continue;
    const ratio = line.hanging[k] === true ? HANG_INSIDE_RATIO : 1;
    return (line.xs[k] as Twips) + (line.ws[k] as Twips) * ratio;
  }
  return 0;
}
