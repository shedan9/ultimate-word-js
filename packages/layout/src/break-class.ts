/**
 * 断点判定与中文禁则（避头尾）。
 *
 * 这是断行算法的**规则表**部分，故意与算法（linebreak.ts）分开：规则表要按 Word 的
 * 实际行为一点点校准，算法不该跟着改。
 *
 * 与 UAX#14 的关系：只取它的骨架（空格后可断、CJK 之间可断、拉丁词内不可断），
 * **不**实现完整的 30 多个 line breaking class。理由是我们的判据不是「符合 UAX#14」，
 * 而是「与 Word 的断行点一致」，而 Word 用的是自己那套东亚版式规则 —— 两者在中文文本上
 * 的差异恰恰集中在禁则与标点挤压，也就是下面这两张表。
 */
import type { LayoutItem } from './types.ts';

/**
 * 内建**首禁则**集：不能出现在行首的字符（后置标点）。
 *
 * 取自 Word「中文版式 → 首尾字符设置」的「标准」级别。文档可以用
 * `w:noLineBreaksBefore` 整个替换它（见 `kinsokuFrom`）。
 *
 * ⚠️ 这张表是照 Word 界面抄的，**没有真值验证**。上 Windows 时值得做一份
 * 「每个候选字符在行首会不会被推下去」的样本把边界钉死 —— 尤其是
 * `—`（破折号）、`…`（省略号）、`％`、`℃` 这几个各家实现分歧最大的。
 */
const DEFAULT_NO_START = '!),.:;?]}¢°\'"′″‰℃、。〃々〉》」』】〕！＂％＇），．：；？］｀｜｝～￠…—';

/** 内建**尾禁则**集：不能出现在行尾的字符（前置标点） */
const DEFAULT_NO_END = '$([{£¥·‘“〈《「『【〔（［｛＄＇＜＠｀￡￥';

/**
 * 全角标点的**空半边在哪一侧** —— 挤压的物理依据，也是「哪两个标点之间能挤」的判据。
 *
 * 只收**中文全角**标点：拉丁标点本来就是紧排的，压它会让英文变形。
 *
 * - `BLANK_LEFT`（开口类）：墨在右半边，空的是左半边 —— 「『（〔【《〈 与左引号
 * - `BLANK_RIGHT`（收口与句读）：墨在左半边，空的是右半边 —— 、。，．：；！？ 与右括号
 * - 省略号「…」与破折号「—」两边都不空（墨横贯整个字宽），**自己不可压**，
 *   但它仍是标点：邻居的空半边挨着它照样能压（见 `punctPairCompressible`）
 */
const BLANK_LEFT = '「『（〔【《〈“‘';
const BLANK_RIGHT = '、。，．：；！？」』）〕】》〉”’';
/** 两边都不空的全角标点。它们只作为**邻居**参与挤压判定 */
const NO_BLANK = '…‥—―';

export interface KinsokuSets {
  noStart: ReadonlySet<number>;
  noEnd: ReadonlySet<number>;
}

function setOf(chars: string): Set<number> {
  const out = new Set<number>();
  for (const ch of chars) out.add(ch.codePointAt(0) as number);
  return out;
}

export const DEFAULT_KINSOKU: KinsokuSets = {
  noStart: setOf(DEFAULT_NO_START),
  noEnd: setOf(DEFAULT_NO_END),
};

const BLANK_LEFT_SET = setOf(BLANK_LEFT);
const BLANK_RIGHT_SET = setOf(BLANK_RIGHT);
const NO_BLANK_SET = setOf(NO_BLANK);

/**
 * 用 `settings.xml` 的自定义禁则集覆盖内建表。
 *
 * **整个替换**而不是追加：Word 界面上那就是一个可编辑的完整列表，用户删掉的字符
 * 必须真的消失。空串表示文件里没写，此时用内建表。
 */
export function kinsokuFrom(settings: {
  noLineBreaksBefore: string;
  noLineBreaksAfter: string;
}): KinsokuSets {
  return {
    // w:noLineBreaksBefore =「这些字符**之前**不许断」= 它们不能出现在行首
    noStart:
      settings.noLineBreaksBefore === '' ? DEFAULT_KINSOKU.noStart : setOf(settings.noLineBreaksBefore),
    noEnd: settings.noLineBreaksAfter === '' ? DEFAULT_KINSOKU.noEnd : setOf(settings.noLineBreaksAfter),
  };
}

/** 列表编号的 item（编号文字与它后面那个分隔符），见 types.ts 的 `CharItem.numbering` */
export function isNumberingItem(item: LayoutItem): boolean {
  return (item.kind === 'char' || item.kind === 'tab') && item.numbering === true;
}

export function kinsokuOf(cp: number, sets: KinsokuSets = DEFAULT_KINSOKU): 'none' | 'noStart' | 'noEnd' {
  if (sets.noStart.has(cp)) return 'noStart';
  if (sets.noEnd.has(cp)) return 'noEnd';
  return 'none';
}

/** 有空半边可以交出去的全角标点。省略号与破折号**不**在内（墨横贯整个字宽） */
export function isCompressiblePunct(cp: number): boolean {
  return BLANK_LEFT_SET.has(cp) || BLANK_RIGHT_SET.has(cp);
}

/**
 * 这两个相邻的字符之间要不要挤掉半个字 —— **两条都要成立**：
 *
 * 1. 两边都是全角标点。标点挨着汉字一点都不压（`spike-punct-01`，26 段实测）
 * 2. 这个**接缝上有空白**：前一个是收口类（空在右），或后一个是开口类（空在左）。
 *    唯一压不了的组合是「开口紧跟收口」（`「，`）—— 开口的墨在右半边、收口的墨在左半边，
 *    接缝上两边都是墨，挤它就是把字形叠起来
 *
 * 第 2 条是从 `gongwen-01` 真值第 10 行（0 起）反推的，那一行把整套标点排成一串：
 * Word 在 `，。`…`）】`…`】…` 这 12 个接缝上各挤了半个字，唯独 `「，` 那个接缝
 * **一点没挤**（前 11 个字的推进宽 163.29pt，只够容下一次挤压加一点行内调整）。
 * 同一行还说明**省略号也参与**：`】…` 那个接缝挤了 —— 挤掉的是「】」的右半边，
 * 与「…」自己压不压无关，所以判定要按「接缝上有没有空白」而不是「两边都可压」。
 */
export function punctPairCompressible(prev: number, next: number): boolean {
  if (!isPunctCp(prev) || !isPunctCp(next)) return false;
  return BLANK_RIGHT_SET.has(prev) || BLANK_LEFT_SET.has(next);
}

function isPunctCp(cp: number): boolean {
  return isCompressiblePunct(cp) || NO_BLANK_SET.has(cp);
}

/**
 * **相邻两个全角标点**要压掉的宽度，单位 em（东亚一侧的字号）。
 *
 * 这个 0.5 是实测的，不在 `uncalibrated.ts` 里 —— 样本 `spike-punct-01`
 * （26 段短句，每段不折行，行宽减去自然宽就是压缩量，与断行 / 悬挂 / 行距全无关）：
 *
 * | 段落 | 压掉 | 说明 |
 * |---|---|---|
 * | 甲，乙 / 甲。乙 / 甲、乙 / 甲；乙 / 甲：乙 / 甲？乙 / 甲！乙 | **0.000 em** | 孤立的标点**一点都不压** |
 * | 甲（乙）丙 / 甲「乙」丙 / 甲《乙》丙 | 0.000 em | 开口与收口各自孤立，同样不压 |
 * | 甲，，乙 / 甲、、乙 / 甲），乙 / 甲。」乙 | **0.504 em** | 相邻一对，压掉半个字 |
 * | 甲，（乙）丙 | 0.504 em | 收口 + 开口中间空着整整 1 em，**也只压 0.5** |
 * | 甲（「乙」）丙 | 1.000 em | 两对，各 0.5 |
 * | 甲，，，乙 | 1.004 em | 三连 = 两对 |
 * | （甲乙）丙 / 甲乙（丙 / 甲乙） | 0.000 em | 行首、行末、紧邻汉字，都不压 |
 *
 * 一句话：**只有「标点紧跟标点」这一种情形压，且固定压半个字**。
 * 这条与「塞不下时临时挤一点」是两件事 —— 后者是下面的 `PUNCT_COMPRESS_MAX_EM`
 * 与 `PUNCT_COMPRESS_STRETCH_K`，各有自己的样本（`spike-compress-01/02`）。
 */
export const PUNCT_PAIR_COMPRESS_EM = 0.5;

/**
 * 悬挂出版心的标点，有多少**留在版心里**（占它自己宽度的比例）。
 *
 * 这个 0.5 也是实测的：`gongwen-01` 真值第 4 与第 13 行（行号 0 起）的行尾「，」，
 * 左边缘落在版心线**内** 7.96pt、右边缘出界 8.05pt（16pt 字号，即 15.96pt 的推进宽），
 * 正好一半一半 —— **吐出版心的是空的那半边，墨留在里面**。
 *
 * | 行 | 版心右边缘 | 「，」右边缘 | 出界 |
 * |---|---|---|---|
 * | 第 4 行 | 521.594 pt | 529.643 pt | 8.049 pt |
 * | 第 13 行 | 521.594 pt | 529.520 pt | 7.926 pt |
 *
 * 直接后果有两个，都在 `linebreak.ts`：① 能不能挂，看的是**半宽**塞不塞得下，
 * 塞不下要先挤压；② 行宽要把这半个字算进去，否则两端对齐会把整行多拉半个字宽 ——
 * 实测第 4 行的「，」左边缘落在 434.30pt，正是「版心宽 442.25 − 半个字 8」。
 *
 * ⚠️ 只对**全角**后置标点有真值。行尾空格仍然整个不计入行宽（那是另一条老规则，
 * 见 `contentWidth`）；拉丁标点（`.` `,`）会不会悬挂、悬挂多少，没有样本。
 */
export const HANG_INSIDE_RATIO = 0.5;

/**
 * **塞不下时**临时挤一个标点的上限，单位 em —— 它自己的空半边，不是 0.5 而是 0.48。
 *
 * 实测 `spike-compress-02` 的 G1 组（20 字一行、行内只有一个孤立的「，」、两端对齐，
 * 亏空从 7.4pt 走到 8.4pt，步长 0.2pt）：7.60pt 时还肯挤（0.4753 em），7.80pt 就换行了
 * （0.4878 em）。取中点 0.48。
 *
 * 与 `PUNCT_PAIR_COMPRESS_EM`（0.5）差的那 0.02 em 是实测差异，不是同一个数的两次测量：
 * 常态的「标点紧跟标点」压的是**两个字之间**的空白，临时挤压动的是**一个标点自己**的空半边。
 */
export const PUNCT_COMPRESS_MAX_EM = 0.48;

/**
 * 「挤一个标点」与「拉一个字距」的兑换率 —— 决定 Word 什么时候**宁可换行也不再挤**。
 *
 * 这条是 `spike-compress-02` 五组阶梯（标点数 1/2/3/4/6 × 行长 14/20/27 字，
 * 刻度 0.1–0.4pt）反推出来的。判据写成一个不等式：
 *
 * ```
 * 挤压量 × 字距数  ≤  K × 标点数 × 换行后要拉伸的量
 * ```
 *
 * 也就是「每个标点挨的挤压」不超过「每个字距挨的拉伸」的 K 倍。K = 30.6 时七组阶梯里
 * 六组的翻转点都落在预测值的 ±0.1pt 内（第七组 P=1 由上面的 0.48 em 上限先卡住）：
 *
 * | 组 | 标点数 | 行长 | 实测翻转点 | 预测 |
 * |---|---|---|---|---|
 * | G1 | 1 | 20 字 | 7.60 → 7.80 pt | 上限 0.48 em = 7.67 |
 * | G2 | 2 | 20 字 | 12.20 → 12.30 | 12.20 |
 * | G3 | 3 | 20 字 | 13.00 → 13.25 | 13.24 |
 * | G4 | 4 | 20 字 | 13.75 → 13.90 | 13.84 |
 * | G5 | 6 | 20 字 | 14.20 → 14.60 | 14.49 |
 * | G6 | 2 | 14 字 | 13.00 → 13.50 | 13.19 |
 * | G7 | 2 | 27 字 | 11.20 → 11.30 | 11.22 |
 *
 * 同一批样本还钉死了两条边界：**标点在行首还是行末不影响**（G8 / G9 与 G2 同格翻转），
 * **段落后面还有没有文字也不影响**（G10 与 G2 同格）。
 *
 * ⚠️ K = 30.6 这个数**不圆**，说明公式的形状多半只是 Word 真规则的一个好近似。
 * **已知的一个反例**：gongwen-01 真值第 10 行 —— 那一行 Word 只差 4.6pt 就能留住「出」，
 * 行内还有四个孤立标点给得起（按这个不等式也该挤），却换了行。那一行的特别之处是
 * 行内有一串 16 个连着的标点、12 个接缝已经在常态挤压里各压掉了半个字；
 * 换句话说「已经压过的行还肯不肯再压」多半另有规则，这份样本没覆盖到。
 * 它是 L2 剩下 2 行（真值第 10 / 11 行）对不上的唯一原因。
 */
export const PUNCT_COMPRESS_STRETCH_K = 30.6;

/** 空白：断点在它之后，且行尾的它不计入行宽（可以吐出版心） */
export function isSpaceCp(cp: number): boolean {
  return cp === 0x20 || cp === 0x09 || cp === 0xa0 || cp === 0x3000 || cp === 0x2007 || cp === 0x202f;
}

/** 连字符之后允许断行 */
function isHyphenCp(cp: number): boolean {
  return cp === 0x2d || cp === 0x2010 || cp === 0x2012 || cp === 0x2013 || cp === 0x2014;
}

/**
 * `next` 能不能作为一行的开头 —— 也就是「能不能在 prev 和 next 之间断」。
 *
 * 禁则在这里生效，而不是等断完再修：把它做成断点判定的一部分，「压不下就回退到
 * 上一个断点」自然就成立了（linebreak.ts 往回找的是**允许的**断点），不必再写一遍回退逻辑。
 */
export function canBreakBetween(prev: LayoutItem | undefined, next: LayoutItem): boolean {
  if (prev === undefined) return false;
  // 硬换行由 linebreak.ts 直接处理，不走这里
  if (prev.kind === 'break' || next.kind === 'break') return false;
  // 列表编号永远不能作为行首 —— 这一条同时管住「编号内部不许断」与「编号必须留在首行」：
  // 编号 item 全在段落最前面，它们既然不能开新行，就只能整体待在第 0 行
  if (isNumberingItem(next)) return false;

  if (prev.kind === 'char') {
    if (prev.noBreak === true) return false; // w:noBreakHyphen
    if (prev.kinsoku === 'noEnd') return false; // 前置标点不能留在行尾
  }
  if (next.kind === 'char') {
    if (next.softHyphen === true) return true; // 软连字符就是为了在这里断
    if (next.kinsoku === 'noStart') return false; // 后置标点不能跑到行首
    if (next.space) return false; // 在空格前断毫无意义，空格跟着上一行走
  }

  if (prev.kind === 'char' && prev.space) return true;
  // 制表位与内嵌对象两侧都可断：它们本来就是「一整块」，不属于任何单词
  if (prev.kind !== 'char' || next.kind !== 'char') return true;

  if (prev.script === 'eastAsia' || next.script === 'eastAsia') return true;
  if (isHyphenCp(prev.cp)) return true;
  // 拉丁词内不断。自动断词（w:autoHyphenation）默认关，中文公文里也不该开
  return false;
}
