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
 * 可挤压的全角标点：字形只占半边，塞不下时可以压掉空着的那半。
 *
 * 只收**中文全角**标点：拉丁标点本来就是紧排的，压它会让英文变形。
 * 挤压量见 uncalibrated.ts 的 `PUNCT_COMPRESS_RATIO`。
 */
const COMPRESSIBLE = '、。，．：；！？「」『』（）〔〕【】《》〈〉“”‘’';

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

const COMPRESSIBLE_SET = setOf(COMPRESSIBLE);

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

export function kinsokuOf(cp: number, sets: KinsokuSets = DEFAULT_KINSOKU): 'none' | 'noStart' | 'noEnd' {
  if (sets.noStart.has(cp)) return 'noStart';
  if (sets.noEnd.has(cp)) return 'noEnd';
  return 'none';
}

export function isCompressiblePunct(cp: number): boolean {
  return COMPRESSIBLE_SET.has(cp);
}

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
