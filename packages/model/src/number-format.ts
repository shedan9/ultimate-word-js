/**
 * 计数值 → 编号文字（`w:numFmt` + `w:lvlText`）。
 *
 * 纯函数、无状态：进来是「第几级各是几」，出去是一串字符。计数器的推进在
 * numbering-counter.ts，级联在 cascade.ts —— 三件事分开是因为只有这一件能被
 * 单测钉死到字符级，另外两件依赖文档顺序与样式表。
 *
 * **降级优于丢失**：规范列了六十多种 `w:numFmt`，认不出的一律按 `decimal` 出数字。
 * 编号错一位是瑕疵，编号消失是 bug —— 用户会以为文档内容丢了。
 *
 * ── 未标定（拿到 Word 真值前不要当结论）────────────────────────────────────
 * 1. `chineseCounting` 与 `chineseCountingThousand` 在这里是**同一套读法**。
 *    两者在小数值上完全重合（十一、二十三），差异只可能出现在 ≥ 100 的位值读法上，
 *    而公文列表极少编到三位数。钉死它的样本：一份 numFmt 分别取这两个值、
 *    编号跑到 105 / 1005 的 docx，看 Word 显示「一百零五」还是别的。
 * 2. `ideographDigital` 同样按中文读法出（一、二、十一）。另一种可能是逐位念
 *    （一〇五），同一份样本能一起分开。
 * 3. 零的插入规则（105 → 一百零五）是按中文习惯写的，没有 Word 真值。
 */

/** 阿拉伯数字之外的几套字表 —— 都是「值 → 字符」的死表，没有算法可言 */
const CN_DIGITS = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九'] as const;
const CN_UNITS = ['', '十', '百', '千'] as const;
const CN_BIG = ['', '万', '亿'] as const;
const CN_LEGAL_DIGITS = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖'] as const;
const CN_LEGAL_UNITS = ['', '拾', '佰', '仟'] as const;
/** 天干，`ideographTraditional` 用的就是它，只有 10 个 */
const HEAVENLY_STEMS = ['甲', '乙', '丙', '丁', '戊', '己', '庚', '辛', '壬', '癸'] as const;
/** 地支，`ideographZodiac` 用，12 个 */
const EARTHLY_BRANCHES = ['子', '丑', '寅', '卯', '辰', '巳', '午', '未', '申', '酉', '戌', '亥'] as const;

const ROMAN: readonly (readonly [number, string])[] = [
  [1000, 'M'],
  [900, 'CM'],
  [500, 'D'],
  [400, 'CD'],
  [100, 'C'],
  [90, 'XC'],
  [50, 'L'],
  [40, 'XL'],
  [10, 'X'],
  [9, 'IX'],
  [5, 'V'],
  [4, 'IV'],
  [1, 'I'],
];

/**
 * 一个计数值按 `w:numFmt` 出文字。
 *
 * 不认识的 `numFmt`、以及超出该格式表示范围的值（罗马数字 > 3999、带圈数字 > 20、
 * 天干 > 10），一律回退成十进制 —— 见文件头「降级优于丢失」。
 */
export function formatNumber(value: number, numFmt: string): string {
  // 负数与 0 只有 decimal 说得清（Word 里 startOverride 可以写 0）
  if (!Number.isFinite(value) || value <= 0) return String(Math.trunc(value));

  switch (numFmt) {
    case 'decimal':
    case 'decimalHalfWidth':
      return String(value);
    case 'decimalZero':
      // 只补到两位：Word 的 decimalZero 是「个位数补一个 0」，不是补到固定宽度
      return value < 10 ? `0${value}` : String(value);
    case 'decimalFullWidth':
    case 'decimalFullWidth2':
      return toFullWidth(value);
    case 'upperLetter':
      return repeatedLetter(value, 'A');
    case 'lowerLetter':
      return repeatedLetter(value, 'a');
    case 'upperRoman':
      return toRoman(value) ?? String(value);
    case 'lowerRoman':
      return toRoman(value)?.toLowerCase() ?? String(value);
    case 'ordinal':
      return `${value}${ordinalSuffix(value)}`;
    case 'decimalEnclosedCircle':
    case 'decimalEnclosedCircleChinese':
      return enclosed(value, 0x2460) ?? String(value);
    case 'decimalEnclosedParen':
      return enclosed(value, 0x2474) ?? String(value);
    case 'decimalEnclosedFullstop':
      return enclosed(value, 0x2488) ?? String(value);
    case 'chineseCounting':
    case 'chineseCountingThousand':
    case 'ideographDigital':
      return chineseNumber(value, CN_DIGITS, CN_UNITS, true);
    case 'chineseLegalSimplified':
      // 法定大写不省略「壹拾」的那个「壹」—— 支票上的写法，省了就变造假
      return chineseNumber(value, CN_LEGAL_DIGITS, CN_LEGAL_UNITS, false);
    case 'ideographTraditional':
      return value <= HEAVENLY_STEMS.length ? (HEAVENLY_STEMS[value - 1] as string) : String(value);
    case 'ideographZodiac':
    case 'ideographZodiacTraditional':
      return value <= EARTHLY_BRANCHES.length ? (EARTHLY_BRANCHES[value - 1] as string) : String(value);
    case 'none':
      return '';
    default:
      // ordinalText / cardinalText（one、first）与各国语言的计数格式都落这儿。
      // 拼英文数词要一张不小的表，而中文公文里它们不出现，先降级
      return String(value);
  }
}

/**
 * `w:lvlText` 模板展开：`%1.%2` → `1.3`。
 *
 * `%n` 里的 n 是**1 起的级别号**（`%1` 是 ilvl 0），这是 OOXML 里最容易差一位的地方。
 * 引用到还没有值的级（多级列表跳级用）时按该级 start 补，由调用方在 `values` 里填好。
 *
 * `%` 后面不是 1–9 的一律原样保留 —— 有的模板里 `%` 就是字面百分号。
 */
export function formatLevelText(
  lvlText: string,
  values: readonly number[],
  formats: readonly string[],
): string {
  let out = '';
  for (let i = 0; i < lvlText.length; i++) {
    const ch = lvlText[i] as string;
    if (ch !== '%') {
      out += ch;
      continue;
    }
    const digit = lvlText.charCodeAt(i + 1) - 0x30;
    if (!(digit >= 1 && digit <= 9)) {
      out += ch;
      continue;
    }
    const value = values[digit - 1];
    // 引用了一个连 start 都补不出来的级：整个占位符吞掉，别把「%3」印到纸上
    out += value === undefined ? '' : formatNumber(value, formats[digit - 1] ?? 'decimal');
    i++;
  }
  return out;
}

function toFullWidth(value: number): string {
  let out = '';
  for (const ch of String(value)) out += String.fromCharCode(ch.charCodeAt(0) - 0x30 + 0xff10);
  return out;
}

/**
 * `upperLetter` 超过 Z 之后**重复字母**（27 = AA、28 = BB），不是 Excel 那种进位（AA、AB）。
 * 这是 Word 的实际行为，按进位实现会从第 27 项起全错。
 */
function repeatedLetter(value: number, base: 'A' | 'a'): string {
  const index = (value - 1) % 26;
  const times = Math.floor((value - 1) / 26) + 1;
  return String.fromCharCode(base.charCodeAt(0) + index).repeat(times);
}

function toRoman(value: number): string | undefined {
  if (value > 3999) return undefined;
  let n = value;
  let out = '';
  for (const [v, s] of ROMAN) {
    while (n >= v) {
      out += s;
      n -= v;
    }
  }
  return out;
}

function ordinalSuffix(value: number): string {
  const rem100 = value % 100;
  if (rem100 >= 11 && rem100 <= 13) return 'th';
  const rem10 = value % 10;
  return rem10 === 1 ? 'st' : rem10 === 2 ? 'nd' : rem10 === 3 ? 'rd' : 'th';
}

/** 带圈 / 带括号数字都是从 1 开始的连续 20 个码位，超出就没有字形了 */
function enclosed(value: number, base: number): string | undefined {
  return value <= 20 ? String.fromCodePoint(base + value - 1) : undefined;
}

/**
 * 中文数字读法。
 *
 * `elide` 控制「一十一 → 十一」那个省略：口语读法省，法定大写不省。
 * 省略只在**整个数小于 20** 时发生 —— 110 读「一百一十」，不读「一百十」。
 */
function chineseNumber(
  value: number,
  digits: readonly string[],
  units: readonly string[],
  elide: boolean,
): string {
  // 亿以上就没有稳定读法了（万亿 / 兆各家不同），直接给数字，不猜
  if (value >= 1e12) return String(value);

  const groups: number[] = [];
  let rest = value;
  while (rest > 0) {
    groups.push(rest % 10000);
    rest = Math.floor(rest / 10000);
  }

  let out = '';
  for (let g = groups.length - 1; g >= 0; g--) {
    const group = groups[g] as number;
    if (group === 0) continue;
    // 低位组不足四位时补「零」：1000001 读「一百万零一」，不读「一百万一」
    if (out !== '' && group < 1000) out += digits[0];
    out += fourDigits(group, digits, units) + (CN_BIG[g] ?? '');
  }
  if (out === '') return digits[0] as string;
  if (elide && value >= 10 && value < 20) return out.slice(1);
  return out;
}

/** 0–9999 的读法：逢 0 补一个「零」，末尾的 0 不读 */
function fourDigits(value: number, digits: readonly string[], units: readonly string[]): string {
  let out = '';
  let zeroPending = false;
  const s = String(value);
  for (let i = 0; i < s.length; i++) {
    const d = s.charCodeAt(i) - 0x30;
    const unit = units[s.length - 1 - i] as string;
    if (d === 0) {
      // 连续的 0 只读一个「零」，且要等到后面真有非 0 位才补上
      zeroPending = out !== '';
      continue;
    }
    if (zeroPending) {
      out += digits[0];
      zeroPending = false;
    }
    out += (digits[d] as string) + unit;
  }
  return out;
}
