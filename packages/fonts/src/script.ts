/**
 * `w:rFonts` 的脚本分桶 —— 中文文档度量出错的头号原因。
 *
 * `w:rFonts` **不是**给一个 run 指定一款字体，而是同时挂四款：
 *
 * ```xml
 * <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"
 *           w:eastAsia="仿宋" w:cs="Times New Roman" w:hint="eastAsia"/>
 * ```
 *
 * 引擎必须**逐字符**判断它属于哪个桶，再取对应字体的度量 ——「汉字用仿宋、数字英文用
 * Times New Roman」就是这么来的。把整个 run 按一款字体量，公文里每一行的宽度都会错。
 * 因此**一个 run 内可以横跨多款字体**，度量与断行的最小单位是 `FontRun`，不是 run。
 *
 * 放在 `@uw/fonts` 而不是 `@uw/layout`：这是**字体选择策略**，不是排版算法（架构 §3.1）。
 * layout 拿到的是已经切好的 `FontRun[]`，断行算法不必知道 rFonts 是什么。
 *
 * 这个文件不 import `@uw/model` —— 依赖方向是 `layout → {model, fonts}`，
 * fonts 不认识 model。`ScriptFonts` 与 `ResolvedRunProps['fonts']` 结构上兼容，
 * 调用方直接传即可。
 */

/** 四个 rFonts 属性，也就是四个桶 */
export type FontBucket = 'ascii' | 'hAnsi' | 'eastAsia' | 'cs';

/** `w:hint`：歧义字符归哪个桶的决断依据 */
export type FontHint = 'default' | 'eastAsia' | 'cs';

/**
 * 歧义字符（Unicode EastAsianWidth = **A**）归哪个桶，按什么判 —— **标定用的接缝**，
 * 正常调用不要传，实测结论与证据表在 `@uw/layout` 的 `WIDTH_RULES`：
 *
 * - `hint`：`w:hint="eastAsia"` 时进 eastAsia 桶，否则进 hAnsi 桶（**实测**）
 * - `eastAsia` / `latin`：不看 hint，一律进那一个桶
 */
export type AmbiguousRule = 'hint' | 'eastAsia' | 'latin';

/**
 * 空格这类**中性字符**跟不跟东亚邻居走 —— 同样是标定用的接缝，见 `neutralTakesEastAsia`：
 *
 * - `either`：任一侧邻居是东亚字就跟（**实测**）
 * - `eitherHinted`：同上，但还要 `w:hint="eastAsia"`（**旧实现**）
 * - `prev`：只看前一个字
 * - `none`：一律按 `bucketOf` 走
 */
export type NeutralRule = 'either' | 'eitherHinted' | 'prev' | 'none';

/**
 * 粗分类，给**排版**用（与 `FontBucket` 的用途不同）：
 * - 行高要不要走东亚的 1.3 系数，看 `eastAsia`
 * - 中西文之间自动加 1/4 em（实测），靠的是 `latin` ↔ `eastAsia` 的边界
 *
 * ascii 与 hAnsi 都是拉丁，排版规则相同，所以合并成 `latin`。
 */
export type ScriptKind = 'latin' | 'eastAsia' | 'complex';

/** 结构上与 `ResolvedRunProps['fonts']` 一致；空串表示「文档没指定这个桶」 */
export interface ScriptFonts {
  ascii: string;
  hAnsi: string;
  eastAsia: string;
  cs: string;
  hint: FontHint;
}

/**
 * 一段「同一款字体、同一种排版脚本」的连续字符。
 *
 * `start` / `end` 是**源字符串的 UTF-16 下标**（可直接 `text.slice(start, end)`），
 * 不是码点序号 —— 布局要拿它回去切片、算命中区间，用码点序号会在 CJK 扩展 B（代理对）上错位。
 */
export interface FontRun {
  start: number;
  end: number;
  /** 解析后的字体名；空串表示这四个桶都没指定，由调用方套默认字体 */
  font: string;
  script: ScriptKind;
}

// ── 码点区间表 ────────────────────────────────────────────────────────────────
//
// 三张表互不相交，查表顺序不影响结果。写成 [lo, hi] 对，模块加载时摊平成
// 一维数组做二分 —— 逐字符要查几万次，线性扫会成为热点。

type Range = readonly [number, number];

/**
 * 无条件走 eastAsia 桶：Unicode EastAsianWidth = W / F 的区间。
 * 注意 U+3000–U+303F（全角标点：「」、。）在这里 —— 它们**不**由 hint 决断。
 */
const EAST_ASIA: readonly Range[] = [
  [0x1100, 0x115f], // 谚文字母
  [0x2e80, 0x2ef3], // 康熙部首补充
  [0x2f00, 0x2fd5], // 康熙部首
  [0x2ff0, 0x2ffb], // 汉字结构描述符
  [0x3000, 0x303e], // CJK 符号与标点（全角空格、「」、。）
  [0x3041, 0x3096], // 平假名
  [0x3099, 0x30ff], // 假名声调符 + 片假名
  [0x3105, 0x312f], // 注音符号
  [0x3131, 0x318e], // 谚文兼容字母
  [0x3190, 0x31e3], // 汉文训读 + 笔画
  [0x31f0, 0x321e], // 片假名语音扩展 + 带括号谚文
  [0x3220, 0x3247], // 带括号的表意文字
  [0x3250, 0x4dbf], // 扩充 A
  [0x4e00, 0xa48c], // 基本区 + 彝文
  [0xa490, 0xa4c6], // 彝文部首
  [0xa960, 0xa97c], // 谚文字母扩展 A
  [0xac00, 0xd7a3], // 谚文音节
  [0xf900, 0xfaff], // 兼容表意文字
  [0xfe10, 0xfe19], // 竖排标点
  [0xfe30, 0xfe52], // CJK 兼容形式
  [0xfe54, 0xfe66],
  [0xfe68, 0xfe6b],
  [0xff01, 0xff60], // 全角 ASCII（Ａ１，。）
  [0xffe0, 0xffe6], // 全角符号（￥ ￡）
  [0x1b000, 0x1b001],
  [0x1f200, 0x1f251],
  [0x20000, 0x3fffd], // 扩充 B 及以后
];

/**
 * 无条件走 cs 桶：需要复杂文字整形的书写系统。
 *
 * RTL 与复杂文字整形本身是**非目标**，但分桶必须认它们 —— 否则阿拉伯文会去查
 * `w:ascii` 的字体，度量整个错掉，页数跟着错。认桶不等于会整形。
 */
const COMPLEX: readonly Range[] = [
  [0x0590, 0x08ff], // 希伯来 → 阿拉伯 → 叙利亚 → 它拿字母 → 恩科
  [0x0900, 0x0dff], // 天城文 → 僧伽罗
  [0x0e00, 0x0eff], // 泰文 / 老挝文
  [0x0f00, 0x0fff], // 藏文
  [0x1000, 0x109f], // 缅甸文
  [0x1780, 0x17ff], // 高棉文
  [0xfb1d, 0xfdff], // 希伯来 / 阿拉伯表现形式 A
  [0xfe70, 0xfeff], // 阿拉伯表现形式 B
];

/**
 * 由 `w:hint` 决断的歧义字符：Unicode EastAsianWidth = **A**（Ambiguous）。
 *
 * 这不是拍脑袋列的 —— 「东亚环境下算全角、其他环境算半角」的字符集，Unicode 已经定义好了，
 * 而 `w:hint` 要回答的正是「这份文档算不算东亚环境」。
 *
 * 公文里真正咬人的几个都在这张表里：`①②③`(U+2460+)、`※`(U+203B)、`℃`(U+2103)、
 * `Ⅰ Ⅱ Ⅲ`(U+2160+)、`■ ●`(U+25A0+)、希腊字母、西里尔字母。hint="eastAsia" 时它们
 * 用中文字体（全角宽），否则用拉丁字体（半角宽）—— 一行里差几个字的宽度，断行点就变了。
 */
const AMBIGUOUS: readonly Range[] = [
  [0x00a1, 0x00a1],
  [0x00a4, 0x00a4],
  [0x00a7, 0x00a8],
  [0x00aa, 0x00aa],
  [0x00ad, 0x00ae],
  [0x00b0, 0x00b4],
  [0x00b6, 0x00ba],
  [0x00bc, 0x00bf],
  [0x00c6, 0x00c6],
  [0x00d0, 0x00d0],
  [0x00d7, 0x00d8],
  [0x00de, 0x00e1],
  [0x00e6, 0x00e6],
  [0x00e8, 0x00ea],
  [0x00ec, 0x00ed],
  [0x00f0, 0x00f0],
  [0x00f2, 0x00f3],
  [0x00f7, 0x00fa],
  [0x00fc, 0x00fc],
  [0x00fe, 0x00fe],
  [0x0101, 0x0101],
  [0x0111, 0x0111],
  [0x0113, 0x0113],
  [0x011b, 0x011b],
  [0x0126, 0x0127],
  [0x012b, 0x012b],
  [0x0131, 0x0133],
  [0x0138, 0x0138],
  [0x013f, 0x0142],
  [0x0144, 0x0144],
  [0x0148, 0x014b],
  [0x014d, 0x014d],
  [0x0152, 0x0153],
  [0x0166, 0x0167],
  [0x016b, 0x016b],
  [0x01ce, 0x01ce],
  [0x01d0, 0x01d0],
  [0x01d2, 0x01d2],
  [0x01d4, 0x01d4],
  [0x01d6, 0x01d6],
  [0x01d8, 0x01d8],
  [0x01da, 0x01da],
  [0x01dc, 0x01dc],
  [0x0251, 0x0251],
  [0x0261, 0x0261],
  [0x02c4, 0x02c4],
  [0x02c7, 0x02c7],
  [0x02c9, 0x02cb],
  [0x02cd, 0x02cd],
  [0x02d0, 0x02d0],
  [0x02d8, 0x02db],
  [0x02dd, 0x02dd],
  [0x02df, 0x02df],
  [0x0300, 0x036f], // 组合附加符号
  [0x0391, 0x03a1], // 希腊大写（Α–Ρ）
  [0x03a3, 0x03a9],
  [0x03b1, 0x03c1], // 希腊小写（α–ρ）
  [0x03c3, 0x03c9],
  [0x0401, 0x0401],
  [0x0410, 0x044f], // 西里尔
  [0x0451, 0x0451],
  [0x2010, 0x2010],
  [0x2013, 0x2016], // – — ‖
  [0x2018, 0x2019], // ' '
  [0x201c, 0x201d], // " "
  [0x2020, 0x2022],
  [0x2024, 0x2027], // … 在内
  [0x2030, 0x2030], // ‰
  [0x2032, 0x2033], // ′ ″
  [0x2035, 0x2035],
  [0x203b, 0x203b], // ※
  [0x203e, 0x203e],
  [0x2074, 0x2074],
  [0x207f, 0x207f],
  [0x2081, 0x2084],
  [0x20ac, 0x20ac], // €
  [0x2103, 0x2103], // ℃
  [0x2105, 0x2105],
  [0x2109, 0x2109],
  [0x2113, 0x2113],
  [0x2116, 0x2116], // №
  [0x2121, 0x2122],
  [0x2126, 0x2126],
  [0x212b, 0x212b],
  [0x2153, 0x2154],
  [0x215b, 0x215e],
  [0x2160, 0x216b], // Ⅰ–Ⅻ
  [0x2170, 0x2179], // ⅰ–ⅹ
  [0x2189, 0x2189],
  [0x2190, 0x2199], // ← ↑ → ↓
  [0x21b8, 0x21b9],
  [0x21d2, 0x21d2],
  [0x21d4, 0x21d4],
  [0x21e7, 0x21e7],
  [0x2200, 0x2200],
  [0x2202, 0x2203],
  [0x2207, 0x2208],
  [0x220b, 0x220b],
  [0x220f, 0x220f],
  [0x2211, 0x2211],
  [0x2215, 0x2215],
  [0x221a, 0x221a],
  [0x221d, 0x2220],
  [0x2223, 0x2223],
  [0x2225, 0x2225],
  [0x2227, 0x222c],
  [0x222e, 0x222e],
  [0x2234, 0x2237],
  [0x223c, 0x223d],
  [0x2248, 0x2248],
  [0x224c, 0x224c],
  [0x2252, 0x2252],
  [0x2260, 0x2261], // ≠ ≡
  [0x2264, 0x2267], // ≤ ≥
  [0x226a, 0x226b],
  [0x226e, 0x226f],
  [0x2282, 0x2283],
  [0x2286, 0x2287],
  [0x2295, 0x2295],
  [0x2299, 0x2299],
  [0x22a5, 0x22a5],
  [0x22bf, 0x22bf],
  [0x2312, 0x2312],
  [0x2460, 0x24e9], // ①②③ 与 ⑴⑵⑶
  [0x24eb, 0x254b],
  [0x2550, 0x2573], // 制表符
  [0x2580, 0x258f],
  [0x2592, 0x2595],
  [0x25a0, 0x25a1], // ■ □
  [0x25a3, 0x25a9],
  [0x25b2, 0x25b3], // ▲ △
  [0x25b6, 0x25b7],
  [0x25bc, 0x25bd],
  [0x25c0, 0x25c1],
  [0x25c6, 0x25c8],
  [0x25cb, 0x25cb], // ○
  [0x25ce, 0x25d1], // ◎ ●
  [0x25e2, 0x25e5],
  [0x25ef, 0x25ef],
  [0x2605, 0x2606], // ★ ☆
  [0x2609, 0x2609],
  [0x260e, 0x260f],
  [0x2614, 0x2615],
  [0x261c, 0x261c],
  [0x261e, 0x261e],
  [0x2640, 0x2640],
  [0x2642, 0x2642],
  [0x2660, 0x2661],
  [0x2663, 0x2665],
  [0x2667, 0x266a],
  [0x266c, 0x266d],
  [0x266f, 0x266f],
  [0x269e, 0x269f],
  [0x26be, 0x26bf],
  [0x26c4, 0x26cd],
  [0x26cf, 0x26e1],
  [0x26e3, 0x26e3],
  [0x26e8, 0x26ff],
  [0x273d, 0x273d],
  [0x2757, 0x2757],
  [0x2776, 0x277f], // ❶❷❸
  [0x2b55, 0x2b59],
  [0x3248, 0x324f],
  [0xe000, 0xf8ff], // 私用区：w:sym 的 Wingdings 符号落在这里
  [0xfffd, 0xfffd],
];

/** [lo, hi] 对摊平成一维，二分查找 */
function flatten(ranges: readonly Range[]): Int32Array {
  const out = new Int32Array(ranges.length * 2);
  for (let i = 0; i < ranges.length; i++) {
    const r = ranges[i] as Range;
    out[i * 2] = r[0];
    out[i * 2 + 1] = r[1];
  }
  return out;
}

const EAST_ASIA_FLAT = flatten(EAST_ASIA);
const COMPLEX_FLAT = flatten(COMPLEX);
const AMBIGUOUS_FLAT = flatten(AMBIGUOUS);

function inRanges(flat: Int32Array, cp: number): boolean {
  let lo = 0;
  let hi = flat.length / 2 - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (cp < (flat[mid * 2] as number)) hi = mid - 1;
    else if (cp > (flat[mid * 2 + 1] as number)) lo = mid + 1;
    else return true;
  }
  return false;
}

/** 该码点是否是东亚文字 —— 行高走不走 1.3 系数由它决定（见 metrics.ts） */
export function isEastAsianCodePoint(cp: number): boolean {
  return inRanges(EAST_ASIA_FLAT, cp);
}

/**
 * 这段文字里有没有东亚文字。
 *
 * ⚠️ **不要拿它判行高走哪一套规则** —— 那一问的答案是「实际画字的那款字体是不是东亚字体」，
 * 与字符无关（实测，见 `@uw/layout` 的 `SCRIPT_RULES`）。`@uw/layout` 只在**字体缺失**、
 * 问不出字体是什么的时候才退回这条路。留着它是因为兜底度量（`fallbackAdvance`）
 * 与保真度脚本还要按字符分全角半角。
 */
export function hasEastAsianText(text: string): boolean {
  for (const ch of text) {
    if (isEastAsianCodePoint(ch.codePointAt(0) as number)) return true;
  }
  return false;
}

/**
 * 码点 → 四个 rFonts 桶中的哪一个。
 *
 * ASCII（< 0x80）**永远**走 ascii 桶，即使 hint="eastAsia" —— 中文文档里
 * `w:hint="eastAsia"` 是常态，若让它把数字和英文也拽进中文字体，「2024年」的
 * 「2024」就会变成全角宽。全角形式是**另外的码点**（U+FF10+），不需要靠 hint 分。
 *
 * 这一条 `spike-width-01` 顺手验过：Ea4 / Ea5 两段是 `hint="eastAsia"` 下的
 * `B§B°B±B`，四个 `B` 在真值里全是 TimesNewRomanPSMT 画的（`§` 才是宋体）。
 */
export function bucketOf(
  cp: number,
  hint: FontHint = 'default',
  ambiguous: AmbiguousRule = 'hint',
): FontBucket {
  if (cp < 0x80) return 'ascii';
  if (inRanges(EAST_ASIA_FLAT, cp)) return 'eastAsia';
  if (inRanges(COMPLEX_FLAT, cp)) return 'cs';
  if (inRanges(AMBIGUOUS_FLAT, cp)) {
    if (ambiguous === 'eastAsia') return 'eastAsia';
    if (ambiguous === 'latin') return 'hAnsi';
    if (hint === 'eastAsia') return 'eastAsia';
    if (hint === 'cs') return 'cs';
    return 'hAnsi';
  }
  return 'hAnsi';
}

/**
 * 空格这类**中性字符**要不要跟着东亚邻居走 —— 中英混排行宽算错的第二个原因（第一个是分桶）。
 *
 * 空格是 ASCII，按 `bucketOf` 一律进 ascii 桶，于是拿 Times New Roman 的 0.25 em 去量。
 * 真值说不是：`gongwen-01` 里 12 个空格，只要**任一侧的邻居是东亚字**，Word 量到的就是
 * 0.5 em（仿宋的空格宽），两侧都是拉丁字时才是 0.25 em。
 *
 * | 上下文 | Word 实测（16pt） | 谁的空格 |
 * |---|---|---|
 * | `以 Word` / `Word 导出` / `2026 年` / `年 8` | 7.95–8.23 pt | 仿宋（0.5 em） |
 * | `0.5 pt` 里那个 | ~3.76 pt | Times New Roman（0.25 em） |
 *
 * 注意**两侧都要看**：`Word 导出` 那个空格的前一个字是拉丁的 `d`，若只看前一个字就会
 * 判成 Times，实测却是 0.5 em。所以规则是「邻居里有一个东亚字就算东亚」，
 * 与 Unicode bidi 里中性字符随强方向的做法同构。
 *
 * ⚠️ **空格是唯一不能按真值里的字体名读的字符**：Word 画它时**不换 Tf**，
 * PDF 里它跟着前一个字的字体走，而**推进宽度**才是另一款字体的。
 * `spike-width-01` 的 Ea9（`B 中`）就是这样 —— 片段字体报的是 TimesNewRomanPSMT，
 * 宽度却是 18.14pt = 36pt 的 0.5 em（Times 自己的空格只有 0.25 em）。
 * 按字体名读会得出「空格只跟前一个字」这个相反的结论。
 *
 * `spike-width-01` 钉死的是**另外一半**：这条规则**与 `w:hint` 无关**。
 * 原来这里要求 `hint="eastAsia"`（公文的常态，猜的），实测 hint=default 的
 * De6（`中 中`）与 De7（`中 B`）里空格照样是 18.03pt = 0.5 em。
 *
 * 同一份样本还钉死了**只有空格**是这样：`/`（EaA / DeA）与 `-`（EaC / DeC）夹在两个
 * 汉字中间，两种 hint 下都是 Times 画的、宽 10.01 / 11.99pt（0.278 / 0.333 em），
 * 也就是老老实实待在 ascii 桶里。所以这个函数只该被空格调用。
 */
export function neutralTakesEastAsia(
  hint: FontHint,
  prev: ScriptKind | undefined,
  next: ScriptKind | undefined,
  rule: NeutralRule = 'either',
): boolean {
  switch (rule) {
    case 'none':
      return false;
    case 'prev':
      return prev === 'eastAsia';
    case 'eitherHinted':
      return hint === 'eastAsia' && (prev === 'eastAsia' || next === 'eastAsia');
    case 'either':
      return prev === 'eastAsia' || next === 'eastAsia';
  }
}

/** 桶 → 排版脚本类型。ascii / hAnsi 排版规则相同，合并成 latin */
export function scriptOfBucket(bucket: FontBucket): ScriptKind {
  if (bucket === 'eastAsia') return 'eastAsia';
  if (bucket === 'cs') return 'complex';
  return 'latin';
}

/**
 * 桶 → 真实字体名，带回退链。
 *
 * 某个桶是空串表示「文档没指定」，**不是错误** —— Word 此时会去别的桶找。
 * 回退顺序按「排版上最接近」排：东亚缺了先找拉丁（多数中文字体也画拉丁字符），
 * 拉丁缺了先找另一个拉丁桶。
 *
 * `hint="eastAsia"` 时 ascii 桶的回退里插一个 eastAsia：只写了 `w:eastAsia` 的
 * `w:rFonts` 在中文模板里很常见，此时 Word 用东亚字体画 ASCII，不去找系统默认字体。
 *
 * 四个桶全空时返回空串，由调用方套自己的默认字体 —— fonts 包不该替调用方决定默认值。
 */
export function bucketFont(fonts: ScriptFonts, bucket: FontBucket): string {
  const { ascii, hAnsi, eastAsia, cs, hint } = fonts;
  switch (bucket) {
    case 'ascii':
      return first(ascii, hint === 'eastAsia' ? eastAsia : '', hAnsi, cs);
    case 'hAnsi':
      return first(hAnsi, ascii, hint === 'eastAsia' ? eastAsia : '', cs);
    case 'eastAsia':
      return first(eastAsia, hAnsi, ascii, cs);
    case 'cs':
      return first(cs, ascii, hAnsi, eastAsia);
  }
}

function first(...names: readonly string[]): string {
  for (const n of names) {
    if (n !== '') return n;
  }
  return '';
}

/**
 * 把一段文字切成「同字体 + 同脚本」的连续段。
 *
 * 为什么不只按字体名合并：中西文之间要自动加 1/4 em 间距（实测，见 @uw/layout 的 WIDTH_RULES），靠的就是 latin ↔ eastAsia
 * 的边界。若「宋体」同时被两个桶引用而合并成一段，这个边界就没了，间距加不上去
 * —— 那是中文排版里肉眼可见的差异。
 *
 * 代理对按整个码点处理，不会在 U+20000 以上的汉字中间切开。
 */
export function splitFontRuns(
  text: string,
  fonts: ScriptFonts,
  ambiguous: AmbiguousRule = 'hint',
): FontRun[] {
  const out: FontRun[] = [];
  let i = 0;
  while (i < text.length) {
    const cp = text.codePointAt(i) as number;
    const width = cp > 0xffff ? 2 : 1;
    const bucket = bucketOf(cp, fonts.hint, ambiguous);
    const font = bucketFont(fonts, bucket);
    const script = scriptOfBucket(bucket);
    const last = out[out.length - 1];
    if (last !== undefined && last.font === font && last.script === script) {
      last.end = i + width;
    } else {
      out.push({ start: i, end: i + width, font, script });
    }
    i += width;
  }
  return out;
}
