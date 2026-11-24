/**
 * Word 字体名 → CSS `font-family`。
 *
 * 这是**渲染**的事，与度量无关：宽度、行高、基线全都已经由 `@uw/fonts` 按 Word 的那份
 * 度量算完了，这里只决定「用哪个字形把字画出来」。所以浏览器上装没装这款中文字体，
 * **不影响排版**，只影响好不好看 —— 这正是自研布局引擎买到的东西。
 *
 * 中文版 Word 存的是「仿宋」这种本地化名，Windows 上叫 `FangSong`，
 * macOS / Linux 上多半两个都没有，得退到 Noto / 思源那一族。所以每一项都是
 * 「本地化名 → 英文名 → 开源替代 → 通用族」的一串，与 `@uw/model` 的
 * `fontNameCandidates()`（那是**查度量**用的）同构但不是同一份表：
 * 那份表回答「Word 用的是哪款字体」，这份表回答「本机拿什么画」。
 */

/** 原字体缺失时，半角等宽的西文字形要能放进布局给的 0.5 em。 */
function cjkStack(names: string, substitutes: string, halfWidthAscii = false) {
  return {
    regular: `${names}, ${substitutes}`,
    ascii: halfWidthAscii ? `${names}, 'Noto Sans Mono CJK SC', Osaka-Mono, monospace` : undefined,
  };
}

const SERIF = "'Noto Serif CJK SC', 'Source Han Serif SC', serif";
const SANS = "'Noto Sans CJK SC', 'Source Han Sans SC', sans-serif";

/** 开发计划 §2.1 A 类（中文）的替代字体。等线与微软雅黑的 ASCII 使用比例宽度。 */
const CJK_STACKS = new Map([
  ['宋体', cjkStack("SimSun, '宋体'", SERIF, true)],
  ['仿宋', cjkStack("FangSong, '仿宋'", SERIF, true)],
  ['黑体', cjkStack("SimHei, '黑体'", SANS, true)],
  ['楷体', cjkStack("KaiTi, '楷体'", "'LXGW WenKai', 'Noto Serif CJK SC', serif", true)],
  ['等线', cjkStack("DengXian, '等线'", SANS)],
  ['微软雅黑', cjkStack("'Microsoft YaHei', '微软雅黑'", SANS)],
  ['新宋体', cjkStack("NSimSun, '新宋体'", "'Noto Serif CJK SC', serif", true)],
]);

/** 英文名也要认：同一款字体在 `w:rFonts` 里两种写法都出现得到 */
const ALIASES: Record<string, string> = {
  simsun: '宋体',
  nsimsun: '新宋体',
  fangsong: '仿宋',
  simhei: '黑体',
  kaiti: '楷体',
  dengxian: '等线',
  microsoftyahei: '微软雅黑',
  'microsoft yahei': '微软雅黑',
};

/** 拉丁字体按「它自己 + 同类通用族」退。度量早已定死，退到什么只影响字形 */
const LATIN_GENERIC: Record<string, string> = {
  'times new roman': 'serif',
  georgia: 'serif',
  cambria: 'serif',
  'courier new': 'monospace',
  consolas: 'monospace',
  arial: 'sans-serif',
  calibri: 'sans-serif',
  verdana: 'sans-serif',
  tahoma: 'sans-serif',
  'segoe ui': 'sans-serif',
  symbol: 'serif',
  wingdings: 'serif',
};

function cjkFamily(family: string) {
  const name = family.trim();
  return CJK_STACKS.get(ALIASES[name.toLowerCase()] ?? name);
}

/** 只替换半角 ASCII 片段的后备字体，中文和全角标点继续按各自字体栈绘制。 */
export function halfWidthAsciiFamily(family: string, text: string): string | undefined {
  const ascii = cjkFamily(family)?.ascii;
  return ascii !== undefined && /^[ -~]+$/.test(text) ? ascii : undefined;
}

/**
 * 默认的字体栈。字体名为空时退到 serif —— 空串在布局里表示「四个桶都没指定」，
 * 那时用的是 `defaultFont`，这里没有更多信息可用。
 */
export function defaultFontFamily(family: string): string {
  if (family === '') return 'serif';
  const key = family.toLowerCase();
  const cjk = cjkFamily(family)?.regular;
  if (cjk !== undefined) return cjk;
  const generic = LATIN_GENERIC[key] ?? 'serif';
  // 带空格 / 中文的名字必须加引号，否则 CSS 会把它当成多个关键字
  return `'${family.replace(/'/g, '')}', ${generic}`;
}
