/**
 * 脚本分桶。
 *
 * 这里的期望值来自规范（ECMA-376 §17.3.2.26）与 Unicode EastAsianWidth，
 * 不是从实现反推的 —— 分桶错了不会报错，只会让整篇文档的宽度悄悄偏掉，
 * 所以每条规则都要有一个说得出理由的测试钉着。
 */
import { describe, expect, it } from 'vitest';
import type { ScriptFonts } from './script.ts';
import {
  bucketFont,
  bucketOf,
  hasEastAsianText,
  isEastAsianCodePoint,
  neutralTakesEastAsia,
  splitFontRuns,
} from './script.ts';

const cp = (s: string): number => s.codePointAt(0) as number;

/**
 * 公文里最典型的一套：正文仿宋、西文 Times New Roman、hint 指向东亚。
 * 字体名跟着 `fixtures/gongwen-01.docx` 走 —— 那份真公文写的是「仿宋」。
 *
 * 别顺手改成 `仿宋_GB2312`：GB/T 9704 里确实常见，但它与「仿宋」是**两款字体、度量不同**，
 * 且在首批支持清单的「暂不支持」一栏（开发计划 §2.1）。测试 fixture 用未支持的字体名，
 * 会让后来者以为我们有它的度量。
 */
const gongwen: ScriptFonts = {
  ascii: 'Times New Roman',
  hAnsi: 'Times New Roman',
  eastAsia: '仿宋',
  cs: 'Times New Roman',
  hint: 'eastAsia',
};

describe('码点 → 桶', () => {
  it('ASCII 永远走 ascii 桶，hint=eastAsia 也不例外', () => {
    // 否则「2024年」的「2024」会被拽进中文字体变成全角宽
    expect(bucketOf(cp('2'), 'eastAsia')).toBe('ascii');
    expect(bucketOf(cp('A'), 'eastAsia')).toBe('ascii');
    expect(bucketOf(cp(' '), 'eastAsia')).toBe('ascii');
    expect(bucketOf(cp('.'), 'eastAsia')).toBe('ascii');
  });

  it('汉字、假名、全角标点无条件走 eastAsia，不由 hint 决断', () => {
    for (const hint of ['default', 'eastAsia', 'cs'] as const) {
      expect(bucketOf(cp('国'), hint)).toBe('eastAsia');
      expect(bucketOf(cp('。'), hint)).toBe('eastAsia'); // U+3002
      expect(bucketOf(cp('（'), hint)).toBe('eastAsia'); // U+FF08 全角括号
      expect(bucketOf(cp('あ'), hint)).toBe('eastAsia');
    }
  });

  it('EastAsianWidth=Ambiguous 的字符由 hint 决断 —— 公文里咬人的就是这批', () => {
    const ambiguous = ['①', '※', '℃', 'Ⅰ', '■', '●', '★', 'α', 'Б', '—', '“'];
    for (const ch of ambiguous) {
      expect(bucketOf(cp(ch), 'eastAsia'), ch).toBe('eastAsia');
      expect(bucketOf(cp(ch), 'default'), ch).toBe('hAnsi');
    }
    // hint=cs 时同一批字符归复杂文字桶
    expect(bucketOf(cp('①'), 'cs')).toBe('cs');
  });

  it('阿拉伯 / 希伯来 / 泰文走 cs —— 不会整形不代表可以分错桶', () => {
    expect(bucketOf(cp('ا'), 'eastAsia')).toBe('cs');
    expect(bucketOf(cp('א'), 'default')).toBe('cs');
    expect(bucketOf(cp('ก'), 'default')).toBe('cs');
  });

  it('其余拉丁扩展走 hAnsi', () => {
    // U+0100 Ā 不在 Ambiguous 表里；注意小写 ā（U+0101）在表里，一个码点之差
    expect(bucketOf(cp('Ā'), 'eastAsia')).toBe('hAnsi');
  });

  it('BMP 之外的汉字（扩充 B）也是 eastAsia', () => {
    expect(isEastAsianCodePoint(0x20000)).toBe(true);
    expect(bucketOf(0x20000, 'default')).toBe('eastAsia');
  });

  it('hasEastAsianText 决定整行走不走 1.3 行高系数', () => {
    expect(hasEastAsianText('Report 2024')).toBe(false);
    expect(hasEastAsianText('Report 2024 年')).toBe(true);
    expect(hasEastAsianText('')).toBe(false);
  });
});

describe('桶 → 字体名的回退链', () => {
  it('桶为空表示「文档没指定」，不是错误，要往别的桶找', () => {
    const onlyEA: ScriptFonts = { ascii: '', hAnsi: '', eastAsia: '黑体', cs: '', hint: 'eastAsia' };
    // hint=eastAsia 且没写 w:ascii：Word 用东亚字体画 ASCII，不去找系统默认字体
    expect(bucketFont(onlyEA, 'ascii')).toBe('黑体');
    expect(bucketFont(onlyEA, 'hAnsi')).toBe('黑体');
  });

  it('hint=default 时 ascii 不回退到东亚字体，宁可交还给调用方的默认字体', () => {
    const onlyEA: ScriptFonts = { ascii: '', hAnsi: '', eastAsia: '黑体', cs: '', hint: 'default' };
    expect(bucketFont(onlyEA, 'ascii')).toBe('');
    const withLatin: ScriptFonts = { ...onlyEA, hAnsi: 'Arial' };
    expect(bucketFont(withLatin, 'ascii')).toBe('Arial');
  });

  it('四个桶全空时返回空串，由调用方套默认字体', () => {
    const none: ScriptFonts = { ascii: '', hAnsi: '', eastAsia: '', cs: '', hint: 'default' };
    expect(bucketFont(none, 'eastAsia')).toBe('');
  });
});

describe('切段', () => {
  it('中英混排按字体切开 —— 一个 run 里横跨多款字体', () => {
    const runs = splitFontRuns('2024年1月', gongwen);
    expect(runs).toEqual([
      { start: 0, end: 4, font: 'Times New Roman', script: 'latin' },
      { start: 4, end: 5, font: '仿宋', script: 'eastAsia' },
      { start: 5, end: 6, font: 'Times New Roman', script: 'latin' },
      { start: 6, end: 7, font: '仿宋', script: 'eastAsia' },
    ]);
  });

  it('字体名相同但脚本不同**不合并** —— 中西文 1/8 em 间距要靠这个边界', () => {
    const same: ScriptFonts = {
      ascii: '等线',
      hAnsi: '等线',
      eastAsia: '等线',
      cs: '等线',
      hint: 'eastAsia',
    };
    const runs = splitFontRuns('A中', same);
    expect(runs.map((r) => r.script)).toEqual(['latin', 'eastAsia']);
  });

  it('ascii 与 hAnsi 同字体时合并 —— 它们排版规则相同', () => {
    const runs = splitFontRuns('Café', { ...gongwen, hint: 'default' });
    // é 是 hAnsi 桶，但字体与 ascii 桶相同，且同为 latin
    expect(runs).toEqual([{ start: 0, end: 4, font: 'Times New Roman', script: 'latin' }]);
  });

  it('代理对不会被切开', () => {
    const text = `${String.fromCodePoint(0x20000)}A`;
    const runs = splitFontRuns(text, gongwen);
    expect(runs).toEqual([
      { start: 0, end: 2, font: '仿宋', script: 'eastAsia' },
      { start: 2, end: 3, font: 'Times New Roman', script: 'latin' },
    ]);
    expect(text.slice(0, 2)).toBe(String.fromCodePoint(0x20000));
  });

  it('hint 改变切段结果 —— ① 跟着换字体', () => {
    expect(splitFontRuns('①', gongwen)[0]?.font).toBe('仿宋');
    expect(splitFontRuns('①', { ...gongwen, hint: 'default' })[0]?.font).toBe('Times New Roman');
  });

  it('空文本切出空数组，结果可结构化克隆', () => {
    expect(splitFontRuns('', gongwen)).toEqual([]);
    const runs = splitFontRuns('中文abc', gongwen);
    expect(structuredClone(runs)).toEqual(runs);
  });
});

/**
 * 空格这类中性字符归谁 —— 判据是 `gongwen-01` 的 12 个空格（见 `neutralTakesEastAsia`）。
 * 它不在 `bucketOf` 里，因为 `bucketOf` 只看一个码点，而这一条要看邻居。
 */
describe('中性字符随邻居', () => {
  it('任一侧邻居是东亚字就跟着走东亚桶', () => {
    expect(neutralTakesEastAsia('eastAsia', 'eastAsia', 'latin')).toBe(true);
    expect(neutralTakesEastAsia('eastAsia', 'latin', 'eastAsia')).toBe(true);
    expect(neutralTakesEastAsia('eastAsia', 'eastAsia', 'eastAsia')).toBe(true);
  });

  it('两侧都是拉丁就留在 ascii 桶 —— 「0.5 pt」里那个空格实测是半角', () => {
    expect(neutralTakesEastAsia('eastAsia', 'latin', 'latin')).toBe(false);
  });

  it('一侧没有邻居（行首行末、制表位旁）也按拉丁算', () => {
    expect(neutralTakesEastAsia('eastAsia', undefined, undefined)).toBe(false);
    expect(neutralTakesEastAsia('eastAsia', undefined, 'eastAsia')).toBe(true);
  });

  it('hint 不是 eastAsia 时不改桶 —— 那时没有真值', () => {
    expect(neutralTakesEastAsia('default', 'eastAsia', 'eastAsia')).toBe(false);
    expect(neutralTakesEastAsia('cs', 'eastAsia', 'eastAsia')).toBe(false);
  });
});
