/**
 * 编号文字的格式化。
 *
 * 这一层是编号里**唯一能被字符级钉死**的部分，所以测得细一点：格式化错一位，
 * 表现是「编号看起来怪」而不是报错，靠肉眼在几十页文档里发现不了。
 *
 * 中文读法与 `chineseCounting` / `chineseCountingThousand` 的分岔尚无 Word 真值
 * （见 number-format.ts 文件头），下面的期望值是按中文习惯写的，**不是实测结论**。
 */
import { describe, expect, it } from 'vitest';
import { formatLevelText, formatNumber } from './number-format.ts';

describe('formatNumber —— 拉丁系', () => {
  it('decimal / decimalZero：补零只补到两位', () => {
    expect(formatNumber(7, 'decimal')).toBe('7');
    expect(formatNumber(7, 'decimalZero')).toBe('07');
    expect(formatNumber(70, 'decimalZero')).toBe('70');
    expect(formatNumber(700, 'decimalZero')).toBe('700');
  });

  it('字母超过 Z 是**重复**不是进位 —— 27 是 AA 不是 AB', () => {
    expect(formatNumber(1, 'upperLetter')).toBe('A');
    expect(formatNumber(26, 'upperLetter')).toBe('Z');
    expect(formatNumber(27, 'upperLetter')).toBe('AA');
    expect(formatNumber(28, 'lowerLetter')).toBe('bb');
    expect(formatNumber(53, 'upperLetter')).toBe('AAA');
  });

  it('罗马数字，超出 3999 退回十进制', () => {
    expect(formatNumber(4, 'upperRoman')).toBe('IV');
    expect(formatNumber(1994, 'upperRoman')).toBe('MCMXCIV');
    expect(formatNumber(9, 'lowerRoman')).toBe('ix');
    expect(formatNumber(4000, 'upperRoman')).toBe('4000');
  });

  it('ordinal 的 11/12/13 是 th，不是 st/nd/rd', () => {
    const ord = (v: number) => formatNumber(v, 'ordinal');
    expect([1, 2, 3, 4, 11, 12, 13, 21, 22].map(ord)).toEqual([
      '1st',
      '2nd',
      '3rd',
      '4th',
      '11th',
      '12th',
      '13th',
      '21st',
      '22nd',
    ]);
  });

  it('全角数字与带圈数字，带圈超过 20 没有字形，退回十进制', () => {
    expect(formatNumber(12, 'decimalFullWidth')).toBe('１２');
    expect(formatNumber(1, 'decimalEnclosedCircle')).toBe('①');
    expect(formatNumber(20, 'decimalEnclosedCircle')).toBe('⑳');
    expect(formatNumber(21, 'decimalEnclosedCircle')).toBe('21');
    expect(formatNumber(1, 'decimalEnclosedParen')).toBe('⑴');
    expect(formatNumber(1, 'decimalEnclosedFullstop')).toBe('⒈');
  });
});

describe('formatNumber —— 中文系', () => {
  it('口语读法省掉「一十」的一，但只在小于 20 时', () => {
    const cn = (v: number) => formatNumber(v, 'chineseCounting');
    expect([1, 9, 10, 11, 19, 20, 21, 99].map(cn)).toEqual([
      '一',
      '九',
      '十',
      '十一',
      '十九',
      '二十',
      '二十一',
      '九十九',
    ]);
    // 110 读「一百一十」—— 省略只发生在最高位就是十位的时候
    expect([100, 105, 110, 1234].map(cn)).toEqual(['一百', '一百零五', '一百一十', '一千二百三十四']);
  });

  it('万 / 亿分组，低位不足四位时补「零」', () => {
    const cn = (v: number) => formatNumber(v, 'chineseCountingThousand');
    expect([10000, 10001, 100000000].map(cn)).toEqual(['一万', '一万零一', '一亿']);
  });

  it('法定大写不省略「壹拾」的壹 —— 省了就是另一个数', () => {
    expect(formatNumber(10, 'chineseLegalSimplified')).toBe('壹拾');
    expect(formatNumber(15, 'chineseLegalSimplified')).toBe('壹拾伍');
    expect(formatNumber(2026, 'chineseLegalSimplified')).toBe('贰仟零贰拾陆');
  });

  it('天干地支超出表长就退回十进制，不循环 —— 循环会出现两个「甲」', () => {
    expect(formatNumber(1, 'ideographTraditional')).toBe('甲');
    expect(formatNumber(10, 'ideographTraditional')).toBe('癸');
    expect(formatNumber(11, 'ideographTraditional')).toBe('11');
    expect(formatNumber(12, 'ideographZodiac')).toBe('亥');
  });
});

describe('formatNumber —— 降级', () => {
  it('认不出的 numFmt 出数字，不出空 —— 编号消失比编号不好看严重得多', () => {
    expect(formatNumber(3, 'ordinalText')).toBe('3');
    expect(formatNumber(3, '未来某个新格式')).toBe('3');
  });

  it('none 是空串；0 与负数只有十进制说得清', () => {
    expect(formatNumber(3, 'none')).toBe('');
    expect(formatNumber(0, 'chineseCounting')).toBe('0');
    expect(formatNumber(-1, 'upperRoman')).toBe('-1');
  });
});

describe('formatLevelText', () => {
  const decimals = ['decimal', 'decimal', 'decimal'];

  it('%n 里的 n 是 1 起的级号：%1 取 ilvl 0', () => {
    expect(formatLevelText('%1.%2.%3', [2, 3, 4], decimals)).toBe('2.3.4');
    expect(formatLevelText('第%1章', [7], decimals)).toBe('第7章');
  });

  it('每级用自己的 numFmt', () => {
    expect(formatLevelText('%1、%2)', [3, 2], ['chineseCounting', 'lowerLetter'])).toBe('三、b)');
  });

  it('% 后面不是 1–9 就是字面百分号', () => {
    expect(formatLevelText('100%', [1], decimals)).toBe('100%');
    expect(formatLevelText('%0%1', [5], decimals)).toBe('%05');
  });

  it('引用了一个连 start 都补不出的级：占位符整个吞掉，不把「%3」印出来', () => {
    expect(formatLevelText('%1.%3', [1], decimals)).toBe('1.');
  });
});
