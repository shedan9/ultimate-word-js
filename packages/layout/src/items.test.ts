/**
 * 段落 → item 流：分桶、宽度、中西文自动间距。
 *
 * 期望值全部手算得出（合成字体：东亚 1 em、ASCII 0.5 em，见 test-fixtures.ts），
 * 失败时能直接看出差了几分之几个字。
 */
import { describe, expect, it } from 'vitest';
import { buildItems } from './items.ts';
import { fakeMeasurer, para, run, runOf, SIZE_5 } from './test-fixtures.ts';
import type { CharItem, LayoutItem } from './types.ts';
import { AUTO_SPACE_EM, SMALL_CAPS_SCALE } from './uncalibrated.ts';

const M = { measurer: fakeMeasurer() };
const chars = (items: LayoutItem[]): CharItem[] => items.filter((i): i is CharItem => i.kind === 'char');

describe('分桶与度量', () => {
  it('一个 run 内横跨两款字体 —— 汉字走 eastAsia 桶，英文数字走 ascii 桶', () => {
    const items = chars(buildItems(para([run('中a')]), M));
    expect(items.map((i) => i.font)).toEqual(['仿宋', 'Times New Roman']);
    expect(items.map((i) => i.script)).toEqual(['eastAsia', 'latin']);
    // 汉字全角、ASCII 半角
    expect(items.map((i) => i.width)).toEqual([SIZE_5, SIZE_5 / 2]);
  });

  it('w:w 缩放乘在字形宽度上，w:spacing 字间距是之后再加的常量', () => {
    const items = chars(buildItems(para([run('中', { scale: 50, charSpacing: 20 })]), M));
    expect(items[0]?.width).toBe(SIZE_5 * 0.5 + 20);
  });

  it('w:vanish 的文字不参与排版 —— 不是画成透明', () => {
    expect(buildItems(para([run('中', { hidden: true }), run('文')]), M)).toHaveLength(1);
  });

  it('w:caps 改变的是宽度，不只是外观 —— 大写化在度量之前做', () => {
    const items = chars(buildItems(para([run('ab', { caps: true })]), M));
    expect(items.map((i) => String.fromCodePoint(i.cp)).join('')).toBe('AB');
  });

  it('w:smallCaps 只把原本小写的那些缩小 —— 不是整段缩小', () => {
    const items = chars(buildItems(para([run('Ab', { smallCaps: true })]), M));
    expect(items.map((i) => String.fromCodePoint(i.cp)).join('')).toBe('AB');
    expect(items.map((i) => i.fontSize)).toEqual([SIZE_5, SIZE_5 * SMALL_CAPS_SCALE]);
  });

  it('域代码与域界桩不占宽度', () => {
    const items = buildItems(
      para([
        runOf([
          { kind: 'fieldChar', charType: 'begin' },
          { kind: 'fieldInstruction', text: 'PAGE' },
          { kind: 'fieldChar', charType: 'end' },
        ]),
      ]),
      M,
    );
    expect(items).toEqual([]);
  });

  it('软连字符平时宽度为 0 —— 它参与排版的是「可断」这个性质', () => {
    const items = chars(buildItems(para([runOf([{ kind: 'softHyphen' }])]), M));
    expect(items[0]?.width).toBe(0);
    expect(items[0]?.softHyphen).toBe(true);
  });

  it('内嵌对象按外框尺寸占位', () => {
    const items = buildItems(
      para([runOf([{ kind: 'object', objectKind: 'drawing', width: 1000, height: 800 }])]),
      M,
    );
    expect(items[0]).toMatchObject({ kind: 'object', width: 1000, height: 800 });
  });

  it('offset 指回源文本的 UTF-16 下标，代理对不会被切开', () => {
    // U+20000 是扩充 B 的汉字，UTF-16 占两个码元
    const items = chars(buildItems(para([run('中\u{20000}文')]), M));
    expect(items.map((i) => i.offset)).toEqual([0, 1, 3]);
  });
});

describe('中西文自动间距', () => {
  const gap = SIZE_5 * AUTO_SPACE_EM;

  it('汉字与拉丁字母之间加 1/8 em，记在后一个字符上', () => {
    const items = chars(buildItems(para([run('中a中')]), M));
    expect(items.map((i) => i.gapBefore)).toEqual([0, gap, gap]);
  });

  it('汉字与数字之间由 autoSpaceDN 单独控制', () => {
    expect(chars(buildItems(para([run('中1')], { autoSpaceDN: false }), M))[1]?.gapBefore).toBe(0);
    // DE 关掉不影响数字那一侧
    expect(chars(buildItems(para([run('中1')], { autoSpaceDE: false }), M))[1]?.gapBefore).toBe(gap);
  });

  it('空格两侧不加 —— 已经有空隙了，再加就成了双份', () => {
    expect(chars(buildItems(para([run('中 a')]), M)).map((i) => i.gapBefore)).toEqual([0, 0, 0]);
  });

  it('两侧都是拉丁或都是汉字时不加', () => {
    expect(chars(buildItems(para([run('ab中中')]), M)).map((i) => i.gapBefore)).toEqual([0, 0, gap, 0]);
  });
});
