/**
 * 段落 → item 流：分桶、宽度、中西文自动间距。
 *
 * 期望值全部手算得出（合成字体：东亚 1 em、ASCII 0.5 em，见 test-fixtures.ts），
 * 失败时能直接看出差了几分之几个字。
 */
import { describe, expect, it } from 'vitest';
import { PUNCT_PAIR_COMPRESS_EM } from './break-class.ts';
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

  it('空格挨着东亚字时走 eastAsia 桶 —— 拿拉丁字体的 0.25 em 量会每个空格差 4pt', () => {
    // 实测 gongwen-01：「以 Word」「Word 导出」两侧的空格 Word 都按仿宋的 0.5 em 排
    const items = chars(buildItems(para([run('中 a')]), M));
    expect(items.map((i) => i.font)).toEqual(['仿宋', '仿宋', 'Times New Roman']);
    expect(items.map((i) => i.script)).toEqual(['eastAsia', 'eastAsia', 'latin']);
  });

  it('空格两侧都是拉丁字时照旧走 ascii 桶 —— 「0.5 pt」里那个实测就是半角', () => {
    const items = chars(buildItems(para([run('a b')]), M));
    expect(items.map((i) => i.font)).toEqual(['Times New Roman', 'Times New Roman', 'Times New Roman']);
    expect(items.map((i) => i.script)).toEqual(['latin', 'latin', 'latin']);
  });

  it('空格的邻居**跨 run** —— 这正是它不能在 splitFontRuns 里做的原因', () => {
    const items = chars(buildItems(para([run('a '), run('中')]), M));
    expect(items.map((i) => i.script)).toEqual(['latin', 'eastAsia', 'eastAsia']);
  });

  it('一串空格整体随邻居，不会因为「邻居也是空格」就判不出来', () => {
    const items = chars(buildItems(para([run('a  中')]), M));
    expect(items.map((i) => i.script)).toEqual(['latin', 'eastAsia', 'eastAsia', 'eastAsia']);
  });

  it('hint 不是 eastAsia 时空格不改桶 —— 那时没有真值，保持规范默认', () => {
    const fonts = {
      ascii: 'Times New Roman',
      hAnsi: 'Times New Roman',
      eastAsia: '仿宋',
      cs: '',
      hint: 'default' as const,
    };
    const items = chars(buildItems(para([run('中 a', { fonts })]), M));
    expect(items.map((i) => i.script)).toEqual(['eastAsia', 'latin', 'latin']);
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

describe('相邻标点挤压', () => {
  // 实测（spike-punct-01，26 段短句）：孤立的标点一点都不压，两个标点相邻固定压掉半个字。
  // 这是**常态排版**，与「行尾塞不下」无关 —— 所以做在 item 流这一步，见 applyPunctPairs
  const gaps = (text: string): number[] => chars(buildItems(para([run(text)]), M)).map((i) => i.gapBefore);
  const HALF = -SIZE_5 * PUNCT_PAIR_COMPRESS_EM;

  it('孤立的标点不压 —— 「甲，乙」的三个字宽精确等于三个字号', () => {
    expect(gaps('一，二')).toEqual([0, 0, 0]);
  });

  it('两个标点相邻，后一个往左挪半个字', () => {
    expect(gaps('一，，二')).toEqual([0, 0, HALF, 0]);
  });

  it('三连标点 = 两对，各压半个字', () => {
    expect(gaps('一，，，二')).toEqual([0, 0, HALF, HALF, 0]);
  });

  it('收口 + 开口中间空着整整一个字，但也只压半个 —— 实测如此', () => {
    expect(gaps('一，（二）三')).toEqual([0, 0, HALF, 0, 0, 0]);
  });

  it('行首、行末、紧邻汉字的标点都不压', () => {
    expect(gaps('（一二）三')).toEqual([0, 0, 0, 0, 0]);
  });

  it('省略号与破折号不在可挤压表里 —— 它们的墨横贯整个字宽，没有空半边', () => {
    expect(gaps('一…—二')).toEqual([0, 0, 0, 0]);
  });

  it('开口紧跟收口（「，）不压 —— 接缝两侧都是墨', () => {
    // gongwen-01 真值第 10 行（0 起）实测：Word 在那一串标点的 12 个接缝上各挤半个字，唯独「，那个没挤
    expect(gaps('一「，二')).toEqual([0, 0, 0, 0]);
  });

  it('收口紧跟省略号（】…）要压 —— 挤掉的是「】」的右半边，与省略号自己压不压无关', () => {
    expect(gaps('一】…二')).toEqual([0, 0, HALF, 0]);
  });

  it('文档关掉 w:characterSpacingControl 时一个都不压', () => {
    const off = { measurer: fakeMeasurer(), compressPunctuation: false };
    expect(chars(buildItems(para([run('一，，二')]), off)).map((i) => i.gapBefore)).toEqual([0, 0, 0, 0]);
  });
});
