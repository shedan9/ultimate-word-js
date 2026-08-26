/**
 * 内嵌对象与浮动对象。
 *
 * 版心与 page.test.ts 同一套（10 字宽 × 3 行高），期望值全部能心算。
 * 三件事分开测：
 * ① **内嵌**图占文字流的宽与高（断行、行高都要跟着变）；
 * ② **浮动**（`wrap="none"`）不占文字流 —— 一个字都不许被它挤走；
 * ③ 浮动对象的**纸坐标**：六种参照物 × 偏移 / 对齐两种写法。
 */
import type { DrawingAnchor, ResolvedBlock, ResolvedBody, SectionProps } from '@uw/model';
import { DEFAULT_SECTION_PROPS, DEFAULT_SETTINGS } from '@uw/model';
import { describe, expect, it } from 'vitest';
import type { LayoutDocumentOptions } from './page.ts';
import { layoutDocument } from './page.ts';
import { layoutParagraph } from './paragraph.ts';
import { fakeMeasurer, NO_GRID, para, paraProps, run, runOf, SIZE_5 } from './test-fixtures.ts';

const EA_LINE = SIZE_5 * 1.3;
const TEN = '一二三四五六七八九十';

/** 版心 10 字宽（2100）× 3 行高（819），左上角在 (600, 600) */
function sect(over: Partial<SectionProps> = {}): SectionProps {
  return {
    ...structuredClone(DEFAULT_SECTION_PROPS),
    page: { width: 3300, height: 2019, orientation: 'portrait' },
    margin: { top: 600, right: 600, bottom: 600, left: 600, header: 0, footer: 0, gutter: 0 },
    docGrid: NO_GRID,
    ...over,
  };
}

function body(blocks: ResolvedBlock[], props: SectionProps = sect()): ResolvedBody {
  return { sections: [{ id: 's0', props, blocks }] };
}

function opts(over: Partial<LayoutDocumentOptions> = {}): LayoutDocumentOptions {
  return { measurer: fakeMeasurer(), settings: DEFAULT_SETTINGS, ...over };
}

/** 一张 w × h 的图 */
function pic(width: number, height: number, anchor?: DrawingAnchor) {
  return runOf([
    {
      kind: 'object' as const,
      objectKind: 'drawing' as const,
      width,
      height,
      image: { id: 'rId5', relId: 'rId5' },
      alt: '图',
      ...(anchor === undefined ? {} : { anchor }),
    },
  ]);
}

function anchorOf(over: Partial<DrawingAnchor> = {}): DrawingAnchor {
  return {
    wrap: 'none',
    behindDoc: false,
    z: 0,
    h: { relativeFrom: 'page', offset: 0 },
    v: { relativeFrom: 'page', offset: 0 },
    dist: { top: 0, bottom: 0, left: 0, right: 0 },
    ...over,
  };
}

const paraOpts = {
  measurer: fakeMeasurer(),
  contentWidth: 2100,
  settings: DEFAULT_SETTINGS,
  docGrid: NO_GRID,
};

describe('内嵌对象', () => {
  it('占宽度：一张 3 个字宽的图 + 8 个字，一行放不下', () => {
    const p = layoutParagraph(para([pic(SIZE_5 * 3, SIZE_5), run(TEN.slice(0, 8))]), paraOpts);
    expect(p.lines).toHaveLength(2);
    // 图 3 字 + 文字 7 字 = 10 字，第 8 个字换行
    expect(p.lines[0]?.fragments[0]?.text).toBe('一二三四五六七');
  });

  it('占高度：图比行高时行跟着变高，图的底边坐在基线上，文字的下伸还留着', () => {
    const tall = SIZE_5 * 4;
    const plain = layoutParagraph(para([run('一')]), paraOpts);
    const textBelow = (plain.lines[0]?.height ?? 0) - (plain.lines[0]?.baseline ?? 0);
    const p = layoutParagraph(para([pic(SIZE_5, tall), run('一')]), paraOpts);
    const line = p.lines[0];
    // 图撑的是基线**以上**那一截，文字自己的下伸照旧留在基线以下（实测，见 OBJECT_RULES ②）
    expect(line?.baseline).toBe(tall);
    expect(line?.height).toBe(tall + textBelow);
    expect(line?.objects).toEqual([
      {
        runId: expect.any(String),
        contentIndex: 0,
        x: 0,
        width: SIZE_5,
        height: tall,
        objectKind: 'drawing',
        image: { id: 'rId5', relId: 'rId5' },
        alt: '图',
      },
    ]);
  });

  it('对象排在文字后面时 x 是累加出来的，不是行首', () => {
    const p = layoutParagraph(para([run('一二'), pic(SIZE_5, SIZE_5)]), paraOpts);
    expect(p.lines[0]?.objects?.[0]?.x).toBe(SIZE_5 * 2);
  });
});

/**
 * 内嵌对象 × 行网格 / 倍数行距（`spike-image-03` 标定，见 `OBJECT_RULES` 第 ⑤ 条）。
 * 期望值全部手算：网格 400、东亚行自然行高 273（= 210 × 1.3）、基线在自然行高里 199.5、
 * 下伸 73.5；图高 480 是 30 的整数倍，绕开盒高量化那一条。
 */
describe('内嵌对象 × 行网格与倍数行距', () => {
  const PITCH = 400;
  const IMG = 480;
  /** 图撑起来的那一截：盒高 + 文字自己的下伸 = 480 + 73.5 */
  const OBJ_NATURAL = IMG + EA_LINE - 199.5;
  const gridOpts = { ...paraOpts, docGrid: { type: 'lines' as const, linePitch: PITCH, charSpace: 0 } };
  const multiple = { spacing: { ...paraProps().spacing, line: 360 } };

  it('含图的行照样吸网格：吸的是「盒高 + 文字下伸」，富余上下均分', () => {
    const line = layoutParagraph(para([pic(SIZE_5, IMG), run('一')]), gridOpts).lines[0];
    // 553.5 要两个网格行
    expect(line?.height).toBe(PITCH * 2);
    expect(line?.baseline).toBeCloseTo(IMG + (PITCH * 2 - OBJ_NATURAL) / 2, 6);
  });

  it('倍数行距**不乘在图撑起来的那一截上**：两侧各算各的再取大', () => {
    const line = layoutParagraph(para([pic(SIZE_5, IMG), run('一')], multiple), gridOpts).lines[0];
    // 文字侧 400 × 1.5 = 600，对象侧 (553.5 + 倍数多留的 136.5) 吸成 800 —— 取 800。
    // 旧实现（两者合成一个自然行高再乘）给的是 800 × 1.5 = 1200，实测 Word 不是那样
    expect(line?.height).toBe(PITCH * 2);
  });

  it('关网格 + 倍数行距：多留出来的空白整个落在基线以下，图底仍坐在基线上', () => {
    const line = layoutParagraph(para([pic(SIZE_5, IMG), run('一')], multiple), paraOpts).lines[0];
    // 对象侧 = 553.5 + 136.5 = 690，文字侧只有 409.5
    expect(line?.height).toBeCloseTo(OBJ_NATURAL + (EA_LINE * 1.5 - EA_LINE), 6);
    // 居中用的是**对象侧的行盒**（553.5），不是推进量 690 —— 否则基线还要往下沉 68.25
    expect(line?.baseline).toBe(IMG);
  });
});

describe('浮动对象（wrap="none"）', () => {
  it('一个字都不挤走：整行 10 个字照排，图不进 objects 而进 floats', () => {
    const anchor = anchorOf();
    const p = layoutParagraph(para([pic(SIZE_5 * 5, SIZE_5 * 5, anchor), run(TEN)]), paraOpts);
    expect(p.lines).toHaveLength(1);
    expect(p.lines[0]?.fragments[0]?.text).toBe(TEN);
    // 行高也不受它影响 —— 5 倍字号的图放进去，行还是一行的高
    expect(p.lines[0]?.height).toBe(EA_LINE);
    expect(p.lines[0]?.objects).toBeUndefined();
    expect(p.lines[0]?.floats?.[0]?.width).toBe(SIZE_5 * 5);
  });

  it('其余环绕方式照样浮动 —— 没做的是「文字让开」，不是「它在行里」', () => {
    // 3 个字宽、比正文高的一张 square 环绕图：当成内嵌就会挤掉 3 个字并把行撑高，
    // 而 Word 里它根本不在这一行上。真实语料里页脚的 topAndBottom 文本框就是这么错的
    const p = layoutParagraph(
      para([pic(SIZE_5 * 3, SIZE_5 * 3, anchorOf({ wrap: 'square' })), run(TEN)]),
      paraOpts,
    );
    expect(p.lines).toHaveLength(1);
    expect(p.lines[0]?.height).toBe(EA_LINE);
    expect(p.lines[0]?.objects).toBeUndefined();
    expect(p.lines[0]?.floats?.[0]?.width).toBe(SIZE_5 * 3);
  });
});

describe('浮动对象的纸坐标', () => {
  /** 一份只有一段的文档，段落里挂一张浮动图 */
  function floatsOf(anchor: DrawingAnchor, blocks?: ResolvedBlock[]) {
    const doc = layoutDocument(body(blocks ?? [para([pic(200, 100, anchor), run(TEN)])]), opts());
    return doc.pages[0]?.floats ?? [];
  }

  it('relativeFrom="page" + 偏移：直接就是纸坐标', () => {
    const [f] = floatsOf(
      anchorOf({ h: { relativeFrom: 'page', offset: 300 }, v: { relativeFrom: 'page', offset: 400 } }),
    );
    expect(f).toMatchObject({ x: 300, y: 400, width: 200, height: 100, behindDoc: false, z: 0 });
  });

  it('relativeFrom="margin" 从版心左上角起算（版心在 600,600）', () => {
    const [f] = floatsOf(
      anchorOf({
        h: { relativeFrom: 'margin', offset: 100 },
        v: { relativeFrom: 'margin', offset: 50 },
      }),
    );
    expect(f?.x).toBe(700);
    expect(f?.y).toBe(650);
  });

  it('align 居中 / 右对齐按参照框算，图自己的宽要减掉', () => {
    const [center] = floatsOf(
      anchorOf({
        h: { relativeFrom: 'page', align: 'center' },
        v: { relativeFrom: 'page', align: 'bottom' },
      }),
    );
    // 纸宽 3300，图宽 200 → (3300 − 200) / 2
    expect(center?.x).toBe(1550);
    expect(center?.y).toBe(2019 - 100);

    const [right] = floatsOf(anchorOf({ h: { relativeFrom: 'margin', align: 'right' } }));
    // 版心右边 600 + 2100，减图宽
    expect(right?.x).toBe(2700 - 200);
  });

  it('inside / outside 看**显示页码**的奇偶：第 1 页 inside 是左边', () => {
    const [f] = floatsOf(anchorOf({ h: { relativeFrom: 'page', align: 'inside' } }));
    expect(f?.x).toBe(0);
    const [g] = floatsOf(anchorOf({ h: { relativeFrom: 'page', align: 'outside' } }));
    expect(g?.x).toBe(3300 - 200);
  });

  it('character / paragraph 跟着锚点走：图排在两个字之后就从那儿起算', () => {
    const doc = layoutDocument(
      body([
        para([run(TEN)]),
        para([
          run('一二'),
          pic(
            200,
            100,
            anchorOf({
              h: { relativeFrom: 'character', offset: 0 },
              v: { relativeFrom: 'paragraph', offset: 0 },
            }),
          ),
        ]),
      ]),
      opts(),
    );
    const [f] = doc.pages[0]?.floats ?? [];
    // 版心左 600 + **一个**五号字：character 参照的是锚点前一个字的左边缘（实测，
    // 见 FLOAT_ORIGIN_RULES ⑥）—— 图排在「一二」之后，参照的是「二」
    expect(f?.x).toBe(600 + SIZE_5);
    // 版心顶 600 + 上面那一段的一行
    expect(f?.y).toBe(600 + EA_LINE);
  });

  it('浮动对象跟着锚点所在的**页**走，不是永远在第一页', () => {
    const blocks = [
      para([run(TEN)]),
      para([run(TEN)]),
      para([run(TEN)]),
      para([run('一'), pic(200, 100, anchorOf({ v: { relativeFrom: 'paragraph', offset: 0 } }))]),
    ];
    const doc = layoutDocument(body(blocks), opts());
    expect(doc.pages).toHaveLength(2);
    expect(doc.pages[0]?.floats).toBeUndefined();
    expect(doc.pages[1]?.floats).toHaveLength(1);
    // 第 2 页的第一行 → 版心顶
    expect(doc.pages[1]?.floats?.[0]?.y).toBe(600);
  });

  it('同一页上按 z 序排好，衬于文字下方的标出来给渲染层分层用', () => {
    const doc = layoutDocument(
      body([
        para([pic(10, 10, anchorOf({ z: 5 })), pic(20, 20, anchorOf({ z: 1, behindDoc: true })), run(TEN)]),
      ]),
      opts(),
    );
    expect(doc.pages[0]?.floats?.map((f) => [f.z, f.behindDoc])).toEqual([
      [1, true],
      [5, false],
    ]);
  });
});
