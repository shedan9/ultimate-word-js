/**
 * 段落边框与底纹的几何：让出多少高、怎么成组、分页时跟着谁走。
 *
 * 这一层没有真值（见 `uncalibrated.ts`「段落边框」一节），这里测的是**自洽**：
 * 让出的高度进了行的 y、成组规则按键表判、页首只丢段前间距不丢边框。
 * 版心与 `page.test.ts` 同一套：一行 10 个字、行高 273、一页 3 行。
 */
import type { Border, ResolvedBlock, ResolvedBody, SectionProps } from '@uw/model';
import { DEFAULT_SECTION_PROPS, DEFAULT_SETTINGS } from '@uw/model';
import { describe, expect, it } from 'vitest';
import type { PlacedParagraph, PlacedTable } from './page.ts';
import { layoutDocument } from './page.ts';
import { layoutParagraph } from './paragraph.ts';
import { cell, fakeMeasurer, NO_GRID, para, paraProps, row, run, SIZE_5, table } from './test-fixtures.ts';

const EA_LINE = SIZE_5 * 1.3;
const measurer = fakeMeasurer();

/** 线宽 1pt（20 twips）、间距 1pt —— 一条线让出 40 */
const line = (over: Partial<Border> = {}): Border => ({
  style: 'single',
  size: 20,
  space: 20,
  color: 'auto',
  ...over,
});

function sect(height = 2019): SectionProps {
  return {
    ...structuredClone(DEFAULT_SECTION_PROPS),
    page: { width: 3300, height, orientation: 'portrait' },
    margin: { top: 600, right: 600, bottom: 600, left: 600, header: 0, footer: 0, gutter: 0 },
    docGrid: NO_GRID,
  };
}

function lay(blocks: ResolvedBlock[], height?: number) {
  const body: ResolvedBody = { sections: [{ id: 's0', props: sect(height), blocks }] };
  return layoutDocument(body, { measurer, settings: DEFAULT_SETTINGS });
}

const paras = (doc: ReturnType<typeof lay>, page: number): PlacedParagraph[] =>
  (doc.pages[page]?.blocks ?? []).filter((b): b is PlacedParagraph => b.kind === 'paragraph');

function boxed(id: string, text: string, over: Parameters<typeof paraProps>[0] = {}) {
  return {
    ...para([run(text)], {
      borders: {
        top: line(),
        bottom: line(),
        left: line(),
        right: line(),
        between: line({ color: 'FF0000' }),
      },
      ...over,
    }),
    id,
  };
}

describe('单独一段', () => {
  it('上下边框让出「线宽 + 间距」，折进段前段后；框是文字区，悬挂的编号在框里', () => {
    const p = para([run('一二')], {
      borders: { top: line(), bottom: line({ size: 10, style: 'double' }) },
      spacing: { ...paraProps().spacing, before: 100, after: 50 },
      indent: { ...paraProps().indent, left: 400, hanging: 200, right: 300 },
    });
    const l = layoutParagraph(p, {
      measurer,
      settings: DEFAULT_SETTINGS,
      contentWidth: 2100,
      docGrid: NO_GRID,
    });
    expect(l.frame?.insetTop).toBe(40);
    // 双线画出来 3 倍厚（与表格格线同一张表）：30 + 20
    expect(l.frame?.insetBottom).toBe(50);
    expect(l.spaceBefore).toBe(140);
    expect(l.spaceAfter).toBe(100);
    expect(l.frame?.x).toBe(200);
    expect(l.frame?.width).toBe(2100 - 300 - 200);
  });

  it('没有边框也没有底纹时不出 frame；nil 的边等于没有', () => {
    const l = layoutParagraph(para([run('一')], { borders: { bottom: line({ style: 'nil' }) } }), {
      measurer,
      settings: DEFAULT_SETTINGS,
      contentWidth: 2100,
      docGrid: NO_GRID,
    });
    expect(l.frame).toBeUndefined();
    expect(l.spaceAfter).toBe(0);
  });
});

describe('成组', () => {
  it('边框相同的相邻段框在一起：组内画 between、不画各自的上下边，缩进不同就断组', () => {
    const doc = lay(
      [
        boxed('a', '一'),
        boxed('b', '二'),
        boxed('c', '三', { indent: { ...paraProps().indent, left: 200 } }),
      ],
      4000,
    );
    const [a, b, c] = paras(doc, 0) as [PlacedParagraph, PlacedParagraph, PlacedParagraph];
    expect([a.frame?.top?.color, a.frame?.bottom, a.frame?.joinNext]).toEqual(['auto', undefined, true]);
    // b 顶上是红色的 between，底下是组末的下边
    expect([b.frame?.top?.color, b.frame?.bottom?.color, b.frame?.joinPrev]).toEqual([
      'FF0000',
      'auto',
      true,
    ]);
    expect(c.frame?.joinPrev).toBe(false);
    // y：a 首行让出上边 40；b 首行让出 between 40；c 前面是 b 的下边 40 + c 自己的上边 40
    expect(a.y).toBe(40);
    expect(b.y).toBe(40 + EA_LINE + 40);
    expect(c.y).toBe(b.y + EA_LINE + 40 + 40);
  });

  it('组内没有 between 时段与段之间什么都不让', () => {
    const plain = { top: line(), bottom: line() };
    const doc = lay([
      { ...para([run('一')], { borders: plain }), id: 'a' },
      { ...para([run('二')], { borders: plain }), id: 'b' },
    ]);
    const [a, b] = paras(doc, 0) as [PlacedParagraph, PlacedParagraph];
    expect(b.frame?.top).toBeUndefined();
    expect(b.y).toBe(a.y + EA_LINE);
  });
});

describe('分页', () => {
  it('页首只丢段前间距，上边框照样让出', () => {
    const doc = lay([
      {
        ...para([run('一')], { borders: { top: line() }, spacing: { ...paraProps().spacing, before: 300 } }),
        id: 'a',
      },
    ]);
    expect(paras(doc, 0)[0]?.y).toBe(40);
  });

  it('整段挪到下一页时上边框跟着首行走；下边框要与末行一起放得下', () => {
    const doc = lay([
      { ...para([run('一')]), id: 'x' },
      { ...para([run('二')]), id: 'y' },
      // 本页剩一行：放得下字，放不下「字 + 下边框」
      { ...para([run('三')], { borders: { top: line(), bottom: line() } }), id: 'z' },
    ]);
    expect(paras(doc, 0).map((p) => p.id)).toEqual(['x', 'y']);
    expect(paras(doc, 1)[0]?.y).toBe(40);
  });
});

describe('表格单元格', () => {
  it('格内相邻段落同样成组，内容高度含让出的高度', () => {
    const doc = lay([table([2100], [row([cell([boxed('a', '一'), boxed('b', '二')])])])]);
    const t = doc.pages[0]?.blocks[0] as PlacedTable;
    const blocks = t.rows[0]?.row.cells[0]?.blocks ?? [];
    const [a, b] = blocks.map((x) => (x.kind === 'paragraph' ? x.layout : undefined));
    expect([a?.frame?.joinNext, b?.frame?.joinPrev]).toEqual([true, true]);
    expect((a?.spaceBefore ?? 0) + (a?.spaceAfter ?? 0) + (b?.spaceBefore ?? 0) + (b?.spaceAfter ?? 0)).toBe(
      40 + 40 + 40,
    );
  });
});
