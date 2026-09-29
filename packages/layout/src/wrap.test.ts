/**
 * 环绕：方形 / 上下型环绕的对象让开文字。
 *
 * 版心 10 字宽（2100）× 10 行高（2730），左上角在 (600, 600)；五号字 210 宽、东亚行高 273，
 * 期望值全部能心算。几何规则本身**没有真值**（见 wrap.ts 文件头），这里钉的是结构：
 * 让开哪一段、推到哪儿、跨页后断点接不接得上、页眉里的对象算不算。
 */
import { createDiagnosticSink } from '@uw/core';
import type { DrawingAnchor, ResolvedBlock, ResolvedBody, SectionProps } from '@uw/model';
import { DEFAULT_SECTION_PROPS, DEFAULT_SETTINGS } from '@uw/model';
import { describe, expect, it } from 'vitest';
import type { HeaderFooterSource } from './header-footer.ts';
import type { DocumentLayout, LayoutDocumentOptions } from './page.ts';
import { layoutDocument } from './page.ts';
import { fakeMeasurer, NO_GRID, para, run, runOf, SIZE_5 } from './test-fixtures.ts';
import type { WrapExclusion } from './wrap.ts';
import { slotAround } from './wrap.ts';

const LINE = SIZE_5 * 1.3;
const TEN = '一二三四五六七八九十';

function sect(over: Partial<SectionProps> = {}): SectionProps {
  return {
    ...structuredClone(DEFAULT_SECTION_PROPS),
    page: { width: 3300, height: 600 + LINE * 10 + 600, orientation: 'portrait' },
    margin: { top: 600, right: 600, bottom: 600, left: 600, header: 0, footer: 0, gutter: 0 },
    docGrid: NO_GRID,
    ...over,
  };
}

function lay(blocks: ResolvedBlock[], over: Partial<LayoutDocumentOptions> = {}, props = sect()) {
  const body: ResolvedBody = { sections: [{ id: 's0', props, blocks }] };
  return layoutDocument(body, { measurer: fakeMeasurer(), settings: DEFAULT_SETTINGS, ...over });
}

function pic(width: number, height: number, anchor: DrawingAnchor) {
  return runOf([
    {
      kind: 'object' as const,
      objectKind: 'drawing' as const,
      width,
      height,
      image: { id: 'rId5', relId: 'rId5' },
      anchor,
    },
  ]);
}

/** 缺省：方形环绕、贴着版心右边、段顶对齐 */
function anchorOf(over: Partial<DrawingAnchor> = {}): DrawingAnchor {
  return {
    wrap: 'square',
    behindDoc: false,
    z: 0,
    h: { relativeFrom: 'margin', align: 'right' },
    v: { relativeFrom: 'paragraph', offset: 0 },
    dist: { top: 0, bottom: 0, left: 0, right: 0 },
    ...over,
  };
}

/** 每页每行的文字 */
function textsOf(doc: DocumentLayout): string[][] {
  return doc.pages.map((p) =>
    p.blocks.flatMap((b) =>
      b.kind === 'paragraph' ? b.lines.map((l) => l.line.fragments.map((f) => f.text).join('')) : [],
    ),
  );
}

function linesOf(doc: DocumentLayout, page = 0) {
  return (doc.pages[page]?.blocks ?? []).flatMap((b) => (b.kind === 'paragraph' ? b.lines : []));
}

describe('slotAround（纯几何）', () => {
  const base = { left: 0, avail: 2100, skip: 0 };
  const ex = (over: Partial<WrapExclusion>): WrapExclusion => ({
    left: 1260,
    right: 2100,
    top: 0,
    bottom: 546,
    mode: 'square',
    side: 'bothSides',
    ...over,
  });

  it('不碰禁区就是原样；碰边不算碰', () => {
    expect(slotAround(base, [ex({})], 546, LINE, SIZE_5)).toEqual(base);
    expect(slotAround(base, [], 0, LINE, SIZE_5)).toEqual(base);
  });

  it('方形：禁区在右边就用左边那一段，在左边就用右边那一段', () => {
    expect(slotAround(base, [ex({})], 0, LINE, SIZE_5)).toEqual({ left: 0, avail: 1260, skip: 0 });
    expect(slotAround(base, [ex({ left: 0, right: 840 })], 0, LINE, SIZE_5)).toEqual({
      left: 840,
      avail: 1260,
      skip: 0,
    });
  });

  it('上下型：整行推到禁区底下，推下去的那一截记在 skip 上', () => {
    expect(slotAround(base, [ex({ mode: 'topAndBottom', left: 1800 })], 100, LINE, SIZE_5)).toEqual({
      left: 0,
      avail: 2100,
      skip: 446,
    });
  });

  it('两侧都窄过一个字就推下去；推下去之后再碰到下一个禁区接着推', () => {
    const narrow = ex({ left: 100, right: 2000, bottom: 300 });
    const next = ex({ mode: 'topAndBottom', top: 400, bottom: 800 });
    // 推到 300，行盒 [300, 573) 又碰上 [400, 800)，再推到 800
    expect(slotAround(base, [narrow, next], 0, LINE, SIZE_5)).toEqual({ left: 0, avail: 2100, skip: 800 });
  });

  it('wrapText="left" 只走对象左边，哪怕右边更宽', () => {
    const got = slotAround(base, [ex({ left: 420, right: 840, side: 'left' })], 0, LINE, SIZE_5);
    expect(got).toEqual({ left: 0, avail: 420, skip: 0 });
  });

  it('bothSides 两侧都放得下时退化成宽的那一侧，并标出来', () => {
    const got = slotAround(base, [ex({ left: 630, right: 1050 })], 0, LINE, SIZE_5);
    expect(got).toEqual({ left: 1050, avail: 1050, skip: 0, bothSidesApproximated: true });
    // largest 本来就是这个意思，不标
    expect(slotAround(base, [ex({ left: 630, right: 1050, side: 'largest' })], 0, LINE, SIZE_5)).toEqual({
      left: 1050,
      avail: 1050,
      skip: 0,
    });
  });

  it('段落自己的缩进与禁区取交集', () => {
    const indented = { left: 420, avail: 1680, skip: 0 };
    expect(slotAround(indented, [ex({ left: 0, right: 840 })], 0, LINE, SIZE_5)).toEqual({
      left: 840,
      avail: 1260,
      skip: 0,
    });
  });
});

describe('分页里的环绕', () => {
  const THIRTY = TEN.repeat(3);

  it('四周型贴右边：锚点段落的头两行让出 4 个字，其余行照常', () => {
    const doc = lay([para([pic(SIZE_5 * 4, LINE * 2, anchorOf()), run(THIRTY)])]);
    expect(textsOf(doc)[0]?.map((t) => t.length)).toEqual([6, 6, 10, 8]);
    expect(textsOf(doc)[0]?.join('')).toBe(THIRTY);
    // 对象自己照旧画在版心右上角
    expect(doc.pages[0]?.floats?.[0]).toMatchObject({ x: 600 + SIZE_5 * 6, y: 600 });
  });

  it('贴左边时文字从对象右边起排', () => {
    const doc = lay([
      para([pic(SIZE_5 * 4, LINE * 2, anchorOf({ h: { relativeFrom: 'margin', offset: 0 } })), run(THIRTY)]),
    ]);
    const lines = linesOf(doc);
    expect(lines.map((l) => l.line.x)).toEqual([SIZE_5 * 4, SIZE_5 * 4, 0, 0]);
    expect(lines[0]?.line.fragments[0]?.x).toBe(SIZE_5 * 4);
  });

  it('dist 把禁区往外扩：左边再留一个字的距离，每行只剩 5 个字', () => {
    const anchor = anchorOf({ dist: { top: 0, bottom: 0, left: SIZE_5, right: 0 } });
    const doc = lay([para([pic(SIZE_5 * 4, LINE * 2, anchor), run(THIRTY)])]);
    expect(textsOf(doc)[0]?.map((t) => t.length)).toEqual([5, 5, 10, 10]);
  });

  it('上下型：锚点段落自己被推到对象底下', () => {
    const anchor = anchorOf({ wrap: 'topAndBottom', h: { relativeFrom: 'margin', offset: 0 } });
    const doc = lay([para([pic(SIZE_5 * 2, 300, anchor), run(TEN)])]);
    const [first] = linesOf(doc);
    expect(first?.y).toBe(300);
    expect(first?.line.skip).toBe(300);
    expect(textsOf(doc)[0]).toEqual([TEN]);
  });

  it('后面的段落也绕着它走，绕过去之后恢复整行', () => {
    const doc = lay([para([pic(SIZE_5 * 4, LINE * 3, anchorOf()), run('甲')]), para([run(THIRTY)])]);
    // 第一段一行（6 字宽里的「甲」），第二段接着两行窄的，然后整行
    expect(textsOf(doc)[0]?.map((t) => t.length)).toEqual([1, 6, 6, 10, 8]);
  });

  it('跨页：这一页绕排、下一页没有禁区就按整行排，断点接得上、一个字不丢不重', () => {
    const text = TEN.repeat(15);
    const doc = lay([para([pic(SIZE_5 * 4, LINE * 10, anchorOf()), run(text)])]);
    const [p1, p2] = textsOf(doc);
    expect(p1).toHaveLength(10);
    expect(p1?.every((t) => t.length === 6)).toBe(true);
    expect(p2?.every((t, i, all) => t.length === 10 || i === all.length - 1)).toBe(true);
    expect([...(p1 ?? []), ...(p2 ?? [])].join('')).toBe(text);
    // 对象只在锚点那一页
    expect(doc.pages[1]?.floats).toBeUndefined();
  });

  it('两侧都窄得放不下一个字：推到对象底下再排', () => {
    const anchor = anchorOf({ h: { relativeFrom: 'margin', offset: 100 } });
    const doc = lay([para([pic(1900, LINE * 2, anchor), run(TEN)])]);
    const [first] = linesOf(doc);
    expect(first?.y).toBe(LINE * 2);
    expect(first?.line.x).toBe(0);
  });

  it('bothSides 两侧都放得下时只排一侧并记诊断，wrap="none" 一个字都不让', () => {
    const diagnostics = createDiagnosticSink();
    const centered = anchorOf({ h: { relativeFrom: 'margin', align: 'center' } });
    const doc = lay([para([pic(SIZE_5 * 2, LINE, centered), run(TEN)])], { diagnostics });
    expect(textsOf(doc)[0]?.map((t) => t.length)).toEqual([4, 6]);
    expect(diagnostics.list().map((d) => d.code)).toContain('wrap-both-sides-approximated');

    const none = lay([para([pic(SIZE_5 * 4, LINE * 2, anchorOf({ wrap: 'none' })), run(THIRTY)])]);
    expect(textsOf(none)[0]?.map((t) => t.length)).toEqual([10, 10, 10]);
  });

  it('页眉里的四周型对象伸进版心时，正文也绕着它走', () => {
    const headerFooters: HeaderFooterSource = {
      h: {
        resolved: [
          para([
            pic(SIZE_5 * 4, LINE * 2, anchorOf({ v: { relativeFrom: 'page', offset: 600 } })),
            run('眉'),
          ]),
        ],
      },
    };
    const doc = lay(
      [para([run(THIRTY)])],
      { headerFooters },
      sect({ headers: [{ type: 'default', relId: 'h' }] }),
    );
    expect(textsOf(doc)[0]?.map((t) => t.length)).toEqual([6, 6, 10, 8]);
  });
});
