/**
 * 文本框的内容：摞在外框里（减内边距、按 vAlign 下移），浮动的跟着 `placeFloats` 的外框，
 * 内嵌的借住进 `page.floats`（带 `inline`）。版心与 object.test.ts 同一套（10 字宽 × 3 行高）。
 */
import type { DrawingAnchor, ResolvedBlock, ResolvedBody, SectionProps, TextBoxRef } from '@uw/model';
import { DEFAULT_SECTION_PROPS, DEFAULT_SETTINGS } from '@uw/model';
import { describe, expect, it } from 'vitest';
import type { LayoutDocumentOptions } from './page.ts';
import { layoutDocument } from './page.ts';
import { fakeMeasurer, NO_GRID, para, run, runOf, SIZE_5 } from './test-fixtures.ts';

const EA_LINE = SIZE_5 * 1.3;

function sect(): SectionProps {
  return {
    ...structuredClone(DEFAULT_SECTION_PROPS),
    page: { width: 3300, height: 2019, orientation: 'portrait' },
    margin: { top: 600, right: 600, bottom: 600, left: 600, header: 0, footer: 0, gutter: 0 },
    docGrid: NO_GRID,
  };
}

function body(blocks: ResolvedBlock[]): ResolvedBody {
  return { sections: [{ id: 's0', props: sect(), blocks }] };
}

const ANCHOR: DrawingAnchor = {
  wrap: 'none',
  behindDoc: false,
  z: 7,
  h: { relativeFrom: 'page', offset: 300 },
  v: { relativeFrom: 'page', offset: 400 },
  dist: { top: 0, bottom: 0, left: 0, right: 0 },
};

function ref(over: Partial<TextBoxRef> = {}): TextBoxRef {
  return { id: 'tb0', inset: { left: 100, top: 50, right: 100, bottom: 50 }, vAlign: 'top', ...over };
}

/** 一个 w × h 的文本框（浮动时带 ANCHOR） */
function box(width: number, height: number, textBox: TextBoxRef, anchor?: DrawingAnchor) {
  return runOf([
    {
      kind: 'object' as const,
      objectKind: 'drawing' as const,
      width,
      height,
      textBox,
      shape: { stroke: { color: '000000', width: 10 } },
      ...(anchor === undefined ? {} : { anchor }),
    },
  ]);
}

/** 框里两段各三个字 */
const CONTENT = { tb0: { resolved: [para([run('一二三')]), para([run('四五六')])] } };

function opts(over: Partial<LayoutDocumentOptions> = {}): LayoutDocumentOptions {
  return { measurer: fakeMeasurer(), settings: DEFAULT_SETTINGS, textBoxes: CONTENT, ...over };
}

describe('文本框的内容', () => {
  it('浮动文本框：内容区 = 外框减内边距，块从内容区顶往下摞', () => {
    const doc = layoutDocument(body([para([box(1000, 800, ref(), ANCHOR), run('正文')])]), opts());
    const [f] = doc.pages[0]?.floats ?? [];
    expect(f).toMatchObject({ x: 300, y: 400, width: 1000, height: 800, shape: { stroke: { width: 10 } } });
    expect(f?.textBox).toMatchObject({ x: 400, y: 450, width: 800, height: 2 * EA_LINE });
    const blocks = f?.textBox?.blocks ?? [];
    expect(blocks.map((b) => b.kind === 'paragraph' && b.lines.map((l) => l.y))).toEqual([[0], [EA_LINE]]);
  });

  it('vAlign 居中 / 靠下把整块往下挪；装不下时顶着上边（不往上溢出）', () => {
    const center = layoutDocument(body([para([box(1000, 800, ref({ vAlign: 'center' }), ANCHOR)])]), opts())
      .pages[0]?.floats?.[0]?.textBox;
    const room = 800 - 100;
    expect(center?.y).toBe(450 + (room - 2 * EA_LINE) / 2);

    const bottom = layoutDocument(body([para([box(1000, 800, ref({ vAlign: 'bottom' }), ANCHOR)])]), opts())
      .pages[0]?.floats?.[0]?.textBox;
    expect(bottom?.y).toBe(450 + room - 2 * EA_LINE);

    const tight = layoutDocument(body([para([box(1000, 200, ref({ vAlign: 'bottom' }), ANCHOR)])]), opts())
      .pages[0]?.floats?.[0]?.textBox;
    expect(tight?.y).toBe(450);
  });

  it('内容宽决定断行：三个字放不下就折成两行', () => {
    // 外框 = 两个字宽 + 左右内边距，内容区正好两个字宽
    const doc = layoutDocument(
      body([para([box(2 * SIZE_5 + 200, 1000, ref(), ANCHOR)])]),
      opts({ textBoxes: { tb0: { resolved: [para([run('一二三')])] } } }),
    );
    const [p] = doc.pages[0]?.floats?.[0]?.textBox?.blocks ?? [];
    expect(
      p?.kind === 'paragraph' && p.lines.map((l) => l.line.fragments.map((x) => x.text).join('')),
    ).toEqual(['一二', '三']);
  });

  it('内嵌文本框借住进 page.floats（带 inline），外框按行算：基线 − 高', () => {
    const doc = layoutDocument(body([para([run('一'), box(400, 300, ref())])]), opts());
    const [f] = doc.pages[0]?.floats ?? [];
    expect(f?.inline).toBe(true);
    // 行首一个字，框的 x 接在它后面
    expect(f?.x).toBe(600 + SIZE_5);
    expect(f?.textBox?.x).toBe(600 + SIZE_5 + 100);
    // 外框没有带过来：内嵌的外框由行画
    expect(f?.shape).toBeUndefined();
  });

  it('没传内容（或内容里没有这个 id）只剩外框，不报错', () => {
    const doc = layoutDocument(body([para([box(1000, 800, ref(), ANCHOR)])]), opts({ textBoxes: {} }));
    expect(doc.pages[0]?.floats?.[0]?.textBox).toBeUndefined();
    const inline = layoutDocument(body([para([box(400, 300, ref())])]), opts({ textBoxes: {} }));
    expect(inline.pages[0]?.floats).toBeUndefined();
  });

  it('同一个框出现在每一页的页眉里：内容只摞一次，各页共用', () => {
    const header = { rIdH: { resolved: [para([box(1000, 800, ref(), ANCHOR)])] } };
    const s = { ...sect(), headers: [{ type: 'default' as const, relId: 'rIdH' }] };
    const many = Array.from({ length: 8 }, () => para([run('一二三四五六七八九十')]));
    const doc = layoutDocument(
      { sections: [{ id: 's0', props: s, blocks: many }] },
      opts({ headerFooters: header }),
    );
    expect(doc.pages.length).toBeGreaterThan(1);
    const boxes = doc.pages.map((p) => p.floats?.[0]?.textBox);
    expect(boxes.every((b) => b !== undefined)).toBe(true);
    expect(boxes[1]?.blocks).toBe(boxes[0]?.blocks);
  });
});
