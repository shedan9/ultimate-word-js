import { twipsToPx } from '@uw/core';
import type { DocumentLayout, FragmentStyle, PageLayout } from '@uw/layout';
import type { DocRange } from '@uw/model';
import { describe, expect, it, vi } from 'vitest';
import type { PageViewport } from './transform.ts';
import { createReadonlyView } from './view.ts';

const STYLE: FragmentStyle = {
  bold: false,
  italic: false,
  color: 'auto',
  underline: 'none',
  strike: false,
  doubleStrike: false,
  vertAlign: 'baseline',
  position: 0,
  scale: 100,
};

/** 两页各两个字，版心偏移特意不为零，防止把页坐标当成版心坐标。 */
function documentLayout(y = 0): DocumentLayout {
  return {
    pages: [0, 1].map(
      (index): PageLayout => ({
        index,
        number: 1,
        sectionIndex: index,
        geometry: { width: 6000, height: 9000, content: { x: 600, y: 900, width: 4800, height: 7200 } },
        blocks: [
          {
            kind: 'paragraph',
            id: `p${index}`,
            y,
            first: true,
            last: true,
            lines: [
              {
                index: 0,
                y,
                line: {
                  start: 0,
                  end: 2,
                  x: 0,
                  width: 600,
                  height: 450,
                  baseline: 350,
                  natural: 450,
                  leaders: [],
                  isLast: true,
                  fragments: [
                    {
                      runId: `r${index}`,
                      contentIndex: 0,
                      offset: 0,
                      font: '宋体',
                      fontSize: 300,
                      script: 'eastAsia',
                      style: STYLE,
                      text: '甲乙',
                      x: 0,
                      width: 600,
                      glyphX: [0, 300],
                    },
                  ],
                },
              },
            ],
          },
        ],
      }),
    ),
  };
}

const range: DocRange = {
  start: { nodeId: 'r0', contentIndex: 0, offset: 0 },
  end: { nodeId: 'r1', contentIndex: 0, offset: 2 },
};

function viewport(zoom = 1, scroll = 0): PageViewport[] {
  return [0, 1].map((page) => ({
    page,
    width: 6000,
    height: 9000,
    matrix: {
      a: twipsToPx(1, zoom),
      b: 0,
      c: 0,
      d: twipsToPx(1, zoom),
      e: 50,
      f: 20 + page * (twipsToPx(9000, zoom) + 24) - scroll,
    },
  }));
}

describe('只读视图连接两次映射', () => {
  it('跨页 range 一次测量，定位字缝与反向矩形一致', () => {
    const measure = vi.fn(() => viewport());
    const view = createReadonlyView(documentLayout(), measure);
    expect(view.rectsOf(range)).toEqual([
      { x: 90, y: 80, width: 40, height: 30 },
      { x: 90, y: 704, width: 40, height: 30 },
    ]);
    expect(measure).toHaveBeenCalledTimes(1);
    expect(view.locate({ clientX: 115, clientY: 710 })).toEqual({
      nodeId: 'r1',
      contentIndex: 0,
      offset: 1,
    });
    expect(view.caretRect({ nodeId: 'r1', contentIndex: 0, offset: 1 })).toEqual({
      x: 110,
      y: 704,
      width: 0,
      height: 30,
    });
  });

  it('滚动与缩放使用新测量，布局完全不变，多视图互不干扰', () => {
    const layout = documentLayout();
    const before = structuredClone(layout);
    let measured = viewport();
    const view = createReadonlyView(layout, () => measured);
    const thumbnail = createReadonlyView(layout, () => viewport(0.5));
    measured = viewport(2, 200);
    expect(view.caretRect(range.start)).toEqual({ x: 130, y: -60, width: 0, height: 60 });
    expect(thumbnail.caretRect(range.start)).toEqual({ x: 70, y: 50, width: 0, height: 15 });
    expect(layout).toEqual(before);
  });

  it('重排后重新建索引，稳定模型位置跟随新几何', () => {
    const view = createReadonlyView(documentLayout(), () => viewport());
    view.update(documentLayout(450));
    expect(view.caretRect(range.start)?.y).toBe(110);
    view.update({ pages: [] });
    expect(view.caretRect(range.start)).toBeNull();
    expect(view.rectsOf(range)).toEqual([]);
  });

  it('未挂载页不返回矩形，空白和未知模型位置不伪造结果', () => {
    const view = createReadonlyView(documentLayout(), () => viewport().slice(0, 1));
    expect(view.rectsOf(range)).toHaveLength(1);
    expect(view.locate({ clientX: 90, clientY: 635 })).toBeNull();
    expect(view.caretRect({ nodeId: 'missing', contentIndex: 0, offset: 0 })).toBeNull();
    const blank = documentLayout();
    for (const page of blank.pages) page.blocks = [];
    view.update(blank);
    expect(view.locate({ clientX: 90, clientY: 80 })).toBeNull();
  });
});
