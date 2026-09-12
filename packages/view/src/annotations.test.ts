import type { DocumentLayout, FragmentStyle } from '@uw/layout';
import { buildLayoutIndex } from '@uw/layout';
import { describe, expect, it } from 'vitest';
import { anchorRect, pageRect, placeOverlay } from './annotations.ts';

const style: FragmentStyle = {
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
const layout: DocumentLayout = {
  pages: [0, 1].map((page) => ({
    index: page,
    number: 1,
    sectionIndex: 0,
    geometry: { width: 6000, height: 9000, content: { x: 600, y: 900, width: 4800, height: 7200 } },
    blocks: [
      {
        kind: 'paragraph',
        id: 'p',
        y: 0,
        first: page === 0,
        last: page === 1,
        lines: [
          {
            index: page,
            y: page * 450,
            line: {
              start: page * 2,
              end: page * 2 + 2,
              x: 300,
              width: 600,
              height: 450,
              baseline: 350,
              natural: 450,
              leaders: [],
              isLast: page === 1,
              fragments: [
                {
                  runId: 'r',
                  contentIndex: 0,
                  offset: page * 2,
                  font: 'serif',
                  fontSize: 300,
                  script: 'eastAsia',
                  style,
                  text: '甲乙',
                  x: 300,
                  width: 600,
                  glyphX: [300, 600],
                },
              ],
            },
          },
        ],
      },
    ],
  })),
};

describe('装饰与批注锚点', () => {
  it('同一 run 跨页时，断行位置跟随下一行而非 run 首行', () => {
    const index = buildLayoutIndex(layout);
    const position = { nodeId: 'r', contentIndex: 0, offset: 2 };
    expect(anchorRect(index, position, 'right-of-line')).toEqual({
      page: 1,
      x: 900,
      y: 1350,
      width: 600,
      height: 450,
    });
    expect(anchorRect(index, position, 'inline')).toEqual({
      page: 1,
      x: 900,
      y: 1350,
      width: 0,
      height: 450,
    });
  });
  it('缺失模型位置不吸附到别的文字', () => {
    expect(
      anchorRect(buildLayoutIndex(layout), { nodeId: 'missing', contentIndex: 0, offset: 0 }, 'above'),
    ).toBeUndefined();
  });
  it('页面留白与非整数缩放采用 SVG meet 比例', () => {
    expect(pageRect({ page: 0, x: 600, y: 900, width: 600, height: 450 }, 6000, 9000, 500, 600)).toEqual({
      x: 90,
      y: 60,
      width: 40,
      height: 30,
    });
    const rect = pageRect({ page: 0, x: 600, y: 900, width: 600, height: 450 }, 6000, 9000, 493.2, 739.8);
    expect(rect.width).toBeCloseTo(49.32);
    expect(rect.x).toBeCloseTo(49.32);
  });
  it('四个 placement 使用行末或字缝，offset 不混入缩放', () => {
    const rect = { x: 40, y: 60, width: 100, height: 30 };
    const size = { width: 80, height: 20 };
    const offset = { x: 8, y: -2 };
    expect(placeOverlay(rect, size, 'right-of-line', offset)).toEqual({ x: 148, y: 58 });
    expect(placeOverlay(rect, size, 'above', offset)).toEqual({ x: 48, y: 38 });
    expect(placeOverlay(rect, size, 'below', offset)).toEqual({ x: 48, y: 88 });
    expect(placeOverlay(rect, size, 'inline', offset)).toEqual({ x: 48, y: 58 });
  });
});
