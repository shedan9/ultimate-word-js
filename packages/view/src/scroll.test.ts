import type { DocumentLayout, FragmentStyle } from '@uw/layout';
import { buildLayoutIndex } from '@uw/layout';
import { describe, expect, it } from 'vitest';
import { scrollTargetRect } from './scroll.ts';

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
/** 三页各一行、各一个 run，版心偏移不为零，防止把页坐标当成版心坐标 */
const layout: DocumentLayout = {
  pages: [0, 1, 2].map((page) => ({
    index: page,
    number: page + 1,
    sectionIndex: 0,
    geometry: { width: 9000, height: 9000, content: { x: 900, y: 900, width: 7200, height: 7200 } },
    blocks: [
      {
        kind: 'paragraph',
        id: `p${page}`,
        y: 0,
        first: true,
        last: true,
        lines: [
          {
            index: 0,
            y: 0,
            line: {
              start: 0,
              end: 4,
              x: 0,
              width: 1200,
              height: 450,
              baseline: 350,
              natural: 450,
              leaders: [],
              isLast: true,
              fragments: [
                {
                  runId: `r${page}`,
                  contentIndex: 0,
                  offset: 0,
                  font: 'serif',
                  fontSize: 300,
                  script: 'eastAsia',
                  style,
                  text: '甲乙丙丁',
                  x: 0,
                  width: 1200,
                  glyphX: [0, 300, 600, 900],
                },
              ],
            },
          },
        ],
      },
    ],
  })),
};

describe('scrollTargetRect', () => {
  const index = buildLayoutIndex(layout);
  const at = (run: number, offset: number) => ({ nodeId: `r${run}`, contentIndex: 0, offset });

  it('位置取光标矩形，range 取首行矩形（不是包围盒中心）', () => {
    expect(scrollTargetRect(index, at(1, 2))).toEqual({ page: 1, x: 1500, y: 900, width: 0, height: 450 });
    const range = scrollTargetRect(index, { start: at(0, 3), end: at(2, 2) });
    expect(range).toEqual({ page: 0, x: 1800, y: 900, width: 300, height: 450 });
  });

  it('排不出来的位置没有矩形', () => {
    expect(scrollTargetRect(index, at(9, 0))).toBeUndefined();
    expect(scrollTargetRect(index, { start: at(9, 0), end: at(9, 1) })).toBeUndefined();
  });
});
