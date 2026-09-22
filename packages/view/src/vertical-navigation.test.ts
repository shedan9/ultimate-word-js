import type { DocumentLayout, FragmentStyle, LineFragment } from '@uw/layout';
import { buildLayoutIndex } from '@uw/layout';
import { describe, expect, it } from 'vitest';
import { moveVertically } from './vertical-navigation.ts';

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
const pos = (nodeId: string, offset = 0) => ({ nodeId, contentIndex: 0, offset });
const fragment = (runId: string, text: string, x = 0, offset = 0): LineFragment => ({
  runId,
  text,
  contentIndex: 0,
  offset,
  x,
  width: [...text].length * 100,
  glyphX: [...text].map((_, i) => x + i * 100),
  font: '宋体',
  fontSize: 200,
  script: 'eastAsia',
  style,
});

interface Row {
  id: string;
  text?: string;
  top: number;
  x?: number;
  height?: number;
  fragments?: LineFragment[];
  frame?: 'header' | 'footer';
  repeated?: boolean;
}

function index(pages: Row[][]) {
  const layout: DocumentLayout = {
    pages: pages.map((rows, page) => ({
      index: page,
      number: page + 1,
      sectionIndex: 0,
      geometry: { width: 6000, height: 9000, content: { x: 600, y: 900, width: 4800, height: 7200 } },
      blocks: rows.map((row) => {
        const fragments = row.fragments ?? (row.text ? [fragment(row.id, row.text, row.x)] : []);
        const height = row.height ?? 200;
        return {
          kind: 'paragraph',
          id: `p-${row.id}`,
          y: row.top,
          first: true,
          last: true,
          lines: [
            {
              index: 0,
              y: row.top,
              line: {
                start: 0,
                end: 0,
                x: row.x ?? 0,
                width: fragments.reduce((sum, f) => sum + f.width, 0),
                height,
                baseline: height * 0.8,
                natural: height,
                leaders: [],
                isLast: true,
                fragments,
                ...(fragments.length ? {} : { emptyPosition: pos(row.id) }),
              },
            },
          ],
        };
      }),
    })),
  };
  const result = buildLayoutIndex(layout);
  // 标记不参与坐标计算，用真实索引检验导航是否排除重复内容。
  for (const [i, row] of pages.flat().entries()) {
    const line = result.lines[i];
    if (line && row.frame) line.frame = row.frame;
    if (line && row.repeated) line.repeated = true;
  }
  return result;
}

describe('上下行导航', () => {
  it('跨页跳过页眉页脚、空白补页、重复表头与无源域结果', () => {
    const idx = index([
      [
        { id: 'a', text: '甲乙丙', top: 5000 },
        { id: 'footer', text: '页脚', top: 6000, frame: 'footer' },
      ],
      [],
      [
        { id: 'header', text: '页眉', top: 0, frame: 'header' },
        { id: 'repeat', text: '表头', top: 200, repeated: true },
        { id: 'field', top: 400, fragments: [fragment('field', '1', 0, -1)] },
        { id: 'b', text: '丁戊己', top: 800 },
      ],
    ]);
    expect(moveVertically(idx, pos('a', 2), 'down')).toEqual({ position: pos('b', 2), x: 800 });
    expect(moveVertically(idx, pos('b', 2), 'up')?.position).toEqual(pos('a', 2));
  });

  it('行高和段间距不同仍逐行移动，经过短行后恢复目标横坐标', () => {
    const idx = index([
      [
        { id: 'a', text: '甲乙丙丁', top: 0, height: 500 },
        { id: 'b', text: '短', top: 2000, height: 100 },
        { id: 'c', text: '一二三四', top: 3000 },
      ],
    ]);
    const first = moveVertically(idx, pos('a', 3), 'down');
    expect(first).toEqual({ position: pos('b', 1), x: 900 });
    if (!first) throw new Error('缺少下一行');
    expect(moveVertically(idx, first.position, 'down', first.x)?.position).toEqual(pos('c', 3));
    expect(moveVertically(idx, first.position, 'down')?.position).toEqual(pos('c', 1));
  });

  it('同一高度并排的格内行按横向距离选目标，不横跳到同一排', () => {
    const idx = index([
      [
        { id: 'left1', text: '甲乙', top: 0 },
        { id: 'left2', text: '丙丁', top: 500 },
        { id: 'right1', text: '一二', top: 0, x: 2000 },
        { id: 'right2', text: '三四', top: 500, x: 2000 },
      ],
    ]);
    expect(moveVertically(idx, pos('right1', 1), 'down')?.position).toEqual(pos('right2', 1));
    expect(moveVertically(idx, pos('left2', 1), 'up')?.position).toEqual(pos('left1', 1));
  });

  it('空段落可停留，文档两端与未知位置不伪造目标', () => {
    const idx = index([
      [
        { id: 'a', text: '甲', top: 0 },
        { id: 'empty', top: 400 },
      ],
    ]);
    expect(moveVertically(idx, pos('a'), 'down')?.position).toEqual(pos('empty'));
    expect(moveVertically(idx, pos('empty'), 'up')?.position).toEqual(pos('a'));
    expect(moveVertically(idx, pos('empty'), 'down')).toBeUndefined();
    expect(moveVertically(idx, pos('a'), 'up')).toBeUndefined();
    expect(moveVertically(idx, pos('missing'), 'down')).toBeUndefined();
  });

  it('软换行字缝归下一行时，向上不能停回原行', () => {
    const idx = index([
      [
        { id: 'a', top: 0, fragments: [fragment('r', '甲乙')] },
        { id: 'b', top: 300, fragments: [fragment('r', '丙丁戊', 0, 2)] },
      ],
    ]);
    const moved = moveVertically(idx, pos('r', 5), 'up');
    expect(moved?.position).toEqual(pos('r', 1));
    if (!moved) throw new Error('缺少上一行');
    expect(idx.caretRect(moved.position)?.y).toBe(900);
    expect(moveVertically(idx, moved.position, 'down', moved.x)?.position).toEqual(pos('r', 5));
  });

  it('跨样式字素内部不产生落点，组合音标与 ZWJ emoji 保持完整', () => {
    const idx = index([
      [
        { id: 'a', text: '一二三四五六', top: 0 },
        {
          id: 'b',
          top: 400,
          fragments: [
            fragment('r1', 'e'),
            fragment('r2', '\u0301👩', 100),
            fragment('r3', '\u200d💻好', 300),
          ],
        },
      ],
    ]);
    expect(moveVertically(idx, pos('a', 1), 'down')?.position).toEqual(pos('r1'));
    expect(moveVertically(idx, pos('a', 3), 'down')?.position).toEqual(pos('r2', 1));
  });
});
