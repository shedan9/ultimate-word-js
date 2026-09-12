import type { IndexedLine, LineFragment, PageLayout } from '@uw/layout';
import type { RElement } from '@uw/render-dom';
import { describe, expect, it } from 'vitest';
import { buildTextLayer } from './text-layer.ts';

const page: PageLayout = {
  index: 2,
  number: 1,
  sectionIndex: 1,
  blocks: [],
  geometry: { width: 6000, height: 9000, content: { x: 600, y: 900, width: 4800, height: 7200 } },
};
function fragment(text: string, over: Partial<LineFragment> = {}): LineFragment {
  return {
    runId: 'r1',
    contentIndex: 0,
    offset: 0,
    font: '宋体',
    fontSize: 300,
    script: 'eastAsia',
    text,
    x: 0,
    width: 600,
    glyphX: [0, 300],
    style: {
      bold: false,
      italic: false,
      color: 'auto',
      underline: 'none',
      strike: false,
      doubleStrike: false,
      vertAlign: 'baseline',
      position: 0,
      scale: 100,
    },
    ...over,
  };
}
function line(fragments: LineFragment[], over: Partial<IndexedLine> = {}): IndexedLine {
  return {
    page: 2,
    originX: 600,
    top: 900,
    repeated: false,
    line: {
      start: 0,
      end: 2,
      x: 0,
      width: 600,
      height: 450,
      baseline: 350,
      natural: 450,
      fragments,
      leaders: [],
      isLast: true,
    },
    ...over,
  };
}
function texts(node: RElement): RElement[] {
  return node.tag === 'tspan' ? [node] : node.children.flatMap(texts);
}

describe('常驻原生文字层', () => {
  it('编号与重复表头排除，域结果和跨内容片的正文保留', () => {
    const tree = buildTextLayer(page, [
      line([
        fragment('一、', { numbering: true }),
        fragment('正文'),
        fragment('7', { field: true, contentIndex: -1, offset: -1 }),
      ]),
      line([fragment('重复表头')], { repeated: true }),
      line([fragment('第二片', { contentIndex: 1, offset: 3 })]),
      line([fragment('其他页')], { page: 3 }),
    ]);
    expect(texts(tree).map((n) => n.text)).toEqual(['正文', '7', '第二片']);
    expect(texts(tree).at(-1)?.attrs['data-content-index']).toBe('1');
    expect(tree.attrs['data-viewport-page']).toBe('2');
  });
  it('直接使用逐字坐标和实际基线，透明文字不改变布局', () => {
    const input = [line([fragment('甲乙')])];
    const before = structuredClone(input);
    const text = texts(buildTextLayer(page, input))[0];
    expect(text?.attrs.x).toBe('30 45');
    expect(text?.attrs.y).toBe('62.5');
    expect(text?.attrs.fill).toBe('transparent');
    expect(text?.attrs['xml:space']).toBe('preserve');
    expect(input).toEqual(before);
  });
  it('压缩、上下标和升降沿用绘制规则', () => {
    const base = fragment('甲乙');
    const text = texts(
      buildTextLayer(page, [
        line([{ ...base, style: { ...base.style, scale: 50, position: 100 }, glyphX: [0, 150] }]),
      ]),
    )[0];
    expect(text?.attrs.transform).toBeUndefined();
    expect(text?.attrs.x).toBe('30 37.5');
    expect(text?.attrs.lengthAdjust).toBe('spacingAndGlyphs');
    expect(text?.attrs.y).toBe('57.5');
  });
  it('关闭文字层后仍保留页面坐标 SVG，空白页没有假文字', () => {
    const disabled = buildTextLayer(page, [line([fragment('正文')])], {}, false);
    expect(disabled.children).toEqual([]);
    expect(disabled.attrs['aria-hidden']).toBe('true');
    expect(disabled.attrs.viewBox).toBe('0 0 300 450');
    expect(buildTextLayer(page, []).children).toEqual([]);
  });
  it('图片说明保留在辅助技术中，不混入复制文字', () => {
    const imageLine = line([]);
    imageLine.line.objects = [
      {
        runId: 'image',
        contentIndex: 0,
        x: 150,
        width: 300,
        height: 200,
        raise: 50,
        objectKind: 'drawing',
        alt: '流程图',
      },
    ];
    const tree = buildTextLayer(page, [imageLine]);
    expect(texts(tree)).toEqual([]);
    expect(tree.children[0]?.attrs).toMatchObject({
      role: 'img',
      'aria-label': '流程图',
      x: '37.5',
      y: '50',
    });
  });
});
