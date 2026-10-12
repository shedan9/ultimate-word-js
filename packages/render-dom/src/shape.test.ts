/**
 * 形状的画法。
 *
 * 手搭 `PageLayout`，只验几处**画错了才看得出来**的：零高的横线照画（原来零尺寸的对象一律跳过，
 * 红头下那条线整条消失）、直线的方向由翻转给、线不填充、虚线按线宽缩放、
 * 箭头跟着线一起翻转且实心三角把线缩到底边、透明的形状不出元素也不画占位框。
 * 几何公式本身照规范抄，挑两种（圆角半径按短边、梯形上窄下宽）钉住单位换算。
 */
import type { PageLayout, PlacedFloat } from '@uw/layout';
import { describe, expect, it } from 'vitest';
import { buildPage } from './paint.ts';
import { dashArray, shapeOutline } from './shape-geometry.ts';
import type { RElement } from './tree.ts';

const GEOMETRY = {
  width: 11906,
  height: 16838,
  content: { x: 1440, y: 1440, width: 9026, height: 13958 },
};

function float(over: Partial<PlacedFloat> = {}): PlacedFloat {
  return {
    runId: 'r1',
    contentIndex: 0,
    x: 1440,
    y: 2880,
    width: 8820,
    height: 0,
    objectKind: 'drawing',
    behindDoc: false,
    z: 0,
    ...over,
  };
}

function page(floats: PlacedFloat[]): PageLayout {
  return { index: 0, number: 1, sectionIndex: 0, geometry: GEOMETRY, blocks: [], floats };
}

function collect(node: RElement, pred: (n: RElement) => boolean, out: RElement[] = []): RElement[] {
  if (pred(node)) out.push(node);
  for (const c of node.children) collect(c, pred, out);
  return out;
}

const shapes = (svg: RElement) => collect(svg, (n) => (n.attrs.class ?? '').includes('uw-shape'));
const placeholders = (svg: RElement) => collect(svg, (n) => (n.attrs.class ?? '').includes('placeholder'));

describe('直线', () => {
  it('零高的横线照画：一条 <line>，不填充，线宽从 twips 换成 pt', () => {
    const svg = buildPage(
      page([float({ shape: { geometry: 'line', stroke: { color: 'FF0000', width: 45 } } })]),
    );
    const [ln] = shapes(svg);
    expect(ln?.tag).toBe('line');
    expect(ln?.attrs).toMatchObject({
      x1: '72',
      y1: '144',
      x2: '513',
      y2: '144',
      fill: 'none',
      stroke: '#ff0000',
      'stroke-width': '2.25',
    });
    expect(placeholders(svg)).toHaveLength(0);
  });

  it('flipV：线还是外框的对角线，方向由绕中心的翻转给（左下 → 右上）', () => {
    const svg = buildPage(
      page([
        float({
          width: 2000,
          height: 1000,
          shape: { geometry: 'line', flipV: true, stroke: { color: '000000', width: 20 } },
        }),
      ]),
    );
    const [ln] = shapes(svg);
    expect(ln?.attrs).toMatchObject({ x1: '72', y1: '144', x2: '172', y2: '194' });
    // 中心 y = 144 + 25 = 169，翻转 = 以 338 为轴
    expect(ln?.attrs.transform).toBe('translate(0 338) scale(1 -1)');
  });

  it('箭头与线装进同一个组、共用翻转；实心三角把线缩到底边', () => {
    const svg = buildPage(
      page([
        float({
          width: 2000,
          shape: {
            geometry: 'line',
            flipH: true,
            stroke: { color: '000000', width: 20, tail: { type: 'triangle', w: 'med', len: 'med' } },
          },
        }),
      ]),
    );
    const [group] = shapes(svg);
    expect(group?.tag).toBe('g');
    expect(group?.attrs.transform).toBe('translate(244 0) scale(-1 1)');
    const [ln] = collect(group as RElement, (n) => n.tag === 'line');
    // 1pt 线、中号箭头长 3pt：线从 72 画到 172 − 3
    expect(ln?.attrs).toMatchObject({ x1: '72', x2: '169' });
    const [head] = collect(group as RElement, (n) => n.tag === 'path');
    expect(head?.attrs.d).toBe('M172 144 L169 145.5 L169 142.5 Z');
    expect(head?.attrs.fill).toBe('#000000');
  });

  it('虚线图案按线宽缩放', () => {
    expect(dashArray('sysDash', 2)).toBe('6 2');
    expect(dashArray('lgDashDot', 1)).toBe('8 3 1 3');
    expect(dashArray('solid', 1)).toBeUndefined();
    const svg = buildPage(
      page([float({ shape: { geometry: 'line', stroke: { color: '000000', width: 40, dash: 'dash' } } })]),
    );
    expect(shapes(svg)[0]?.attrs['stroke-dasharray']).toBe('8 6');
  });
});

describe('闭合形状', () => {
  it('椭圆、旋转绕外框中心', () => {
    const svg = buildPage(
      page([
        float({
          width: 2000,
          height: 1000,
          shape: { geometry: 'ellipse', rotation: 30, fill: 'FFFF00' },
        }),
      ]),
    );
    const [e] = shapes(svg);
    expect(e?.tag).toBe('ellipse');
    expect(e?.attrs).toMatchObject({ cx: '122', cy: '169', rx: '50', ry: '25', fill: '#ffff00' });
    expect(e?.attrs.transform).toBe('rotate(30 122 169)');
  });

  it('圆角半径 = 短边 × adj；缺省 16667', () => {
    const box = { x: 0, y: 0, width: 100, height: 60 };
    expect(shapeOutline({ geometry: 'roundRect' }, box).attrs.rx).toBe('10');
    expect(shapeOutline({ geometry: 'roundRect', adjust: 50000 }, box).attrs.rx).toBe('30');
    // 超过上限夹回 50000
    expect(shapeOutline({ geometry: 'roundRect', adjust: 90000 }, box).attrs.rx).toBe('30');
  });

  it('梯形上窄下宽（Word 2007 起），菱形四个顶点在四条边的中点', () => {
    const box = { x: 0, y: 0, width: 100, height: 40 };
    expect(shapeOutline({ geometry: 'trapezoid' }, box).attrs.d).toBe('M0 40 L10 0 L90 0 L100 40 Z');
    expect(shapeOutline({ geometry: 'diamond' }, box).attrs.d).toBe('M50 0 L100 20 L50 40 L0 20 Z');
  });

  it('透明的形状不出元素，也不画占位框', () => {
    const svg = buildPage(page([float({ height: 1000, shape: { geometry: 'rect' } })]));
    expect(shapes(svg)).toHaveLength(0);
    expect(placeholders(svg)).toHaveLength(0);
  });

  it('认不出几何的形状（没有 shape）照旧画占位框', () => {
    const svg = buildPage(page([float({ height: 1000, graphic: 'wordprocessingShape' })]));
    expect(placeholders(svg)).toHaveLength(1);
  });
});

describe('连接符', () => {
  it('肘形：横 → 竖 → 横，拐在 adjust 占宽的比例上，不填充', () => {
    const svg = buildPage(
      page([
        float({
          width: 2000,
          height: 1000,
          shape: {
            geometry: 'bentConnector3',
            adjust: 25000,
            fill: 'FFFFFF',
            stroke: { color: '000000', width: 20 },
          },
        }),
      ]),
    );
    const [p] = shapes(svg);
    expect(p?.tag).toBe('path');
    expect(p?.attrs.d).toBe('M72 144 L97 144 L97 194 L172 194');
    expect(p?.attrs.fill).toBe('none');
  });

  it('箭头顺着最后一段（横的）而不是外框的对角线；零高时跳过缩成一点的那一竖', () => {
    const tail = { type: 'triangle' as const, w: 'med' as const, len: 'med' as const };
    const svg = buildPage(
      page([
        float({
          width: 2000,
          height: 1000,
          shape: { geometry: 'bentConnector3', stroke: { color: '000000', width: 20, tail } },
        }),
        float({
          width: 2000,
          height: 0,
          shape: { geometry: 'bentConnector3', stroke: { color: '000000', width: 20, head: tail } },
        }),
      ]),
    );
    const [bent, flat] = shapes(svg);
    const [line] = collect(bent as RElement, (n) => n.tag === 'path' && n.attrs.fill === 'none');
    // 末段从 (122, 194) 往右到 172，缩回 3pt 到箭头底边
    expect(line?.attrs.d).toBe('M72 144 L122 144 L122 194 L169 194');
    const [arrow] = collect(bent as RElement, (n) => n.tag === 'path' && n.attrs.fill === '#000000');
    expect(arrow?.attrs.d).toBe('M172 194 L169 195.5 L169 192.5 Z');
    const [head] = collect(flat as RElement, (n) => n.tag === 'path' && n.attrs.fill === '#000000');
    expect(head?.attrs.d).toBe('M72 144 L75 142.5 L75 145.5 Z');
  });

  it('曲线：两段三次贝塞尔在拐点那一竖的正中接上', () => {
    const svg = buildPage(
      page([
        float({
          width: 2000,
          height: 1000,
          shape: { geometry: 'curvedConnector3', stroke: { color: '000000', width: 20 } },
        }),
      ]),
    );
    expect(shapes(svg)[0]?.attrs.d).toBe('M72 144 C97 144 122 156.5 122 169 C122 181.5 147 194 172 194');
  });
});

describe('组合', () => {
  const groups = (svg: RElement) =>
    collect(svg, (n) => (n.attrs.class ?? '').split(' ').includes('uw-group'));

  it('成员相对组合外框画，认不出的成员画占位框，组合自己不画占位框', () => {
    const svg = buildPage(
      page([
        float({
          width: 4000,
          height: 2000,
          graphic: 'wordprocessingGroup',
          group: [
            {
              x: 0,
              y: 0,
              width: 2000,
              height: 1000,
              shape: { geometry: 'rect', fill: 'FFFFFF', stroke: { color: '000000', width: 20 } },
            },
            { x: 2000, y: 1000, width: 2000, height: 1000 },
          ],
        }),
      ]),
    );
    const [g] = groups(svg);
    expect(g?.tag).toBe('g');
    expect(g?.attrs['data-run']).toBe('r1');
    const [r] = shapes(g as RElement);
    expect(r?.attrs).toMatchObject({ x: '72', y: '144', width: '100', height: '50' });
    const [ph] = placeholders(svg);
    expect(ph?.attrs).toMatchObject({ x: '172', y: '194', width: '100', height: '50' });
  });

  it('成员图片按转之前的框画，旋转翻转绕成员中心（不像独立图片那样把长宽比缩回去）', () => {
    const svg = buildPage(
      page([
        float({
          width: 4000,
          height: 2000,
          group: [
            { x: 0, y: 0, width: 2000, height: 1000, image: { id: 'rId9', relId: 'rId9', rotation: 90 } },
          ],
        }),
      ]),
      { imageHref: (id) => `img:${id}` },
    );
    const [img] = collect(svg, (n) => n.tag === 'image');
    expect(img?.attrs).toMatchObject({ href: 'img:rId9', x: '72', y: '144', width: '100', height: '50' });
    expect(img?.attrs.transform).toBe('rotate(90 122 169)');
  });

  it('全是横线的组合零高也照画', () => {
    const svg = buildPage(
      page([
        float({
          width: 4000,
          height: 0,
          group: [
            {
              x: 0,
              y: 0,
              width: 4000,
              height: 0,
              shape: { geometry: 'line', stroke: { color: '000000', width: 20 } },
            },
          ],
        }),
      ]),
    );
    expect(groups(svg)).toHaveLength(1);
    expect(shapes(svg)[0]?.tag).toBe('line');
  });
});
