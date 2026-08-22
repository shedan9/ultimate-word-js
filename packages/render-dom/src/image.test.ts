/**
 * 图片的画法。
 *
 * 与 paint.test.ts 同一路数：手搭 `PageLayout`，只验「布局给的 twips 有没有原封不动
 * 翻译成 pt」以及那几处**画错了才看得出来**的地方 —— 裁剪要放大后再切、
 * 旋转 90° 要横竖对调、画不出来的要留一个尺寸正确的框。
 */
import type { LineLayout, LineObject, PageLayout, PlacedFloat } from '@uw/layout';
import { describe, expect, it } from 'vitest';
import { imageHrefOf, imageHrefResolver } from './image.ts';
import { buildPage } from './paint.ts';
import type { RElement } from './tree.ts';

const GEOMETRY = {
  width: 11906,
  height: 16838,
  content: { x: 1440, y: 1440, width: 9026, height: 13958 },
};

/** 1×1 的透明 PNG，够短，能整个写进期望值 */
const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);

const IMAGES = {
  rId5: { contentType: 'image/png', bytes: PNG },
  emf: { contentType: 'image/x-emf', bytes: PNG },
  ext: { contentType: 'image/gif', url: 'https://example.com/a.gif' },
};

const href = imageHrefResolver(IMAGES);

/** `image: null` = 一个画不出来的对象（图表 / 形状）—— `exactOptionalPropertyTypes`
 *  不许把 `undefined` 当成「不写」，所以用 null 表示「明确没有」 */
function object(
  over: Omit<Partial<LineObject>, 'image'> & { image?: LineObject['image'] | null } = {},
): LineObject {
  const { image, ...rest } = over;
  return {
    runId: 'r1',
    contentIndex: 0,
    x: 0,
    width: 2880, // 144pt
    height: 1440, // 72pt
    objectKind: 'drawing',
    ...(image === null ? {} : { image: image ?? { id: 'rId5', relId: 'rId5' } }),
    ...rest,
  };
}

function line(objects: LineObject[]): LineLayout {
  return {
    start: 0,
    end: 1,
    x: 0,
    width: 0,
    height: 1440,
    baseline: 1440,
    natural: 1440,
    fragments: [],
    leaders: [],
    objects,
    isLast: true,
  };
}

function pageWith(objects: LineObject[], floats?: PlacedFloat[]): PageLayout {
  const p: PageLayout = {
    index: 0,
    number: 1,
    sectionIndex: 0,
    geometry: GEOMETRY,
    blocks: [
      {
        kind: 'paragraph',
        id: 'p1',
        y: 0,
        lines: [{ index: 0, y: 0, line: line(objects) }],
        first: true,
        last: true,
      },
    ],
  };
  if (floats !== undefined) p.floats = floats;
  return p;
}

function collect(node: RElement, tag: string, out: RElement[] = []): RElement[] {
  if (node.tag === tag) out.push(node);
  for (const c of node.children) collect(c, tag, out);
  return out;
}

describe('href', () => {
  it('位图变成 data URI，外链原样给地址', () => {
    expect(imageHrefOf(IMAGES.rId5)).toBe('data:image/png;base64,iVBORw0KGgo=');
    expect(imageHrefOf(IMAGES.ext)).toBe('https://example.com/a.gif');
  });

  it('EMF / WMF 返回 undefined —— 浏览器画不出来，给了 href 只会得到一个破图', () => {
    expect(imageHrefOf(IMAGES.emf)).toBeUndefined();
  });

  it('同一张图只编码一次', () => {
    const resolver = imageHrefResolver(IMAGES);
    expect(resolver('rId5')).toBe(resolver('rId5'));
    expect(resolver('没这张')).toBeUndefined();
  });
});

describe('内嵌图', () => {
  it('画在基线之上：y = 基线 − 高度，尺寸从 twips 换成 pt', () => {
    const svg = buildPage(pageWith([object()]), { imageHref: href });
    const [img] = collect(svg, 'image');
    // 行高 1440、基线 1440 → 图的顶边在版心顶（`<g>` 已经平移过，所以这里是 0）
    expect(img?.attrs).toMatchObject({ x: '0', y: '0', width: '144', height: '72' });
    expect(img?.attrs.href).toBe('data:image/png;base64,iVBORw0KGgo=');
    // 外框比例是用户拖出来的，可以与图片本身不一致 —— 不许浏览器替我们「保持比例」
    expect(img?.attrs.preserveAspectRatio).toBe('none');
  });

  it('裁剪：整张图放大后再切回外框，切的框还是原来那个', () => {
    const svg = buildPage(
      pageWith([
        object({
          image: { id: 'rId5', relId: 'rId5', crop: { left: 0.25, top: 0, right: 0.25, bottom: 0.5 } },
        }),
      ]),
      { imageHref: href },
    );
    const [img] = collect(svg, 'image');
    // 左右各裁 25% → 只剩一半，整张图要画成两倍宽；上下裁掉一半 → 两倍高
    expect(img?.attrs.width).toBe('288');
    expect(img?.attrs.height).toBe('144');
    // 往左挪掉被裁掉的那 25%（288 × 0.25）
    expect(img?.attrs.x).toBe('-72');
    expect(img?.attrs.y).toBe('0');
    const [clip] = collect(svg, 'clipPath');
    expect(collect(clip as RElement, 'rect')[0]?.attrs).toMatchObject({ width: '144', height: '72' });
    expect(img?.attrs['clip-path']).toBe(`url(#${clip?.attrs.id})`);
  });

  it('两张裁过的图不共用 clipPath id —— 共用会让后一张按前一张的框去切', () => {
    const crop = { left: 0.1, top: 0, right: 0, bottom: 0 };
    const svg = buildPage(
      pageWith([
        object({ image: { id: 'rId5', relId: 'rId5', crop } }),
        object({ x: 2880, image: { id: 'rId5', relId: 'rId5', crop } }),
      ]),
      { imageHref: href },
    );
    const ids = collect(svg, 'clipPath').map((c) => c.attrs.id);
    expect(new Set(ids).size).toBe(2);
  });

  it('旋转 90°：绕中心转，并且把长宽比缩回去（extent 是转完的外接矩形）', () => {
    const svg = buildPage(pageWith([object({ image: { id: 'rId5', relId: 'rId5', rotation: 90 } })]), {
      imageHref: href,
    });
    const t = collect(svg, 'image')[0]?.attrs.transform ?? '';
    expect(t).toContain('rotate(90 72 36)');
    expect(t).toContain('scale(0.5 2)');
  });

  it('翻转绕中心做', () => {
    const svg = buildPage(pageWith([object({ image: { id: 'rId5', relId: 'rId5', flipH: true } })]), {
      imageHref: href,
    });
    expect(collect(svg, 'image')[0]?.attrs.transform).toBe('translate(144 0) scale(-1 1)');
  });
});

describe('画不出来的对象', () => {
  it('图表画一个尺寸正确的虚线框，可选文本进 <title>', () => {
    const svg = buildPage(pageWith([object({ image: null, graphic: 'chart', alt: '季度图表' })]), {
      imageHref: href,
    });
    expect(collect(svg, 'image')).toHaveLength(0);
    const rect = collect(svg, 'rect').find((r) => r.attrs.class?.includes('uw-object-placeholder'));
    expect(rect?.attrs).toMatchObject({ width: '144', height: '72', 'data-graphic': 'chart' });
    expect(collect(svg, 'title')[0]?.text).toBe('季度图表');
  });

  it('没传 imageHref 时也画框 —— 版面上留一个正确尺寸的空当，而不是莫名其妙的空白', () => {
    const svg = buildPage(pageWith([object()]));
    expect(collect(svg, 'image')).toHaveLength(0);
    expect(collect(svg, 'rect').some((r) => r.attrs.class?.includes('uw-object-placeholder'))).toBe(true);
  });

  it('零尺寸的对象什么都不画', () => {
    const svg = buildPage(pageWith([object({ width: 0, height: 0 })]), { imageHref: href });
    expect(collect(svg, 'image')).toHaveLength(0);
    expect(collect(svg, 'rect').some((r) => r.attrs.class?.includes('uw-object'))).toBe(false);
  });
});

describe('浮动对象', () => {
  const float = (over: Partial<PlacedFloat> = {}): PlacedFloat => ({
    runId: 'r2',
    contentIndex: 0,
    x: 1440,
    y: 2880,
    width: 1440,
    height: 1440,
    objectKind: 'drawing',
    image: { id: 'rId5', relId: 'rId5' },
    behindDoc: false,
    z: 0,
    ...over,
  });

  it('坐标相对纸左上角，不套版心的 translate', () => {
    const svg = buildPage(pageWith([], [float()]), { imageHref: href });
    const [img] = collect(svg, 'image');
    expect(img?.attrs).toMatchObject({ x: '72', y: '144', class: 'uw-float' });
  });

  it('衬于文字下方的画在正文之前，浮于上方的画在最后 —— SVG 里「层」就是先后', () => {
    const svg = buildPage(pageWith([], [float({ behindDoc: true, z: 1 }), float({ z: 2 })]), {
      imageHref: href,
    });
    const kinds = svg.children.map((c) => c.attrs.class ?? c.tag);
    const content = kinds.indexOf('uw-content');
    const floats = kinds.flatMap((k, i) => (k === 'uw-float' ? [i] : []));
    expect(floats[0]).toBeLessThan(content);
    expect(floats[1]).toBeGreaterThan(content);
  });
});
