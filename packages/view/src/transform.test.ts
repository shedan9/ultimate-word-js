import { twipsToPx } from '@uw/core';
import { describe, expect, it } from 'vitest';
import type { PageViewport } from './transform.ts';
import { createViewTransform } from './transform.ts';

function page(page = 0, zoom = 1, top = 100): PageViewport {
  return {
    page,
    width: 6000,
    height: 9000,
    matrix: { a: twipsToPx(1, zoom), b: 0, c: 0, d: twipsToPx(1, zoom), e: 200, f: top },
  };
}

describe('屏幕与布局坐标', () => {
  it.each([0.5, 1, 1.75, 2])('缩放 %s 下往返相同字缝，保留小数精度', (zoom) => {
    const t = createViewTransform([page(0, zoom)]);
    const client = t.toClient({ page: 0, x: 1234.5, y: 4567.25 });
    expect(client?.clientX).toBeCloseTo(200 + twipsToPx(1234.5, zoom));
    if (client === undefined) throw new Error('页面未映射');
    const back = t.toLayout(client);
    expect(back?.x).toBeCloseTo(1234.5);
    expect(back?.y).toBeCloseTo(4567.25);
  });

  it('多页使用物理页序；纸外与页间空隙不命中', () => {
    const t = createViewTransform([page(3), page(8, 1, 724)]);
    expect(t.toLayout({ clientX: 210, clientY: 734 })).toEqual({ page: 8, x: 150, y: 150 });
    expect(t.toLayout({ clientX: 210, clientY: 712 })).toBeUndefined();
    expect(t.toLayout({ clientX: 199, clientY: 110 })).toBeUndefined();
    expect(t.toClient({ page: 1, x: 0, y: 0 })).toBeUndefined();
  });

  it('横向页采用自身几何，重叠时后绘制的页优先', () => {
    const landscape = { ...page(1), width: 9000, height: 6000 };
    const t = createViewTransform([page(), landscape]);
    expect(t.toLayout({ clientX: 700, clientY: 200 })?.page).toBe(1);
    expect(t.toLayout({ clientX: 300, clientY: 200 })?.page).toBe(1);
  });

  it('旋转后按实际纸面命中，矩形取四角包围盒', () => {
    const t = createViewTransform([
      {
        page: 0,
        width: 100,
        height: 200,
        matrix: { a: 0, b: 2, c: -2, d: 0, e: 500, f: 100 },
      },
    ]);
    expect(t.toLayout({ clientX: 440, clientY: 140 })).toEqual({ page: 0, x: 20, y: 30 });
    expect(t.rectToClient({ page: 0, x: 20, y: 30, width: 10, height: 40 })).toEqual({
      x: 360,
      y: 140,
      width: 80,
      height: 20,
    });
    expect(t.toLayout({ clientX: 510, clientY: 140 })).toBeUndefined();
  });

  it('倾斜后的包围盒空白不属于纸面', () => {
    const t = createViewTransform([
      {
        page: 0,
        width: 100,
        height: 100,
        matrix: { a: 1, b: 0, c: 1, d: 1, e: 0, f: 0 },
      },
    ]);
    expect(t.toLayout({ clientX: 10, clientY: 90 })).toBeUndefined();
    expect(t.toLayout({ clientX: 100, clientY: 90 })).toEqual({ page: 0, x: 10, y: 90 });
  });

  it('不可见尺寸、不可逆矩阵和非法输入不产生伪命中', () => {
    const p = page();
    for (const bad of [
      { ...p, width: 0 },
      { ...p, matrix: { ...p.matrix, a: 0 } },
      { ...p, matrix: { ...p.matrix, e: Number.NaN } },
    ]) {
      expect(createViewTransform([bad]).toClient({ page: 0, x: 0, y: 0 })).toBeUndefined();
    }
    expect(createViewTransform([p]).toLayout({ clientX: Infinity, clientY: 100 })).toBeUndefined();
  });

  it('快照不受宿主随后修改测量对象影响', () => {
    const p = page();
    const t = createViewTransform([p]);
    p.matrix.e = 900;
    expect(t.toClient({ page: 0, x: 0, y: 0 })?.clientX).toBe(200);
  });
});
