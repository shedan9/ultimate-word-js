/**
 * 形状的几何：预设名 + 外框 → SVG 元素。
 *
 * 单独一个文件，因为这里全是**规范里写死的公式**（ECMA-376 Part 1 §20.1.10.56 的预设几何定义、
 * §20.1.10.48 的虚线图案），与 paint.ts 里那些要对着真值调的画法不是一回事 —— 公式抄错了
 * 对照规范就能查，不必去翻样本。没有规范依据的常数（箭头多大）在 `uncalibrated.ts`。
 *
 * 外框单位是 **pt**（渲染层的坐标）。几何名与调整值在 `@uw/model` 的 `ShapeStyle` 上定义，
 * 路径在这里现算：外框是布局之后才定下来的东西，模型那一侧手上没有它。
 */
import { twipsToPt } from '@uw/core';
import type { LineObject } from '@uw/layout';
import type { RElement } from './tree.ts';
import { el, fmt } from './tree.ts';
import { LINE_END_MIN_WIDTH_PT, LINE_END_SCALE } from './uncalibrated.ts';

type ShapeStyle = NonNullable<LineObject['shape']>;
type LineEnd = NonNullable<NonNullable<ShapeStyle['stroke']>['head']>;

export interface Box {
  x: number;
  y: number;
  width: number;
  height: number;
}

/**
 * 各几何的调整值缺省（`a:avLst` 没写时，规范 `presetShapeDefinitions.xml` 里的 `avLst`）
 * 与上限（`pin 0 adj max`，超出的夹回来）。上限是短边比例的那几种都是 50000
 */
const ADJUST: Partial<Record<NonNullable<ShapeStyle['geometry']>, { def: number; max: number }>> = {
  roundRect: { def: 16667, max: 50000 },
  triangle: { def: 50000, max: 100000 },
  parallelogram: { def: 25000, max: 100000 },
  trapezoid: { def: 25000, max: 50000 },
  hexagon: { def: 25000, max: 50000 },
  octagon: { def: 29289, max: 50000 },
};

function adjust(shape: ShapeStyle): number {
  const spec = ADJUST[shape.geometry ?? 'rect'];
  if (spec === undefined) return 0;
  return Math.min(spec.max, Math.max(0, shape.adjust ?? spec.def)) / 100000;
}

/** 多边形顶点（相对外框左上角，0–1 之外的用绝对 pt 算好再传进来）→ 闭合路径 */
function polygon(box: Box, points: readonly (readonly [number, number])[]): RElement {
  const d = points.map(([px, py], i) => `${i === 0 ? 'M' : 'L'}${fmt(box.x + px)} ${fmt(box.y + py)}`);
  return el('path', { d: `${d.join(' ')} Z` });
}

/**
 * 几何 → 一个**没有样式**的 SVG 元素（填充 / 描边 / 变换由调用方补）。
 * 公式照规范的预设定义抄：`ss` 是短边，调整值多数按短边折算（三角形的顶点除外，按宽）
 */
export function shapeOutline(shape: ShapeStyle, box: Box): RElement {
  const { x, y, width: w, height: h } = box;
  const ss = Math.min(w, h);
  const a = adjust(shape);
  switch (shape.geometry ?? 'rect') {
    case 'rect':
      return el('rect', { x: fmt(x), y: fmt(y), width: fmt(w), height: fmt(h) });
    case 'roundRect': {
      const r = ss * a;
      return el('rect', { x: fmt(x), y: fmt(y), width: fmt(w), height: fmt(h), rx: fmt(r), ry: fmt(r) });
    }
    case 'ellipse':
      return el('ellipse', { cx: fmt(x + w / 2), cy: fmt(y + h / 2), rx: fmt(w / 2), ry: fmt(h / 2) });
    case 'line':
      // 规范方向永远是左上 → 右下，反过来的由翻转给（`shapeTransform`）
      return el('line', { x1: fmt(x), y1: fmt(y), x2: fmt(x + w), y2: fmt(y + h) });
    case 'triangle':
      return polygon(box, [
        [0, h],
        [w * a, 0],
        [w, h],
      ]);
    case 'rtTriangle':
      // 直角在左下
      return polygon(box, [
        [0, 0],
        [0, h],
        [w, h],
      ]);
    case 'diamond':
      return polygon(box, [
        [w / 2, 0],
        [w, h / 2],
        [w / 2, h],
        [0, h / 2],
      ]);
    case 'parallelogram': {
      // 调整值的上限其实是 100000 × w / ss，夹到不超过宽就够了
      const dx = Math.min(w, ss * a);
      return polygon(box, [
        [0, h],
        [dx, 0],
        [w, 0],
        [w - dx, h],
      ]);
    }
    case 'trapezoid': {
      // Word 2007 起的梯形是**上窄下宽**（与 2003 的 VML 梯形相反）
      const dx = Math.min(w / 2, ss * a);
      return polygon(box, [
        [0, h],
        [dx, 0],
        [w - dx, 0],
        [w, h],
      ]);
    }
    case 'hexagon': {
      // 缺省的 vf（115470）让上下两条边正好贴着外框，这里不认别的 vf
      const dx = Math.min(w / 2, ss * a);
      return polygon(box, [
        [0, h / 2],
        [dx, 0],
        [w - dx, 0],
        [w, h / 2],
        [w - dx, h],
        [dx, h],
      ]);
    }
    case 'octagon': {
      const d = ss * a;
      return polygon(box, [
        [0, d],
        [d, 0],
        [w - d, 0],
        [w, d],
        [w, h - d],
        [w - d, h],
        [d, h],
        [0, h - d],
      ]);
    }
    case 'flowChartTerminator': {
      // 规范在 21600 × 21600 的格子里画：两头是半椭圆，横半径 3475 / 21600 个宽、纵半径半个高
      const rx = (w * 3475) / 21600;
      const ry = h / 2;
      const r = `${fmt(rx)} ${fmt(ry)} 0 0 1`;
      return el('path', {
        d: `M${fmt(x + rx)} ${fmt(y)} H${fmt(x + w - rx)} A${r} ${fmt(x + w - rx)} ${fmt(y + h)} H${fmt(x + rx)} A${r} ${fmt(x + rx)} ${fmt(y)} Z`,
      });
    }
    case 'flowChartInputOutput':
      // 5 × 5 的格子：上边从 1 到 5、下边从 0 到 4 —— 斜进去的是宽的五分之一，与短边无关
      return polygon(box, [
        [0, h],
        [w / 5, 0],
        [w, 0],
        [(w * 4) / 5, h],
      ]);
  }
}

/**
 * 旋转与翻转，都绕外框中心。SVG 的变换列表**从右往左**作用在点上：点先翻转、再旋转 ——
 * 与 DrawingML 一致（翻转发生在形状自己的坐标里，旋转在最外面）
 */
export function shapeTransform(shape: ShapeStyle, box: Box): string | undefined {
  const cx = box.x + box.width / 2;
  const cy = box.y + box.height / 2;
  const parts: string[] = [];
  if (shape.rotation !== undefined && shape.rotation !== 0) {
    parts.push(`rotate(${fmt(shape.rotation)} ${fmt(cx)} ${fmt(cy)})`);
  }
  if (shape.flipH === true) parts.push(`translate(${fmt(2 * cx)} 0) scale(-1 1)`);
  if (shape.flipV === true) parts.push(`translate(0 ${fmt(2 * cy)}) scale(1 -1)`);
  return parts.length === 0 ? undefined : parts.join(' ');
}

/**
 * 预设虚线的图案，单位是**线宽**（§20.1.10.48 `ST_PresetLineDashVal` 的说明）。
 * `sys*` 那一组的空隙只有一个线宽，是「屏幕上看着像实线也能分辨」的细碎虚线
 */
const DASH_PATTERNS: Readonly<Record<string, readonly number[]>> = {
  dot: [1, 3],
  dash: [4, 3],
  lgDash: [8, 3],
  dashDot: [4, 3, 1, 3],
  lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3],
  sysDash: [3, 1],
  sysDot: [1, 1],
  sysDashDot: [3, 1, 1, 1],
  sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

/** 虚线 → `stroke-dasharray`（pt）。认不出 / 实线返回 undefined */
export function dashArray(dash: string | undefined, widthPt: number): string | undefined {
  const pattern = dash === undefined ? undefined : DASH_PATTERNS[dash];
  if (pattern === undefined) return undefined;
  return pattern.map((n) => fmt(n * widthPt)).join(' ');
}

/**
 * 直线两端的箭头。坐标是**翻转之前**的（线从外框左上画到右下），由调用方把它们与线一起
 * 放进同一个带 `shapeTransform` 的组里 —— 箭头跟着翻转走，不用自己再算一遍方向。
 *
 * 返回箭头元素，以及线要缩回到哪儿：实心三角的尖就是线的端点，线若画到尖上，平头的线端会从
 * 三角形的两条斜边旁边探出来（线越粗越明显），所以三角形 / 隐形箭头把线缩到箭头的底边
 */
export function lineEnds(
  stroke: NonNullable<ShapeStyle['stroke']>,
  box: Box,
  color: string,
): { ends: RElement[]; start: [number, number]; end: [number, number] } {
  const start: [number, number] = [box.x, box.y];
  const end: [number, number] = [box.x + box.width, box.y + box.height];
  const len = Math.hypot(box.width, box.height);
  const ends: RElement[] = [];
  if (len === 0) return { ends, start, end };
  const ux = box.width / len;
  const uy = box.height / len;
  const widthPt = Math.max(LINE_END_MIN_WIDTH_PT, twipsToPt(stroke.width));
  const head = stroke.head === undefined ? undefined : arrow(stroke.head, start, [-ux, -uy], widthPt, color);
  const tail = stroke.tail === undefined ? undefined : arrow(stroke.tail, end, [ux, uy], widthPt, color);
  if (head !== undefined) ends.push(head.el);
  if (tail !== undefined) ends.push(tail.el);
  return { ends, start: head?.back ?? start, end: tail?.back ?? end };
}

/**
 * 一个箭头。`tip` 是线的端点，`u` 是从线内指向端点的单位向量；`w` 跨在线两侧、`len` 沿着线。
 * `back` 是线该缩回到的点（只有实心的三角与隐形箭头要缩）
 */
function arrow(
  end: LineEnd,
  tip: [number, number],
  u: [number, number],
  widthPt: number,
  color: string,
): { el: RElement; back?: [number, number] } {
  const L = LINE_END_SCALE[end.len] * widthPt;
  const W = LINE_END_SCALE[end.w] * widthPt;
  const [tx, ty] = tip;
  const [ux, uy] = u;
  // 法向量：u 转 90°
  const nx = -uy;
  const ny = ux;
  const at = (along: number, across: number): [number, number] => [
    tx + ux * along + nx * across,
    ty + uy * along + ny * across,
  ];
  const path = (pts: [number, number][], close: boolean): string =>
    pts.map(([px, py], i) => `${i === 0 ? 'M' : 'L'}${fmt(px)} ${fmt(py)}`).join(' ') + (close ? ' Z' : '');
  const solid = { fill: color, stroke: 'none' };
  switch (end.type) {
    case 'triangle':
      return {
        el: el('path', { d: path([at(0, 0), at(-L, W / 2), at(-L, -W / 2)], true), ...solid }),
        back: at(-L, 0),
      };
    case 'stealth':
      // 后缘凹进去四分之一个长度
      return {
        el: el('path', {
          d: path([at(0, 0), at(-L, W / 2), at(-L * 0.75, 0), at(-L, -W / 2)], true),
          ...solid,
        }),
        back: at(-L * 0.75, 0),
      };
    case 'arrow':
      // 开口箭头是两笔线，与线同粗
      return {
        el: el('path', {
          d: path([at(-L, W / 2), at(0, 0), at(-L, -W / 2)], false),
          fill: 'none',
          stroke: color,
          'stroke-width': fmt(widthPt),
          'stroke-linejoin': 'miter',
        }),
      };
    case 'diamond':
      return {
        el: el('path', {
          d: path([at(L / 2, 0), at(0, W / 2), at(-L / 2, 0), at(0, -W / 2)], true),
          ...solid,
        }),
      };
    case 'oval': {
      const angle = (Math.atan2(uy, ux) * 180) / Math.PI;
      return {
        el: el('ellipse', {
          cx: fmt(tx),
          cy: fmt(ty),
          rx: fmt(L / 2),
          ry: fmt(W / 2),
          transform: `rotate(${fmt(angle)} ${fmt(tx)} ${fmt(ty)})`,
          ...solid,
        }),
      };
    }
  }
}
