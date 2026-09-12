import type { Twips } from '@uw/core';
import type { LayoutPoint, LayoutRect } from '@uw/layout';

export interface ClientPoint {
  clientX: number;
  clientY: number;
}

/** 与 DOMRect 同一坐标系的纯数据；无需浏览器也能测试。 */
export interface ClientRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** x' = a*x + c*y + e，y' = b*x + d*y + f；输入 twips，输出 CSS px。 */
export interface ClientMatrix {
  a: number;
  b: number;
  c: number;
  d: number;
  e: number;
  f: number;
}

export interface PageViewport {
  /** 物理页序，不是可能重新起算的显示页码。 */
  page: number;
  width: Twips;
  height: Twips;
  matrix: ClientMatrix;
}

export interface ViewTransform {
  /** 页间空隙与纸外不吸附到正文；纸内的最近字缝交给 LayoutIndex。 */
  toLayout(point: ClientPoint): LayoutPoint | undefined;
  toClient(point: LayoutPoint): ClientPoint | undefined;
  /** 旋转时返回四个角的轴对齐包围盒，可直接用于 DOM overlay。 */
  rectToClient(rect: LayoutRect): ClientRect | undefined;
}

function project(m: ClientMatrix, x: number, y: number): ClientPoint {
  return { clientX: m.a * x + m.c * y + m.e, clientY: m.b * x + m.d * y + m.f };
}

/**
 * 一次测量的快照。调用方每次查询重新测量，不能跨滚动 / 缩放缓存这个对象。
 * 保留完整仿射变换：用 bounding box 反推比例会把旋转后的空白三角形也当成纸。
 */
export function createViewTransform(viewports: readonly PageViewport[]): ViewTransform {
  const pages = new Map<number, PageViewport>();
  for (const viewport of viewports) {
    const { width, height, matrix: m } = viewport;
    const det = m.a * m.d - m.b * m.c;
    if (
      ![width, height, m.a, m.b, m.c, m.d, m.e, m.f, det].every(Number.isFinite) ||
      width <= 0 ||
      height <= 0 ||
      det === 0
    )
      continue;
    pages.set(viewport.page, { ...viewport, matrix: { ...m } });
  }
  const backToFront = [...pages.values()].reverse();
  return {
    toLayout(point) {
      if (![point.clientX, point.clientY].every(Number.isFinite)) return undefined;
      for (const { page, width, height, matrix: m } of backToFront) {
        const det = m.a * m.d - m.b * m.c;
        const dx = point.clientX - m.e;
        const dy = point.clientY - m.f;
        const x = (m.d * dx - m.c * dy) / det;
        const y = (m.a * dy - m.b * dx) / det;
        if (x >= 0 && x <= width && y >= 0 && y <= height) return { page, x, y };
      }
      return undefined;
    },
    toClient(point) {
      const page = pages.get(point.page);
      if (page === undefined || ![point.x, point.y].every(Number.isFinite)) return undefined;
      return project(page.matrix, point.x, point.y);
    },
    rectToClient(rect) {
      const page = pages.get(rect.page);
      if (page === undefined || ![rect.x, rect.y, rect.width, rect.height].every(Number.isFinite)) {
        return undefined;
      }
      const corners = [
        project(page.matrix, rect.x, rect.y),
        project(page.matrix, rect.x + rect.width, rect.y),
        project(page.matrix, rect.x, rect.y + rect.height),
        project(page.matrix, rect.x + rect.width, rect.y + rect.height),
      ];
      const x = Math.min(...corners.map((p) => p.clientX));
      const y = Math.min(...corners.map((p) => p.clientY));
      return {
        x,
        y,
        width: Math.max(...corners.map((p) => p.clientX)) - x,
        height: Math.max(...corners.map((p) => p.clientY)) - y,
      };
    },
  };
}
