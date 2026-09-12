import type { LayoutIndex, LayoutRect } from '@uw/layout';
import type { DocPosition } from '@uw/model';
import type { ClientRect } from './transform.ts';

export type OverlayPlacement = 'right-of-line' | 'above' | 'below' | 'inline';

/** 优先采用 caretRect 的断行归属，不能只按 runId 找它第一次出现的行。 */
export function anchorRect(
  index: LayoutIndex,
  position: DocPosition,
  placement: OverlayPlacement,
): LayoutRect | undefined {
  const caret = index.caretRect(position);
  if (caret === undefined || placement !== 'right-of-line') return caret;
  const line = index.lines.find(
    (entry) =>
      entry.page === caret.page &&
      entry.top === caret.y &&
      entry.line.fragments.some(
        (fragment) =>
          fragment.runId === position.nodeId &&
          fragment.contentIndex === position.contentIndex &&
          fragment.offset >= 0 &&
          position.offset >= fragment.offset &&
          position.offset <= fragment.offset + fragment.text.length,
      ),
  );
  if (line === undefined) return undefined;
  return { ...caret, x: line.originX + line.line.x, width: line.line.width };
}

/** 与页面 SVG 默认 xMidYMid meet 一致；坐标在壳内，祖先滚动 / CSS 变换由浏览器继承。 */
export function pageRect(
  rect: LayoutRect,
  pageWidth: number,
  pageHeight: number,
  width: number,
  height: number,
): ClientRect {
  const scale = Math.min(width / pageWidth, height / pageHeight);
  return {
    x: (width - pageWidth * scale) / 2 + rect.x * scale,
    y: (height - pageHeight * scale) / 2 + rect.y * scale,
    width: rect.width * scale,
    height: rect.height * scale,
  };
}

/** offset 是页面壳内的 CSS px；inline 对齐字缝，其余位置围绕所属行或字缝。 */
export function placeOverlay(
  rect: ClientRect,
  size: { width: number; height: number },
  placement: OverlayPlacement,
  offset: { x: number; y: number },
): { x: number; y: number } {
  let x = rect.x;
  let y = rect.y;
  if (placement === 'right-of-line') x += rect.width;
  if (placement === 'above') y -= size.height;
  if (placement === 'below') y += rect.height;
  return { x: x + offset.x, y: y + offset.y };
}
