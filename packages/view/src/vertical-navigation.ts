import type { IndexedLine, LayoutIndex } from '@uw/layout';
import type { DocPosition } from '@uw/model';

export interface VerticalMove {
  position: DocPosition;
  /** 纸坐标 twips；经过短行、缩放与滚动仍沿用第一次上下移动的横向位置。 */
  x: number;
}

const graphemes = new Intl.Segmenter('zh', { granularity: 'grapheme' });

/** 从排版行导航，不在屏幕上猜行高；页边距、空白补页和段间距都不是一行。 */
export function moveVertically(
  index: LayoutIndex,
  from: DocPosition,
  direction: 'up' | 'down',
  preferredX?: number,
): VerticalMove | undefined {
  const caret = index.caretRect(from);
  if (!caret) return undefined;
  const x = preferredX ?? caret.x;
  const sign = direction === 'down' ? 1 : -1;
  const candidates = index.lines
    .filter(
      (line) =>
        !line.frame &&
        !line.repeated &&
        (line.page === caret.page ? sign * (line.top - caret.y) > 0 : sign * (line.page - caret.page) > 0),
    )
    .sort((a, b) => sign * (a.page - b.page || a.top - b.top));
  let best: { position: DocPosition; page: number; top: number; distance: number } | undefined;
  for (const line of candidates) {
    if (best && (line.page !== best.page || line.top !== best.top)) break;
    for (const position of positionsIn(line)) {
      const rect = index.caretRect(position);
      // 软换行的共享字缝归下一行；不能选一个反查后仍在原行的位置，否则向上键会卡住。
      if (!rect || rect.page !== line.page || rect.y !== line.top) continue;
      const distance = Math.abs(rect.x - x);
      if (!best || distance < best.distance) best = { position, page: line.page, top: line.top, distance };
    }
  }
  return best && { position: best.position, x };
}

/** 同行片段先拼接再切字素，字号 / 样式拆片不能产生组合音标或 emoji 内部的落点。 */
function positionsIn(line: IndexedLine): DocPosition[] {
  if (line.line.emptyPosition) return [line.line.emptyPosition];
  const positions: DocPosition[] = (line.line.caretAnchors ?? []).map((anchor) => anchor.position);
  let text = '';
  const spans: { start: number; end: number; position: DocPosition }[] = [];
  for (const fragment of line.line.fragments) {
    if (fragment.offset < 0) {
      text += '\ufffc';
      continue;
    }
    const start = text.length;
    text += fragment.text;
    spans.push({
      start,
      end: text.length,
      position: { nodeId: fragment.runId, contentIndex: fragment.contentIndex, offset: fragment.offset },
    });
  }
  const boundaries = new Set<number>();
  for (const segment of graphemes.segment(text)) {
    boundaries.add(segment.index);
    boundaries.add(segment.index + segment.segment.length);
  }
  for (const span of spans) {
    for (const offset of boundaries) {
      if (offset >= span.start && offset <= span.end)
        positions.push({ ...span.position, offset: span.position.offset + offset - span.start });
    }
  }
  return positions;
}
