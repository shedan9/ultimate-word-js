/**
 * 按值比较模型位置 / range / 装饰选项。React 侧每次渲染都会造新对象，按引用比会让每个
 * overlay 每帧重挂一次（批注气泡里的输入框会失焦）—— 所以「变没变」一律按值判。
 */
import type { DecorationOptions, DocPosition, DocRange } from 'ultimate-word';

export function samePosition(a: DocPosition, b: DocPosition): boolean {
  return a.nodeId === b.nodeId && a.contentIndex === b.contentIndex && a.offset === b.offset;
}

export function sameRange(a: DocRange, b: DocRange): boolean {
  return samePosition(a.start, b.start) && samePosition(a.end, b.end);
}

export function sameOffset(
  a: { x: number; y: number } | undefined,
  b: { x: number; y: number } | undefined,
): boolean {
  if (a === undefined || b === undefined) return a === b;
  return a.x === b.x && a.y === b.y;
}

export function sameStyle(
  a: Readonly<Record<string, string>> | undefined,
  b: Readonly<Record<string, string>> | undefined,
): boolean {
  if (a === undefined || b === undefined) return a === b;
  const ka = Object.keys(a);
  const kb = Object.keys(b);
  return ka.length === kb.length && ka.every((k) => a[k] === b[k]);
}

export function sameDecoration(a: DecorationOptions, b: DecorationOptions): boolean {
  return a.className === b.className && a.layer === b.layer && sameStyle(a.style, b.style);
}
