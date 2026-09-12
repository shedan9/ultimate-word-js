import type { LayoutIndex, LayoutRect } from '@uw/layout';
import type { DocPosition, DocRange } from '@uw/model';

/** `{ page }` 是物理页序（0 起，与 `PageLayout.index` / 壳上的 `data-page` 同一套），不是显示页码。 */
export type ScrollTarget = DocPosition | DocRange | { page: number };

export interface ScrollOptions {
  /** 目标行贴视口的哪一边，默认 `start`。 */
  align?: 'start' | 'center' | 'end';
  behavior?: 'auto' | 'smooth';
}

/**
 * 滚动目标在布局空间里的矩形：位置取光标矩形，range 取**首行**的矩形 ——
 * 一个跨三页的选区滚到它的开头才是用户要的，滚到包围盒中心会落在空白页上。
 * 排不出来的位置（空 run、隐藏 run）没有矩形，返回 `undefined` 让调用方答 false，
 * 而不是退到页首 —— 「滚到了别处」比「没滚」更难察觉。
 */
export function scrollTargetRect(index: LayoutIndex, target: DocPosition | DocRange): LayoutRect | undefined {
  if ('start' in target) return index.rectsOf(target)[0] ?? index.caretRect(target.start);
  return index.caretRect(target);
}
