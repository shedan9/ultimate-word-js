import type { DocumentLayout, IndexedLine, LayoutIndex } from '@uw/layout';
import { buildLayoutIndex } from '@uw/layout';
import type { DocPosition, DocRange } from '@uw/model';
import type { ClientPoint, ClientRect, PageViewport } from './transform.ts';
import { createViewTransform } from './transform.ts';

export interface ReadonlyView {
  /** 当前布局索引；DOM 装饰复用它，不能另建一份文档索引。 */
  readonly index: LayoutIndex;
  /** 文档序的布局行；文字层复用索引，避免另建一份全量索引。 */
  readonly lines: readonly IndexedLine[];
  locate(point: ClientPoint): DocPosition | null;
  rectsOf(range: DocRange): ClientRect[];
  caretRect(position: DocPosition): ClientRect | null;
  /** 只有重新排版后才换索引；滚动和缩放不走这里。 */
  update(layout: DocumentLayout): void;
}

/**
 * 把两个转换接起来。页面测量由宿主注入，同一布局可以挂到不同缩放的多个视图。
 * 一次 range 查询只测量一次，避免按片段交错读 DOM 与写样式造成反复同步布局。
 */
export function createReadonlyView(
  layout: DocumentLayout,
  measurePages: () => readonly PageViewport[],
): ReadonlyView {
  let index = buildLayoutIndex(layout);
  return {
    get index() {
      return index;
    },
    get lines() {
      return index.lines;
    },
    locate(point) {
      const at = createViewTransform(measurePages()).toLayout(point);
      return at === undefined ? null : (index.positionAt(at) ?? null);
    },
    rectsOf(range) {
      const transform = createViewTransform(measurePages());
      return index.rectsOf(range).flatMap((rect) => {
        const client = transform.rectToClient(rect);
        return client === undefined ? [] : [client];
      });
    },
    caretRect(position) {
      const rect = index.caretRect(position);
      if (rect === undefined) return null;
      return createViewTransform(measurePages()).rectToClient(rect) ?? null;
    },
    update(next) {
      index = buildLayoutIndex(next);
    },
  };
}
