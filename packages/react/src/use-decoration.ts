/**
 * `useDecoration(view, range, options)` —— 组件活着期间在视图上挂一条装饰。
 *
 * range 与 options 按**值**比（`same.ts`）：调用方每次渲染都会造新对象，按引用比会让高亮
 * 每帧闪一次。视图为 null（还没挂上）或 range 为 null 时什么都不挂；视图重挂后旧句柄已经
 * 随旧视图没了，effect 跟着 `view` 重跑就是重挂。
 */
import { useEffect, useRef } from 'react';
import type { DecorationOptions, DocRange, UwView } from 'ultimate-word';
import { sameDecoration, sameRange } from './same.ts';

/** 值没变就沿用上一次的引用，让它能当 effect 依赖 */
function useStable<T>(value: T, same: (a: T, b: T) => boolean): T {
  const ref = useRef(value);
  if (ref.current !== value && !same(ref.current, value)) ref.current = value;
  return ref.current;
}

export function useDecoration(
  view: UwView | null | undefined,
  range: DocRange | null | undefined,
  options: DecorationOptions = {},
): void {
  const stableRange = useStable(range ?? null, (a, b) =>
    a === null || b === null ? a === b : sameRange(a, b),
  );
  const stableOptions = useStable(options, sameDecoration);
  useEffect(() => {
    if (view === null || view === undefined || stableRange === null) return;
    const handle = view.decorate(stableRange, stableOptions);
    return () => handle.dispose();
  }, [view, stableRange, stableOptions]);
}
