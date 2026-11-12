import type { UwView } from 'ultimate-word';

// 父组件的被动 effect 可能先于 onViewChange 引起的新渲染执行，仍握着旧视图。
// 只标记本封装销毁的句柄，不把尚未收到新状态误当成存活；弱引用不延长视图寿命。
export const disposedViews = new WeakSet<UwView>();
