/**
 * `@uw/react` —— `ultimate-word` 的 React 包装（api.md §14）。
 *
 * 三个导出：`<UltimateWordView>`（`doc.mount()` 的声明式形态，`overlays` 按 key 调和）、
 * `useDocument`（`UltimateWord.load()` 的 hook 形态）、`useDecoration`（组件活着期间挂一条装饰）。
 * 它**不另起 API**：能做的事与 `ultimate-word` 一样多，拿 `ref` 就是原来那个 `UwView`。
 */
export type { OverlayKey, OverlaySpec } from './overlays.ts';
export { useDecoration } from './use-decoration.ts';
export type { DocumentState } from './use-document.ts';
export { useDocument } from './use-document.ts';
export type { ReactOverlaySpec, UltimateWordViewProps } from './view.tsx';
export { UltimateWordView } from './view.tsx';
