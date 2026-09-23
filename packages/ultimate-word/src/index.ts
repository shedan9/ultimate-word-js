/**
 * `ultimate-word` —— 门面包。
 *
 * ```ts
 * import { UltimateWord } from 'ultimate-word';
 * const doc = await UltimateWord.load(arrayBuffer);
 * const view = doc.mount('#container');
 * ```
 *
 * 两行。它不实现任何东西，只把 `@uw/*` 六个包收成 api.md 的形状；`@uw/*` 视为内部实现，
 * 只有这个包遵循 semver（api.md §16）。四个概念：`UwDocument`（内容 + 布局，无关屏幕）、
 * `UwView`（一次屏幕呈现）、`DocPosition` / `DocRange`（指向内容的坐标，重排后依然有效）、
 * `Disposable`（凡是挂上去的都能摘下来）。
 *
 * 类型名带 `Uw` 前缀是因为 `Document` / `View` 与 DOM 的全局类型撞名 —— 在一个满是
 * `HTMLElement` 的调用方文件里，裸的 `Document` 十有八九会被当成 DOM 那个。
 */
import { createFontsApi, createRegistry } from './fonts.ts';
import type { LoadOptions, LoadSource } from './load.ts';
import { load } from './load.ts';

export type { Diagnostic, DiagnosticSeverity } from '@uw/core';
export { UwError, UwErrorCode } from '@uw/core';
export type { FontRegistry, FontStatus, MetricsPack } from '@uw/fonts';
export type {
  DocPosition,
  DocRange,
  NodeId,
  TextChangeSet,
  TextTransaction,
  TextTransactionOptions,
} from '@uw/model';
export type { ClientPoint, ClientRect } from '@uw/view';
export type {
  DecorationHandle,
  DecorationOptions,
  OverlayHandle,
  OverlayOptions,
  OverlayPlacement,
  ScrollOptions,
  ScrollTarget,
} from '@uw/view/dom';
export type { DocNode, DocumentEvents, FindOptions } from './document.ts';
export { DOCX_MIME, UwDocument } from './document.ts';
export type { FontsApi } from './fonts.ts';
export type { LoadOptions, LoadSource } from './load.ts';
export type { Disposable, ElementHit, UwView, ViewEvents, ViewOptions, ZoomSpec } from './view.ts';

const registry = createRegistry();

export const UltimateWord = {
  /** 全局字体注册表，随库 17 款度量包已经在里面 */
  fonts: createFontsApi(registry),
  /** 唯一的异步入口。之后所有查询都是同步的 */
  load: (source: LoadSource, options?: LoadOptions) => load(source, registry, options),
};
