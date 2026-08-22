/**
 * 绝对定位 DOM 渲染器 —— 流水线的出口。
 *
 * 分成两个入口，界线是**碰不碰浏览器**：
 *
 * - `@uw/render-dom`（本文件）：`DocumentLayout` → **纯数据元素树** → 标记文本。
 *   跑在任何地方 —— 单测在纯 Node 里跑，`apps/fidelity` 的 `preview` 拿它落盘成 HTML，
 *   将来 Worker 里预生成也走这条路。**这条路上一个 DOM API 都不碰**
 * - `@uw/render-dom/dom`：元素树 → 真 DOM。只有它需要 `document`，且那个 `Document`
 *   是注入的
 *
 * 分成两个入口不是洁癖：`@uw/fidelity` 是个 Node 工具，workspace 包的 `exports` 又直接指向
 * `src/*.ts`（开发时不 build），于是**消费方的 tsc 会把入口的整个 import 图都检一遍** ——
 * 主入口只要牵进 `dom.ts`，fidelity 就被迫在自己的 tsconfig 里打开 `lib: ["DOM"]`。
 * 与 `@uw/fonts` 的主入口刻意不依赖 fontkit 是同一条理由：**别让不需要的那一半传染出去**。

反过来，这也是「主入口不碰 DOM」这条**唯一的自动执行者**：本包自己的 tsconfig 为了
`dom.ts` 打开了 `lib: ["DOM"]`，所以 paint.ts 里误用 `document` 在本包内编不出错来 ——
是 `@uw/fidelity` 的 typecheck（不带 DOM lib，且 import 的是主入口）替我们把这条守住的。
改 fidelity 的 tsconfig 去「顺手加上 DOM lib」等于拆掉这道闸。
 *
 * 一页一个 `<svg>`，viewBox 的单位是 **pt**（与 `fixtures/*.truth.json` 同一套坐标，
 * 肉眼比对不用换算），逐字 x 走 `<text x="x1 x2 …">`。缩放只改 `<svg>` 的 width /
 * height，布局结果一个字节不动 —— 架构 §4.1 的「缩放永不触发重排」在这里落地。
 *
 * 没画的东西集中写在 `paint.ts` 的文件头（图片、run 级高亮、可选文本层）。
 */
export const PACKAGE_NAME = '@uw/render-dom';

export * from './font-stack.ts';
export * from './paint.ts';
export * from './tree.ts';
export * from './uncalibrated.ts';
