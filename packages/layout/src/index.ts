/**
 * 布局引擎。
 *
 * 现状：流水线里**行盒之前**的部分已经完整 —— 分桶 / 度量（`@uw/fonts` 提供）→ 断行
 * （UAX#14 骨架 + 中文禁则 + 标点挤压 + 中西文自动间距）→ 行内水平几何（缩进、
 * 对齐、制表位、列表编号）→ 行高总量（行距规则 + 行网格吸附）。
 *
 * 行盒（`baseline`）与**分页**（`page.ts`）也已经做完：`layoutDocument()` 把段落与表格
 * 摞进一页页的版心，每一行都拿到了页号与页内的 y。gongwen-01 的 18 行基线 y 与 Word
 * 真值最大差 0.06pt（L3 的判据是 0.5pt）。
 *
 * **域求值**（`fields.ts`）也接上了：`layoutDocumentWithFields()` 把「排版 → 算页码 → 再排版」
 * 迭代到自洽，PAGE / NUMPAGES / SECTIONPAGES 已经是算出来的而不是文件里存的旧值。
 *
 * **页眉页脚**（`header-footer.ts`）同样做完了：每一页选出该用的那一份、排出高度、
 * 按实测的三条几何规则摆在纸上，并**反过来把版心挤窄**（页边距是最小值不是固定值）。
 * 页脚里的 `{ PAGE }` 因此一趟就是准的 —— 页码在开页那一刻就定了。
 *
 * **表格拆行**（`table-split.ts`）也做完了：一行放不下时从行间切开，本页一片、下一页一片
 * （`w:cantSplit` 与表头行除外）。切出来的是两份各自自洽的 `RowLayout`，
 * 渲染层不需要裁剪窗口。
 *
 * **布局索引**（`layout-index.ts`）是 Phase 6 的地基：`buildLayoutIndex()` 把整份布局摊平成
 * 一张行表，答「点了哪个字」「这个 range 该画在哪几个矩形上」「光标在哪」「谁先谁后」。
 * 它**在消费侧现建**（带方法的对象过不了结构化克隆），屏幕坐标那一跳不在这里。
 *
 * **还没有的**：TOC / SEQ 的求值、脚注尾注。
 *
 * 约束（从第一天就守住，否则后面搬不动）：
 * - 本包**不得** import 任何 DOM API，度量能力靠注入的 `TextMeasurer` 传进来 ——
 *   这样才能在 Worker / Node 里跑，也是未来把 linebreak + measure 换成 Rust→WASM 的前提
 * - 输出必须可结构化克隆：没有类实例、闭包、DOM 引用、反向指针
 * - 单位只有 twips，px 只在渲染出口出现
 */
export const PACKAGE_NAME = '@uw/layout';

export * from './break-class.ts';
export * from './fields.ts';
export * from './header-footer.ts';
export * from './items.ts';
export * from './layout-index.ts';
export * from './line-height.ts';
export * from './linebreak.ts';
export * from './page.ts';
export * from './paragraph.ts';
export * from './table.ts';
export * from './table-borders.ts';
export * from './table-split.ts';
export * from './types.ts';
export * from './uncalibrated.ts';
