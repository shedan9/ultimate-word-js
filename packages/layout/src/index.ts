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
 * **还没有的**：页眉页脚（部件没解析）、表格拆行（行是原子的）、TOC / SEQ 的求值。
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
export * from './items.ts';
export * from './line-height.ts';
export * from './linebreak.ts';
export * from './page.ts';
export * from './paragraph.ts';
export * from './table.ts';
export * from './table-borders.ts';
export * from './types.ts';
export * from './uncalibrated.ts';
