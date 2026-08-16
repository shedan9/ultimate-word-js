/**
 * 布局引擎。
 *
 * 现状：流水线里**行盒之前**的部分已经完整 —— 分桶 / 度量（`@uw/fonts` 提供）→ 断行
 * （UAX#14 骨架 + 中文禁则 + 标点挤压 + 中西文自动间距）→ 行内水平几何（缩进、
 * 对齐、制表位、列表编号）→ 行高总量（行距规则 + 行网格吸附）。
 *
 * **没有的**：基线、y、页。东亚行高里那 30% 额外行距在基线上下如何分配还没标定
 * （要 Word COM 做一次「首行基线到版心顶」的穿刺），行盒装配与分页因此全部停工。
 * 补完之后 `LineLayout` 加 `baseline`、把行摞起来即可，现有字段不用动。
 *
 * 约束（从第一天就守住，否则后面搬不动）：
 * - 本包**不得** import 任何 DOM API，度量能力靠注入的 `TextMeasurer` 传进来 ——
 *   这样才能在 Worker / Node 里跑，也是未来把 linebreak + measure 换成 Rust→WASM 的前提
 * - 输出必须可结构化克隆：没有类实例、闭包、DOM 引用、反向指针
 * - 单位只有 twips，px 只在渲染出口出现
 */
export const PACKAGE_NAME = '@uw/layout';

export * from './break-class.ts';
export * from './items.ts';
export * from './line-height.ts';
export * from './linebreak.ts';
export * from './paragraph.ts';
export * from './table.ts';
export * from './types.ts';
export * from './uncalibrated.ts';
