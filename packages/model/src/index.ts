/**
 * 文档模型 + 样式级联。
 *
 * Phase 1 现状：`loadDocument()` 这条链已经通 —— OPC → 样式表 / 主题 / 设置 / 字体表 / 编号定义
 * → 正文节点树（段落 / run / 表格结构 / 分节）→ 级联完的纯数据树。
 * 编号已经消费到「每段的编号文字」（`ResolvedParaProps.numbering.label`）；
 * 未建：表格属性（Phase 4）、编号文字接进行首几何（`@uw/layout`）。
 */
export * from './cascade.ts';
export * from './font-table.ts';
export * from './load.ts';
export * from './nodes.ts';
export * from './number-format.ts';
export * from './numbering.ts';
export * from './numbering-counter.ts';
export * from './parse-body.ts';
export * from './parse-props.ts';
export * from './props.ts';
export * from './resolve-body.ts';
export * from './section.ts';
export * from './settings.ts';
export * from './styles.ts';
export * from './theme.ts';
export * from './xml-values.ts';
