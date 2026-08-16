/**
 * 文档模型 + 样式级联。
 *
 * 现状：`loadDocument()` 这条链已经通 —— OPC → 样式表 / 主题 / 设置 / 字体表 / 编号定义
 * → 正文节点树（段落 / run / 表格 / 分节）→ 级联完的纯数据树。
 * 编号已经消费到「每段的编号文字」（`ResolvedParaProps.numbering.label`）；
 * 表格的属性与级联（含条件格式）也已完成，**列宽算法在 `@uw/layout`**，还没写。
 */
export * from './cascade.ts';
export * from './cascade-table.ts';
export * from './font-table.ts';
export * from './load.ts';
export * from './nodes.ts';
export * from './number-format.ts';
export * from './numbering.ts';
export * from './numbering-counter.ts';
export * from './parse-body.ts';
export * from './parse-props.ts';
export * from './parse-table-props.ts';
export * from './props.ts';
export * from './resolve-body.ts';
export * from './section.ts';
export * from './settings.ts';
export * from './styles.ts';
export * from './table-props.ts';
export * from './theme.ts';
export * from './xml-values.ts';
