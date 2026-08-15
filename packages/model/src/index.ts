/**
 * 文档模型 + 样式级联。
 *
 * Phase 1 现状：`loadDocument()` 这条链已经通 ——
 * OPC → 样式表 / 主题 → 正文节点树（段落 / run / 表格结构 / 分节）→ 级联完的纯数据树。
 * 未建：编号（`numbering.xml`）、`settings.xml`、`fontTable.xml`、表格属性（Phase 4）。
 */
export * from './cascade.ts';
export * from './load.ts';
export * from './nodes.ts';
export * from './parse-body.ts';
export * from './parse-props.ts';
export * from './props.ts';
export * from './resolve-body.ts';
export * from './section.ts';
export * from './styles.ts';
export * from './theme.ts';
export * from './xml-values.ts';
