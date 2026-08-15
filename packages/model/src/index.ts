/**
 * 文档模型 + 样式级联。
 *
 * Phase 1 现状：样式级联这条链已经通（docDefaults → 样式链 → 直接格式 → Resolved*），
 * 正文节点树（段落 / run / 表格）与编号尚未建。
 */
export * from './cascade.ts';
export * from './load.ts';
export * from './parse-props.ts';
export * from './props.ts';
export * from './styles.ts';
export * from './theme.ts';
export * from './xml-values.ts';
