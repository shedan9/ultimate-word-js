/**
 * OPC 容器 + XML → 原始 OOXML 树。
 *
 * 这一层不认识 WordprocessingML 的任何语义 —— 语义在 `@uw/model`。
 */
export * from './content-types.ts';
export * from './opc.ts';
export * from './part-names.ts';
export * from './rels.ts';
export * from './xml.ts';
export * from './zip.ts';
