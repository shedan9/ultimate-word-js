/**
 * 模型 → docx（Phase 8）。补丁式回写：没改的部件逐字节照搬，改过的正文以原文为底，
 * 只重写变了的那部分 —— 模型不认识的 XML 因此原样留在文件里（架构原则 1.4）。
 */
export * from './body-writer.ts';
export * from './docx.ts';
export * from './numbering-writer.ts';
export * from './props-writer.ts';
