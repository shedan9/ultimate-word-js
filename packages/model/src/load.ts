/**
 * 从 OPC 包里装配出级联所需的上下文。
 *
 * 样式表和主题**都可能不存在** —— 那不是错误，是「这份文档没定义样式 / 没用主题」。
 * 所以这里一个异常都不抛，缺了就用空表（架构 §10）。会抛的只有 `OpcPackage.open`
 * 那一步的结构性错误。
 */
import type { DiagnosticSink } from '@uw/core';
import type { OpcPackage } from '@uw/ooxml';
import { RelType } from '@uw/ooxml';
import type { CascadeContext } from './cascade.ts';
import type { Body, ResolvedBody } from './nodes.ts';
import { parseBody } from './parse-body.ts';
import { resolveBody } from './resolve-body.ts';
import { parseStyles } from './styles.ts';
import { EMPTY_THEME, parseTheme } from './theme.ts';

export function loadCascadeContext(pkg: OpcPackage, diagnostics: DiagnosticSink): CascadeContext {
  const stylesPart = pkg.partNameByRelType(RelType.STYLES);
  const themePart = pkg.partNameByRelType(RelType.THEME);

  if (stylesPart === undefined) {
    diagnostics.info('styles-missing', '文档没有 styles.xml，全部属性走 Word 内建默认值');
  }
  if (themePart === undefined) {
    diagnostics.info('theme-missing', '文档没有 theme1.xml，主题字体引用会解析成空字体名');
  }

  return {
    styles: parseStyles(stylesPart === undefined ? undefined : pkg.xml(stylesPart), diagnostics),
    theme: themePart === undefined ? EMPTY_THEME : parseTheme(pkg.xml(themePart)),
  };
}

/**
 * 一份 docx 解析完的样子 —— Phase 1 的产出，也是 `@uw/layout` 的输入来源。
 *
 * 三样东西都留着，因为用途不同：
 * - `body` 是**可编辑**的那棵树（直接格式），编辑期改它
 * - `resolved` 是**给布局**的那棵（级联完、纯数据），编辑后重算
 * - `cascade` 是重算所需的上下文，**不可结构化克隆**，不许过 Worker 边界
 */
export interface LoadedDocument {
  cascade: CascadeContext;
  body: Body;
  resolved: ResolvedBody;
}

export function loadDocument(pkg: OpcPackage, diagnostics: DiagnosticSink): LoadedDocument {
  const cascade = loadCascadeContext(pkg, diagnostics);
  const partName = pkg.mainDocumentPartName();
  const body = parseBody(pkg.xml(partName), diagnostics, partName);
  return { cascade, body, resolved: resolveBody(cascade, body) };
}
