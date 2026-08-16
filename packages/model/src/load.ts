/**
 * 从 OPC 包里装配出整份文档。
 *
 * 除主文档外，**每个部件都可能不存在** —— 那不是错误，是「这份文档没定义样式 / 没用主题 /
 * 没有编号」。所以这里一个异常都不抛，缺了就用空表（架构 §10）。会抛的只有 `OpcPackage.open`
 * 那一步的结构性错误。
 *
 * 部件一律**按关系类型找**，不按路径猜：`word/styles.xml` 是惯例不是规范，
 * 换个生成器路径就变了，而关系类型 URI 是规范定死的。
 */
import type { DiagnosticSink } from '@uw/core';
import type { OpcPackage, XmlDocument } from '@uw/ooxml';
import { RelType } from '@uw/ooxml';
import type { CascadeContext } from './cascade.ts';
import type { FieldRegion } from './fields.ts';
import { fieldHyperlinks, scanFields } from './fields.ts';
import type { FontTable } from './font-table.ts';
import { parseFontTable } from './font-table.ts';
import type { Body, ResolvedBody } from './nodes.ts';
import type { Numbering } from './numbering.ts';
import { parseNumbering } from './numbering.ts';
import { parseBody } from './parse-body.ts';
import { resolveBody } from './resolve-body.ts';
import { parseSettings } from './settings.ts';
import { parseStyles } from './styles.ts';
import { EMPTY_THEME, parseTheme } from './theme.ts';

/** 按关系类型取一个部件的 XML；没有这个关系就返回 undefined */
function partXml(pkg: OpcPackage, relType: string): XmlDocument | undefined {
  const name = pkg.partNameByRelType(relType);
  return name === undefined ? undefined : pkg.xml(name);
}

export function loadCascadeContext(pkg: OpcPackage, diagnostics: DiagnosticSink): CascadeContext {
  const styles = partXml(pkg, RelType.STYLES);
  const theme = partXml(pkg, RelType.THEME);
  const settings = partXml(pkg, RelType.SETTINGS);

  if (styles === undefined) {
    diagnostics.info('styles-missing', '文档没有 styles.xml，全部属性走 Word 内建默认值');
  }
  if (theme === undefined) {
    diagnostics.info('theme-missing', '文档没有 theme1.xml，主题字体引用会解析成空字体名');
  }

  return {
    styles: parseStyles(styles, diagnostics),
    theme: theme === undefined ? EMPTY_THEME : parseTheme(theme),
    // settings.xml 缺席不发诊断：它的每一项都有规范默认值，缺了不影响正确性
    settings: parseSettings(settings),
    // 编号级自带的缩进要参与段落级联，所以 numbering.xml 属于级联上下文的一部分
    numbering: parseNumbering(partXml(pkg, RelType.NUMBERING), diagnostics),
  };
}

/**
 * 一份 docx 解析完的样子 —— Phase 1 的产出，也是 `@uw/layout` 的输入来源。
 *
 * 分开放是因为用途不同：
 * - `body` 是**可编辑**的那棵树（直接格式），编辑期改它
 * - `resolved` 是**给布局**的那棵（级联完、纯数据），编辑后重算
 * - `cascade` 是重算所需的上下文，**不可结构化克隆**，不许过 Worker 边界
 * - `fonts` / `numbering` 是纯数据，可以跟着 `resolved` 一起过界
 */
export interface LoadedDocument {
  cascade: CascadeContext;
  body: Body;
  resolved: ResolvedBody;
  /** 字体表：本地化字体名 →`altName` 的桥，`@uw/fonts` 查字体要它 */
  fonts: FontTable;
  /**
   * 编号定义。段落上算好的编号在 `resolved` 里（`ResolvedParaProps.numbering.label`），
   * 这份原始定义留给编辑期（改编号、加一级）与回写。
   */
  numbering: Numbering;
  /**
   * 配对好的域，按 begin 的先后排。
   *
   * 单独放一份而不是挂在树上，是因为域**跨段落**（TOC 能跨几十段），挂在任何一个节点上
   * 都得再补一堆跨节点引用 —— 那正是原则 1.1 要挡的反向指针。
   * 现在的消费者只有 HYPERLINK（已铺进 `resolved` 的 run 上）；PAGE / TOC 的求值等分页。
   */
  fields: FieldRegion[];
}

export function loadDocument(pkg: OpcPackage, diagnostics: DiagnosticSink): LoadedDocument {
  const cascade = loadCascadeContext(pkg, diagnostics);
  const partName = pkg.mainDocumentPartName();
  const body = parseBody(pkg.xml(partName), diagnostics, partName);
  // 域要在级联**之前**扫：级联是按段落递归的，跨段落的配对在那儿看不见
  const fields = scanFields(body, diagnostics);
  return {
    cascade,
    body,
    resolved: resolveBody(cascade, body, { hyperlinks: fieldHyperlinks(fields) }),
    fields,
    fonts: parseFontTable(partXml(pkg, RelType.FONT_TABLE)),
    // 不重新解析一遍：同一份定义解析两次会让 numbering.xml 的诊断也报两次
    numbering: cascade.numbering,
  };
}
