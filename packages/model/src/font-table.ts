/**
 * `fontTable.xml` → 字体表。
 *
 * **为什么这不是可选项**：中文版 Word 把字体写成本地化名字 ——「等线」「黑体」「仿宋」。
 * 而字体文件在磁盘上、在 CSS 里、在 fontkit 解出来的 name 表里，叫的是 `DengXian`、`SimHei`。
 * `w:altName` 就是这两个名字之间唯一的桥。不读它，非中文系统上每一款中文字体都查不到，
 * 度量退化成回退字体 —— 而这个引擎的判据是「与真值差多少 pt」，字体一错全盘皆错。
 *
 * 其余字段（panose / charset / family / pitch）是**字体缺失时**挑替代品的依据：
 * panose 描述字形骨架，pitch 分等宽与比例，charset 说这款字体覆盖哪个代码页。
 * Phase 2 的字体回退会来查它们，现在先如实收着。
 */
import type { XmlDocument, XmlElement } from '@uw/ooxml';
import { attr, child, children } from '@uw/ooxml';
import { attrOf, enumVal } from './xml-values.ts';

const FAMILIES = ['auto', 'decorative', 'modern', 'roman', 'script', 'swiss'] as const;
const PITCHES = ['default', 'fixed', 'variable'] as const;

export type FontFamilyKind = (typeof FAMILIES)[number];
export type FontPitch = (typeof PITCHES)[number];

/** 内嵌字体：`w:embedRegular` 等，指向包里的一个字体部件 */
export interface EmbeddedFont {
  style: 'regular' | 'bold' | 'italic' | 'boldItalic';
  relId: string;
  /** `w:fontKey`：Word 的字体混淆密钥（GUID）。真要用内嵌字体时得拿它解混淆 */
  fontKey?: string;
  /** `w:subsetted`：是不是子集化过的字体（缺字符很正常） */
  subsetted: boolean;
}

export interface FontInfo {
  /** `w:name`，文档里引用的名字（中文版 Word 写的是「等线」这种本地化名） */
  name: string;
  /**
   * `w:altName`：这款字体的**另一个名字**。
   *
   * 注意它有两种截然不同的用法，规范没区分，只能靠实际内容判断：
   * - 本地化名 → 英文名（「等线」→ `DengXian`）：**同一款字体**，这是查找时的第二把钥匙
   * - 找不到时的替代字体名：**另一款字体**
   *
   * 中文文档里几乎总是前者，所以查找顺序是「先按 name 找，找不到再按 altName 找」。
   */
  altName?: string;
  /** 10 个字节的 PANOSE 分类码，十六进制串。挑替代字体时按它比字形骨架 */
  panose?: string;
  /** Windows 字符集编号，十六进制串。`86` = GB2312（简体中文），`00` = ANSI */
  charset?: string;
  family: FontFamilyKind;
  pitch: FontPitch;
  embedded: EmbeddedFont[];
}

export interface FontTable {
  /** 按 `w:name` 索引 */
  byName: Record<string, FontInfo>;
  /** 出现顺序，保留下来是为了让诊断和调试输出稳定 */
  order: string[];
}

export const EMPTY_FONT_TABLE: FontTable = { byName: {}, order: [] };

export function parseFontTable(doc: XmlDocument | undefined): FontTable {
  if (doc === undefined) return structuredClone(EMPTY_FONT_TABLE);
  const byName: Record<string, FontInfo> = {};
  const order: string[] = [];

  for (const el of children(doc.root, 'w:font')) {
    const name = attr(el, 'w:name');
    if (name === undefined || name === '') continue;
    // 同名重复时后者胜（与 Word 一致），但顺序表里只留一份
    if (byName[name] === undefined) order.push(name);
    byName[name] = parseFont(el, name);
  }
  return { byName, order };
}

const EMBED_TAGS: Record<string, EmbeddedFont['style']> = {
  'w:embedRegular': 'regular',
  'w:embedBold': 'bold',
  'w:embedItalic': 'italic',
  'w:embedBoldItalic': 'boldItalic',
};

function parseFont(el: XmlElement, name: string): FontInfo {
  const out: FontInfo = {
    name,
    family: enumVal(valOf(el, 'w:family'), FAMILIES) ?? 'auto',
    pitch: enumVal(valOf(el, 'w:pitch'), PITCHES) ?? 'default',
    embedded: [],
  };
  const altName = valOf(el, 'w:altName');
  const panose = valOf(el, 'w:panose1');
  const charset = valOf(el, 'w:charset');
  if (altName !== undefined && altName !== '') out.altName = altName;
  if (panose !== undefined) out.panose = panose;
  if (charset !== undefined) out.charset = charset;

  for (const e of children(el)) {
    const style = EMBED_TAGS[e.name];
    const relId = attr(e, 'r:id');
    if (style === undefined || relId === undefined) continue;
    const embed: EmbeddedFont = { style, relId, subsetted: attr(e, 'w:subsetted') === 'true' };
    const key = attr(e, 'w:fontKey');
    if (key !== undefined) embed.fontKey = key;
    out.embedded.push(embed);
  }
  return out;
}

function valOf(el: XmlElement, name: string): string | undefined {
  return attrOf(child(el, name), 'w:val');
}

/**
 * 一个字体名的**候选名字列表**，按优先级排列。
 *
 * 这就是 `@uw/fonts` 查字体时该用的顺序：文档里写的名字优先，查不到再试 `altName`。
 * 反过来（先试英文名）会在两款字体互为 altName 时挑错人。
 *
 * 字体表里没有这个名字时返回它自己 —— 文档引用了字体表没登记的字体是合法的，
 * 不是错误，只是查找少了一条线索。
 */
export function fontNameCandidates(table: FontTable, name: string): string[] {
  const info = table.byName[name];
  if (info?.altName === undefined || info.altName === name) return [name];
  return [name, info.altName];
}
