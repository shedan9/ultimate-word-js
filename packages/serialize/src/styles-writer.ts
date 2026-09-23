/**
 * 编辑期新增的样式定义（`Body.styles`，套用文档里没有的「标题 1」时补的）→ `styles.xml`。
 *
 * 与 numbering-writer.ts 同一个思路：编辑只会**追加**定义，原有的 `w:style` 一个字节都不碰
 * （它们带着 `w:rsid`、`w:latentStyles` 里的开关，这些模型都没解析）。
 * 新样式追加在根的末尾 —— `CT_Styles` 的顺序是 docDefaults → latentStyles → style*，
 * 末尾正是 style 该在的地方。
 *
 * `w:latentStyles` 里通常已经挂着一条同名的 `w:lsdException`（`heading 1` 带着 `semiHidden` 之类）；
 * 一旦文档里有了真正的 `w:style`，Word 以它为准，latent 那条不必改。
 */
import type { StyleDefinition, StyleSheet } from '@uw/model';
import type { XmlDocument, XmlElement, XmlNode } from '@uw/ooxml';
import { W_NS } from './numbering-writer.ts';
import { patchParaProps, patchRunProps } from './props-writer.ts';
import { el } from './xml-edit.ts';

/** 原样式表里没有的那些定义；没有就是 undefined（不必动 styles.xml） */
export function addedStyles(
  original: StyleSheet,
  edited: readonly StyleDefinition[] | undefined,
): StyleDefinition[] | undefined {
  const added = (edited ?? []).filter((d) => original.byId(d.id) === undefined);
  return added.length ? added : undefined;
}

/** `base` 是原来的 styles.xml（没有就新建一份） */
export function writeStyles(base: XmlDocument | undefined, added: readonly StyleDefinition[]): XmlDocument {
  const doc = base ?? {
    declaration: { version: '1.0', encoding: 'UTF-8', standalone: 'yes' },
    root: el('w:styles', { 'xmlns:w': W_NS }),
  };
  return {
    declaration: doc.declaration,
    root: { ...doc.root, children: [...doc.root.children, ...added.map(styleElement)] },
  };
}

/** `CT_Style` 的子元素顺序：name, aliases, basedOn, next, link, …, uiPriority, semiHidden, unhideWhenUsed, qFormat, …, pPr, rPr */
function styleElement(d: StyleDefinition): XmlElement {
  const children: XmlNode[] = [el('w:name', { 'w:val': d.name })];
  if (d.basedOn !== undefined) children.push(el('w:basedOn', { 'w:val': d.basedOn }));
  if (d.next !== undefined) children.push(el('w:next', { 'w:val': d.next }));
  if (d.uiPriority !== undefined) children.push(el('w:uiPriority', { 'w:val': String(d.uiPriority) }));
  if (d.quickFormat) children.push(el('w:qFormat'));
  const pPr = patchParaProps(undefined, d.paraProps);
  if (pPr !== undefined) children.push(pPr);
  const rPr = patchRunProps(undefined, d.runProps);
  if (rPr !== undefined) children.push(rPr);
  return el('w:style', { 'w:type': 'paragraph', 'w:styleId': d.id }, children);
}
