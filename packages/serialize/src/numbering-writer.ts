/**
 * 新建列表加的编号定义 → `numbering.xml`。
 *
 * 编辑期只会**追加**定义（`addListDefinition`，每次新的一对 abstractNum + num），
 * 不改已有的 —— 所以这里只比「多了哪些 id」，原有定义一个字节都不碰（它们带着
 * `w:nsid` / `w:tmpl` / 图片项目符号，这些模型都没解析）。
 *
 * 顺序有硬规定：`w:numbering` 里**全部 `w:abstractNum` 在全部 `w:num` 之前**（§17.9.20），
 * 新的 abstractNum 不能直接追加到末尾 —— 那会落在 num 后面，Word 报文件损坏。
 */
import type { AbstractNumbering, Numbering, NumberingInstance, NumberingLevel } from '@uw/model';
import type { XmlDocument, XmlElement, XmlNode } from '@uw/ooxml';
import { patchParaProps, patchRunProps } from './props-writer.ts';
import { el } from './xml-edit.ts';

export const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** 编辑后多出来的定义；没有就是 undefined（不必动 numbering.xml） */
export function addedNumbering(original: Numbering, edited: Numbering | undefined): Numbering | undefined {
  if (edited === undefined) return undefined;
  const abstract: Record<number, AbstractNumbering> = {};
  const instances: Record<number, NumberingInstance> = {};
  for (const [id, a] of Object.entries(edited.abstract))
    if (original.abstract[Number(id)] === undefined) abstract[Number(id)] = a;
  for (const [id, n] of Object.entries(edited.instances))
    if (original.instances[Number(id)] === undefined) instances[Number(id)] = n;
  if (Object.keys(abstract).length === 0 && Object.keys(instances).length === 0) return undefined;
  return { abstract, instances };
}

/** `base` 是原来的 numbering.xml（没有就新建一份） */
export function writeNumbering(base: XmlDocument | undefined, added: Numbering): XmlDocument {
  const doc = base ?? {
    declaration: { version: '1.0', encoding: 'UTF-8', standalone: 'yes' },
    root: el('w:numbering', { 'xmlns:w': W_NS }),
  };
  const children = [...doc.root.children];
  const abstracts = Object.values(added.abstract).map(abstractElement);
  const nums = Object.values(added.instances).map(numElement);

  const lastIndex = (name: string) => children.findLastIndex((c) => c.kind === 'element' && c.name === name);
  const firstIndex = (name: string) => children.findIndex((c) => c.kind === 'element' && c.name === name);
  // abstractNum：接在最后一个 abstractNum 后面；一个都没有就放在第一个 num 前面
  let at = lastIndex('w:abstractNum');
  at = at >= 0 ? at + 1 : firstIndex('w:num') >= 0 ? firstIndex('w:num') : afterPictureBullets(children);
  children.splice(at, 0, ...abstracts);
  // num：接在最后一个 num 后面；`w:numIdMacAtCleanup` 规定排在所有 num 之后
  const lastNum = lastIndex('w:num');
  const cleanup = firstIndex('w:numIdMacAtCleanup');
  children.splice(lastNum >= 0 ? lastNum + 1 : cleanup >= 0 ? cleanup : children.length, 0, ...nums);
  return { declaration: doc.declaration, root: { ...doc.root, children } };
}

/** 没有 abstractNum 也没有 num 时，位置在图片项目符号（`w:numPicBullet`，schema 里排最前）之后 */
function afterPictureBullets(children: readonly XmlNode[]): number {
  const last = children.findLastIndex((c) => c.kind === 'element' && c.name === 'w:numPicBullet');
  return last + 1;
}

function abstractElement(a: AbstractNumbering): XmlElement {
  const children: XmlNode[] = [el('w:multiLevelType', { 'w:val': a.multiLevelType })];
  for (const level of Object.values(a.levels).sort((x, y) => x.level - y.level))
    children.push(levelElement(level));
  return el('w:abstractNum', { 'w:abstractNumId': String(a.id) }, children);
}

/** `CT_Lvl` 的子元素顺序：start, numFmt, lvlRestart, pStyle, isLgl, suff, lvlText, lvlJc, pPr, rPr */
function levelElement(l: NumberingLevel): XmlElement {
  const children: XmlNode[] = [
    el('w:start', { 'w:val': String(l.start) }),
    el('w:numFmt', { 'w:val': l.numFmt }),
  ];
  if (l.restartAfter !== undefined) children.push(el('w:lvlRestart', { 'w:val': String(l.restartAfter) }));
  if (l.paraStyleId !== undefined) children.push(el('w:pStyle', { 'w:val': l.paraStyleId }));
  if (l.isLegal) children.push(el('w:isLgl'));
  // 缺省就是 tab，不写与 Word 自己存的一致
  if (l.suffix !== 'tab') children.push(el('w:suff', { 'w:val': l.suffix }));
  children.push(el('w:lvlText', { 'w:val': l.lvlText }), el('w:lvlJc', { 'w:val': l.justification }));
  const pPr = patchParaProps(undefined, l.paraProps);
  if (pPr !== undefined) children.push(pPr);
  const rPr = patchRunProps(undefined, l.runProps);
  if (rPr !== undefined) children.push(rPr);
  return el('w:lvl', { 'w:ilvl': String(l.level) }, children);
}

function numElement(n: NumberingInstance): XmlElement {
  const children: XmlNode[] = [el('w:abstractNumId', { 'w:val': String(n.abstractNumId) })];
  for (const [level, o] of Object.entries(n.overrides)) {
    const inner: XmlNode[] = [];
    if (o.start !== undefined) inner.push(el('w:startOverride', { 'w:val': String(o.start) }));
    if (o.level !== undefined) inner.push(levelElement(o.level));
    children.push(el('w:lvlOverride', { 'w:ilvl': level }, inner));
  }
  return el('w:num', { 'w:numId': String(n.numId) }, children);
}
