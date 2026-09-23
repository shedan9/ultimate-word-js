/**
 * 造 / 改 XML 树的小工具。
 *
 * 一律**不改入参**：源元素来自 `OpcPackage` 的解析缓存，改了它等于改了「原文」，
 * 同一个包再回写一次就对不上了。要改就浅拷一层（`{ ...el, attrs, children }`）。
 */
import type { XmlElement, XmlNode } from '@uw/ooxml';

export function el(name: string, attrs: Record<string, string> = {}, children: XmlNode[] = []): XmlElement {
  return { kind: 'element', name, attrs, children };
}

/**
 * 按 schema 顺序插入。WordprocessingML 的属性容器（`w:pPr` / `w:rPr`）是 `xsd:sequence`，
 * **顺序错了 Word 报「无法读取的内容」**，不是宽容地忽略 —— 所以新元素不能随手 push 到末尾。
 *
 * 插在第一个「序号比它大」的已知元素前面；不认识的元素（扩展命名空间里的）没有序号，
 * 跳过它们比较，原来在哪儿还在哪儿。
 */
export function insertOrdered(children: XmlNode[], node: XmlElement, order: readonly string[]): XmlNode[] {
  const rank = order.indexOf(node.name);
  const at = children.findIndex((c) => c.kind === 'element' && order.indexOf(c.name) > rank);
  if (at < 0) return [...children, node];
  return [...children.slice(0, at), node, ...children.slice(at)];
}

export function withoutChildren(children: readonly XmlNode[], names: ReadonlySet<string>): XmlNode[] {
  return children.filter((c) => c.kind !== 'element' || !names.has(c.name));
}

/**
 * 属性级的补丁：`owned` 里的属性全部由 `values` 说了算（缺席 = 删），其余属性原样保留。
 *
 * 这是「改一个值不丢邻居」的关键：`<w:lang w:val="en-US" w:eastAsia="zh-CN"/>` 里我们只认
 * `w:eastAsia`，重写整个元素会把 `w:val` 弄丢；`<w:u w:val="single" w:color="FF0000"/>` 同理。
 */
export function patchAttrs(
  attrs: Readonly<Record<string, string>>,
  owned: readonly string[],
  values: Readonly<Record<string, string | undefined>>,
): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [k, v] of Object.entries(attrs)) if (!owned.includes(k)) out[k] = v;
  for (const k of owned) {
    const v = values[k];
    if (v !== undefined) out[k] = v;
  }
  return out;
}

/**
 * 结构相等，**不看键序**。事务改属性时常以展开重建对象（`{ ...props, bold: true }`），
 * 键序与解析出来的不同，`JSON.stringify` 比较会把没改的节点当成改了 —— 然后白白重写一遍。
 */
export function same(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (typeof a !== 'object' || typeof b !== 'object' || a === null || b === null) return false;
  if (Array.isArray(a) !== Array.isArray(b)) return false;
  if (Array.isArray(a)) {
    const bb = b as unknown[];
    return a.length === bb.length && a.every((v, i) => same(v, bb[i]));
  }
  const ka = Object.keys(a).filter((k) => (a as Record<string, unknown>)[k] !== undefined);
  const kb = Object.keys(b).filter((k) => (b as Record<string, unknown>)[k] !== undefined);
  if (ka.length !== kb.length) return false;
  return ka.every((k) => same((a as Record<string, unknown>)[k], (b as Record<string, unknown>)[k]));
}
