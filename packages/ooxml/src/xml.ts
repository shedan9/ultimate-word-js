/**
 * XML ↔ 纯数据树。
 *
 * 这一层只做一件事：把字节变成**保序、保属性、保未知元素**的普通对象树，
 * 不认识 w:p 也不认识 w:r —— 语义是 `@uw/model` 的事。
 *
 * 为什么不直接用 fast-xml-parser 的输出：它 `preserveOrder` 模式产出的是
 * `[{ 'w:p': [...], ':@': {...} }]` 这种「键即标签名」的形状，消费方每读一个节点
 * 都要先 `Object.keys()[0]` 找出标签名，还得躲开 `:@`。转成显式的 `{kind,name,attrs,children}`
 * 一次，后面所有遍历代码都省事，而且这棵树满足原则 1.1（可结构化克隆），
 * 能直接过 Worker 边界、能当 golden file 存进仓库。
 *
 * 保留注释与 XML 声明是为了原则 1.4（round-trip 安全）。Word 自己不写 XML 注释，
 * 但用户的文档可能经过别的工具。
 */
import { UwError, UwErrorCode } from '@uw/core';
import { XMLParser } from 'fast-xml-parser';

export type XmlNode = XmlElement | XmlText | XmlComment;

export interface XmlElement {
  kind: 'element';
  /** 带前缀的原始标签名，如 `w:p`。**不做**命名空间展开 —— 见文件末尾的说明 */
  name: string;
  attrs: Record<string, string>;
  children: XmlNode[];
}

export interface XmlText {
  kind: 'text';
  text: string;
}

export interface XmlComment {
  kind: 'comment';
  text: string;
}

export interface XmlDocument {
  /** `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` 的属性，回写时吐回去 */
  declaration: Record<string, string> | undefined;
  root: XmlElement;
}

const TEXT_KEY = '#text';
const ATTRS_KEY = ':@';
const COMMENT_KEY = '#comment';
const DECLARATION_KEY = '?xml';

/**
 * 解析器配置每一项都是有原因的，别随手改：
 * - `trimValues: false` —— `<w:t xml:space="preserve"> </w:t>` 里的空格是正文内容，
 *   trim 掉会让「张三 李四」变成「张三李四」
 * - `parseTagValue` / `parseAttributeValue: false` —— 一切保持字符串。
 *   `w:val="0011"` 变成数字 11 就再也回不去了，`w:val="00"`（布尔假）会变成 0
 * - `alwaysCreateTextNode: true` —— 让文本节点形状统一，省掉一堆分支
 * - `htmlEntities: true` —— 名字有误导：**数值字符引用（`&#9;` / `&#x41;`）只有开这个才解**，
 *   而它们是 XML 核心语法不是 HTML 扩展。不开的话 `w:val="a&#9;b"` 会读成字面量
 *   `a&#9;b`，回写时那个 `&` 还会被再转义一次成 `&amp;#9;` —— 悄悄改坏用户的文档。
 *   顺带认得 `&nbsp;` 之类 HTML 命名实体，那本来就不是合法 XML，解出来只会更好不会更坏
 */
const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '',
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
  alwaysCreateTextNode: true,
  processEntities: true,
  htmlEntities: true,
  commentPropName: COMMENT_KEY,
});

/** fast-xml-parser `preserveOrder` 模式下的一个条目：唯一的标签键 + 可选的 `:@` 属性键 */
type FxpEntry = Record<string, unknown>;

export function parseXml(xml: string, partName?: string): XmlDocument {
  let entries: FxpEntry[];
  try {
    entries = parser.parse(xml) as FxpEntry[];
  } catch (cause) {
    throw new UwError(UwErrorCode.MALFORMED_XML, `XML 解析失败${partName ? `：${partName}` : ''}`, {
      ...(partName === undefined ? {} : { part: partName }),
      cause,
    });
  }

  let declaration: Record<string, string> | undefined;
  let root: XmlElement | undefined;

  for (const entry of entries) {
    const name = tagNameOf(entry);
    if (name === DECLARATION_KEY) {
      declaration = attrsOf(entry);
      continue;
    }
    if (name === TEXT_KEY || name === COMMENT_KEY) continue; // 根元素之外的文本/注释，丢掉无损
    const node = toNode(entry, name);
    if (node.kind === 'element' && root === undefined) root = node;
  }

  if (root === undefined) {
    throw new UwError(UwErrorCode.MALFORMED_XML, `XML 里没有根元素${partName ? `：${partName}` : ''}`, {
      ...(partName === undefined ? {} : { part: partName }),
    });
  }
  return { declaration, root };
}

function tagNameOf(entry: FxpEntry): string {
  for (const key of Object.keys(entry)) {
    if (key !== ATTRS_KEY) return key;
  }
  return TEXT_KEY;
}

function attrsOf(entry: FxpEntry): Record<string, string> {
  const raw = entry[ATTRS_KEY];
  if (raw === undefined || raw === null) return {};
  const out: Record<string, string> = {};
  for (const [k, v] of Object.entries(raw as Record<string, unknown>)) out[k] = String(v);
  return out;
}

function toNode(entry: FxpEntry, name: string): XmlNode {
  const value = entry[name];
  if (name === TEXT_KEY) return { kind: 'text', text: String(value ?? '') };
  if (name === COMMENT_KEY) return { kind: 'comment', text: textOf(value) };

  const children: XmlNode[] = [];
  for (const child of (value ?? []) as FxpEntry[]) {
    children.push(toNode(child, tagNameOf(child)));
  }
  return { kind: 'element', name, attrs: attrsOf(entry), children };
}

/** 注释体：fxp 把它包成 `[{ '#text': '...' }]` */
function textOf(value: unknown): string {
  if (!Array.isArray(value)) return String(value ?? '');
  let out = '';
  for (const item of value as FxpEntry[]) out += String(item[TEXT_KEY] ?? '');
  return out;
}

// ── 回写 ──────────────────────────────────────────────────────────────────────

/**
 * 树 → XML 文本。是 `parseXml` 的逆运算，两者必须成对改。
 *
 * 不保证与原文**逐字节**相同（空元素写法、属性引号、实体的可选转义都有多种合法形式），
 * 保证的是**语义等价**：`parse(serialize(parse(x)))` 深等于 `parse(x)`。
 * 这才是 round-trip 安全真正需要的东西 —— Word 关心的是语义，不是字节。
 */
export function serializeXml(doc: XmlDocument): string {
  const head =
    doc.declaration === undefined
      ? ''
      : `<?xml${Object.entries(doc.declaration)
          .map(([k, v]) => ` ${k}="${escapeAttr(v)}"`)
          .join('')}?>`;
  return head + serializeNode(doc.root);
}

function serializeNode(node: XmlNode): string {
  if (node.kind === 'text') return escapeText(node.text);
  if (node.kind === 'comment') return `<!--${node.text}-->`;

  let out = `<${node.name}`;
  for (const [k, v] of Object.entries(node.attrs)) out += ` ${k}="${escapeAttr(v)}"`;
  if (node.children.length === 0) return `${out}/>`;
  out += '>';
  for (const child of node.children) out += serializeNode(child);
  return `${out}</${node.name}>`;
}

/** `>` 其实只在 `]]>` 里非转不可，但一律转掉更省心，且合法 */
function escapeText(s: string): string {
  return s.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
}

/**
 * 属性值里的制表符 / 换行必须转成数字实体：XML 的属性值规范化会把它们变成空格，
 * 不转的话 `w:val="a&#9;b"` 读回来就成了 `a b`。
 */
function escapeAttr(s: string): string {
  return escapeText(s)
    .replaceAll('"', '&quot;')
    .replaceAll('\t', '&#x9;')
    .replaceAll('\n', '&#xA;')
    .replaceAll('\r', '&#xD;');
}

// ── 遍历辅助 ──────────────────────────────────────────────────────────────────
// model 层会大量用到，放这里免得每个消费方各写一份。

export function isElement(node: XmlNode): node is XmlElement {
  return node.kind === 'element';
}

/** 直接子元素里第一个叫 `name` 的 */
export function child(parent: XmlElement, name: string): XmlElement | undefined {
  for (const c of parent.children) {
    if (c.kind === 'element' && c.name === name) return c;
  }
  return undefined;
}

/** 直接子元素里所有叫 `name` 的（不传则全部子元素） */
export function children(parent: XmlElement, name?: string): XmlElement[] {
  const out: XmlElement[] = [];
  for (const c of parent.children) {
    if (c.kind === 'element' && (name === undefined || c.name === name)) out.push(c);
  }
  return out;
}

export function attr(el: XmlElement, name: string): string | undefined {
  return el.attrs[name];
}

/** 元素下所有文本节点拼起来（`w:t` 取值用） */
export function textContent(el: XmlElement): string {
  let out = '';
  for (const c of el.children) {
    if (c.kind === 'text') out += c.text;
    else if (c.kind === 'element') out += textContent(c);
  }
  return out;
}

// ── 关于命名空间 ──────────────────────────────────────────────────────────────
// 标签名保留原始前缀（`w:p` 而不是展开成 `{http://…/wordprocessingml/2006/main}p`）。
// 理由：docx 的前缀事实上是固定的（w / r / a / wp / mc…），Word 自己也按前缀写死；
// 展开命名空间要维护前缀栈、要处理默认命名空间，成本换来的收益在这个项目里是零。
// 代价是遇到用非常规前缀的文档会认不出 —— 真碰上了再在这里加一层前缀归一化，
// 位置是对的，改动是局部的。
