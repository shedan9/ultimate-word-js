/**
 * 元素树 → 真 DOM。**本包里唯一碰浏览器 API 的文件**，别把 `document` 泄漏到别处去。
 *
 * `Document` 是**注入**的，不是直接读全局：一来 SSR / 离屏渲染时全局没有 `document`，
 * 二来它让「渲染器需要 DOM」这件事变成一个显式参数，而不是一个隐藏的环境依赖。
 * 默认值仍然是全局的那个，浏览器里的调用方感觉不到区别。
 */
import type { DocumentLayout, PageLayout } from '@uw/layout';
import type { RenderOptions } from './paint.ts';
import { buildDocument, buildPage } from './paint.ts';
import type { RElement } from './tree.ts';

const SVG_NS = 'http://www.w3.org/2000/svg';
const HTML_NS = 'http://www.w3.org/1999/xhtml';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';

export interface MountOptions extends RenderOptions {
  /** 造元素用的 `Document`。缺省取全局的那个 */
  document?: Document;
}

/**
 * 元素树 → DOM 节点。
 *
 * 命名空间靠**进没进过 `<svg>`** 判断，不靠标签名白名单：`<a>` `<title>` `<style>`
 * 在两个命名空间里都存在，按名字猜必然有猜错的一天。一旦进了 svg 子树就全是 SVG，
 * 这与 HTML 解析器的行为一致（foreignObject 是唯一的例外，本渲染器不产出它）。
 */
export function toDom(node: RElement, doc: Document = globalThis.document, svg = false): Element {
  const inSvg = svg || node.tag === 'svg';
  const element = doc.createElementNS(inSvg ? SVG_NS : HTML_NS, node.tag);
  for (const [k, v] of Object.entries(node.attrs)) {
    // `xmlns` 只是序列化时给 XML 解析器看的；DOM 里命名空间已经由 createElementNS 定了，
    // 再 setAttribute 一次在部分浏览器上直接抛 NamespaceError
    if (k === 'xmlns') continue;
    if (k.startsWith('xml:')) element.setAttributeNS(XML_NS, k, v);
    else element.setAttribute(k, v);
  }
  if (node.text !== undefined) element.appendChild(doc.createTextNode(node.text));
  for (const child of node.children) element.appendChild(toDom(child, doc, inSvg));
  return element;
}

/** 整份文档 → 一个 `<div>` 元素，直接 append 就能看 */
export function renderDocument(layout: DocumentLayout, opts: MountOptions = {}): Element {
  return toDom(buildDocument(layout, opts), opts.document ?? globalThis.document);
}

/** 单页 → 一个 `<svg>` 元素。视口虚拟化按页装卸时走这个入口 */
export function renderPage(page: PageLayout, opts: MountOptions = {}): Element {
  return toDom(buildPage(page, opts), opts.document ?? globalThis.document);
}

/**
 * 画进容器，**先清空**。
 *
 * 这是「全量重画」，v1 只有这一种：增量更新要比对前后两棵元素树，而那件事的前提是
 * 增量排版（架构 §7）还没做 —— 现在每次重排本来就是全量的，做局部 DOM 更新没有收益。
 */
export function mount(container: Element, layout: DocumentLayout, opts: MountOptions = {}): Element {
  const doc = opts.document ?? container.ownerDocument ?? globalThis.document;
  container.replaceChildren();
  const root = toDom(buildDocument(layout, opts), doc);
  container.appendChild(root);
  return root;
}
