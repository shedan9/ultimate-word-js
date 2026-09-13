/**
 * 打印 —— 按文档自带的页面设置分页（api.md §13），**不重排**。
 *
 * 屏幕上的那一份不能直接拿去打：它住在宿主的滚动容器里（`overflow:auto` 在打印时只印
 * 第一屏）、带着屏幕缩放、页与页之间还隔着页间距。所以打印走的是**另一份 DOM**：
 * 一张挂在 `<body>` 直下的「打印页」，每页一个 pt 尺寸的盒子，里面是同一份
 * `PageLayout` 在 zoom = 1 下的重画，`@page { size }` 取纸张尺寸、边距 0 ——
 * 页面设置里的边距早已由布局算进版心，打印机再加一遍就成了双份。
 *
 * 为什么不走 `<iframe>`：iframe 拿不到宿主的 `@font-face` 与 `document.fonts`，
 * 挂在同一个 document 里才能与屏幕上用的是同一批字体；代价是要用一段打印样式把
 * `<body>` 的其余直接子元素藏起来（`display:none` 只在 `@media print` 里生效）。
 *
 * 装饰与 overlay **不打**：这一份是从布局重画的，不是屏幕 DOM 的克隆 —— Word 打批注也是
 * 可选项，等有人要再做成选项。
 *
 * 这个文件只造 DOM 与样式串，不碰 `window.print()`；调不调打印对话框是 `dom.ts` 的事，
 * 样式串是纯函数（`printStyles`），Node 里就能测。
 */
import { twipsToPt } from '@uw/core';
import type { PageLayout } from '@uw/layout';
import type { MountOptions } from '@uw/render-dom/dom';
import { renderPage } from '@uw/render-dom/dom';

/** 打印页的标记属性。**与 classPrefix 无关**：同一页面上几个视图要靠它认出「已经有人造过了」 */
export const PRINT_SHEET_ATTR = 'data-uw-print';

export interface PageSize {
  /** twips */
  width: number;
  height: number;
}

export interface PrintSheet {
  readonly element: HTMLElement;
  dispose(): void;
}

function fmt(n: number): string {
  return String(Math.round(n * 100) / 100);
}

/**
 * 每页归到哪一种纸张：第一种做默认 `@page`，其余各起一个命名页（`page: <name>`，
 * Chrome 85+ / Firefox 110+ 认，不认的浏览器把它印在默认纸张上 —— 顶多裁掉一角，不会少页）。
 * 答每页的命名页名（默认纸张为 undefined）与去重后的纸张表。
 */
export function paperOf(
  pages: readonly PageSize[],
  prefix: string,
): { names: (string | undefined)[]; papers: { name: string | undefined; size: PageSize }[] } {
  const papers: { name: string | undefined; size: PageSize }[] = [];
  const names = pages.map((size) => {
    let paper = papers.find((p) => p.size.width === size.width && p.size.height === size.height);
    if (paper === undefined) {
      paper = { name: papers.length === 0 ? undefined : `${prefix}-print-${papers.length}`, size };
      papers.push(paper);
    }
    return paper.name;
  });
  return { names, papers };
}

/** 打印样式串。只在 `@media print` 里起作用，屏幕上打印页整个 `display:none` */
export function printStyles(pages: readonly PageSize[], prefix: string): string {
  const { papers } = paperOf(pages, prefix);
  const sel = `[${PRINT_SHEET_ATTR}]`;
  const rules: string[] = [];
  for (const { name, size } of papers) {
    rules.push(
      `@page ${name ?? ''}{size:${fmt(twipsToPt(size.width))}pt ${fmt(twipsToPt(size.height))}pt;margin:0}`,
    );
    if (name !== undefined) rules.push(`${sel}>.${name}{page:${name}}`);
  }
  return [
    `@media screen{${sel}{display:none}}`,
    '@media print{',
    ...rules,
    // 宿主常把 body 钉成一屏高 + overflow:hidden（应用壳的惯用布局），打印时只会印出第一页
    'html,body{margin:0!important;padding:0!important;height:auto!important;overflow:visible!important;background:#fff}',
    `body>:not(${sel}){display:none!important}`,
    // 行高归零：盒子之间的空白文本节点会把下一页顶出一张空纸
    `${sel}{display:block!important;line-height:0}`,
    `${sel}>div{overflow:hidden;break-inside:avoid;page-break-inside:avoid}`,
    // 用 break-before 而不是 break-after：最后一页后面的强制分页会在部分浏览器里多印一张白纸
    `${sel}>div+div{break-before:page;page-break-before:always}`,
    '}',
  ].join('\n');
}

/**
 * 造一张打印页并挂到 `<body>` 直下。**要重画每一页**（不复用屏幕上的绘制层 ——
 * 虚拟化之下大多数页根本没画），所以只在真要打印那一刻造，打完就拆。
 * 没有 `<body>`（脱离文档的容器）时答 undefined，调用方退回原地打印。
 */
export function buildPrintSheet(
  document: Document,
  pages: readonly PageLayout[],
  opts: MountOptions,
): PrintSheet | undefined {
  const body = document.body;
  if (body === null) return undefined;
  const prefix = opts.classPrefix ?? 'uw';
  const sizes = pages.map((p) => p.geometry);
  const { names } = paperOf(sizes, prefix);
  const sheet = document.createElement('div');
  sheet.setAttribute(PRINT_SHEET_ATTR, '');
  sheet.className = `${prefix}-print`;
  sheet.setAttribute('aria-hidden', 'true');
  const style = document.createElement('style');
  style.textContent = printStyles(sizes, prefix);
  sheet.append(style);
  const paint: MountOptions = { ...opts, document, zoom: 1 };
  pages.forEach((page, i) => {
    const box = document.createElement('div');
    const name = names[i];
    box.className = `${prefix}-print-page${name === undefined ? '' : ` ${name}`}`;
    box.dataset.page = String(page.index);
    box.style.cssText = `width:${fmt(twipsToPt(page.geometry.width))}pt;height:${fmt(twipsToPt(page.geometry.height))}pt;position:relative;box-sizing:border-box`;
    const svg = renderPage(page, paint) as SVGSVGElement;
    svg.style.cssText = 'display:block;width:100%;height:100%';
    box.append(svg);
    sheet.append(box);
  });
  body.append(sheet);
  return {
    element: sheet,
    dispose: () => sheet.remove(),
  };
}
