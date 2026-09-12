import { twipsToPt } from '@uw/core';
import type { IndexedLine, PageLayout } from '@uw/layout';
import type { RElement, RenderOptions } from '@uw/render-dom';
import { buildTextFragment, el, fmt, fmtList } from '@uw/render-dom';

function imageAlternative(x: number, y: number, width: number, height: number, label: string): RElement {
  return el('rect', {
    role: 'img',
    'aria-label': label,
    fill: 'transparent',
    x: fmt(twipsToPt(x)),
    y: fmt(twipsToPt(y)),
    width: fmt(twipsToPt(width)),
    height: fmt(twipsToPt(height)),
    'pointer-events': 'none',
  });
}

/**
 * 常驻的轻量文字 SVG。直接复用绘制片段，原生选区与画面采用同一套字形坐标。
 * 填充透明而非 display:none / visibility:hidden，浏览器查找与辅助技术仍能读取。
 * 重复表头只保留原本那一份，编号跳过，域的计算结果正常保留。
 */
export function buildTextLayer(
  page: PageLayout,
  lines: readonly IndexedLine[],
  options: RenderOptions = {},
  enabled = true,
): RElement {
  const children: RElement[] = [];
  if (enabled) {
    for (const line of lines) {
      if (line.page !== page.index || line.repeated) continue;
      const fragments = line.line.fragments.filter((fragment) => !fragment.numbering && fragment.text !== '');
      if (fragments.length > 0)
        children.push(
          el(
            'text',
            { 'data-copy-line': '', 'xml:space': 'preserve' },
            fragments.map((fragment) => {
              const text = buildTextFragment(fragment, line.originX, line.top + line.line.baseline, options);
              text.tag = 'tspan';
              // tspan 的 transform 在浏览器中不能可靠参与文字布局；用原始字形位置与
              // textLength 表达横向压缩，保持同一行是一棵可跨样式搜索的 text 子树。
              if (fragment.style.scale !== 100 && fragment.style.scale !== 0) {
                delete text.attrs.transform;
                text.attrs.x = fmtList(fragment.glyphX.map((x) => twipsToPt(line.originX + x)));
                text.attrs.textLength = fmt(twipsToPt(fragment.width));
                text.attrs.lengthAdjust = 'spacingAndGlyphs';
              }
              text.attrs.fill = 'transparent';
              text.attrs.style = 'user-select:text;-webkit-user-select:text;pointer-events:all;cursor:text';
              text.attrs['data-content-index'] = String(fragment.contentIndex);
              text.attrs['data-offset'] = String(fragment.offset);
              return text;
            }),
          ),
        );
      // 绘制层是 inert，图片的替代说明必须在常驻层保留；不用文本节点，避免混入复制。
      for (const object of line.line.objects ?? []) {
        children.push(
          imageAlternative(
            line.originX + object.x,
            line.top + line.line.baseline - object.height - (object.raise ?? 0),
            object.width,
            object.height,
            object.alt || object.graphic || '图片',
          ),
        );
      }
    }
    for (const object of page.floats ?? []) {
      children.push(
        imageAlternative(
          object.x,
          object.y,
          object.width,
          object.height,
          object.alt || object.graphic || '图片',
        ),
      );
    }
  }
  return el(
    'svg',
    {
      xmlns: 'http://www.w3.org/2000/svg',
      class: `${options.classPrefix ?? 'uw'}-text-layer`,
      'data-viewport-page': String(page.index),
      'data-text-layer': String(enabled),
      viewBox: `0 0 ${fmt(twipsToPt(page.geometry.width))} ${fmt(twipsToPt(page.geometry.height))}`,
      width: '100%',
      height: '100%',
      style: 'position:absolute;inset:0;overflow:visible;pointer-events:none;z-index:2',
      ...(enabled ? {} : { 'aria-hidden': 'true' }),
    },
    children,
  );
}
