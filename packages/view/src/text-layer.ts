import { twipsToPt } from '@uw/core';
import type { IndexedLine, LineLayout, PageLayout, PlacedTextBox } from '@uw/layout';
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

/** 一行的透明文字（编号跳过）。没有可见文字的行返回 undefined */
function lineText(
  line: LineLayout,
  originX: number,
  baseline: number,
  options: RenderOptions,
): RElement | undefined {
  const fragments = line.fragments.filter((fragment) => !fragment.numbering && fragment.text !== '');
  if (fragments.length === 0) return undefined;
  return el(
    'text',
    { 'data-copy-line': '', 'xml:space': 'preserve' },
    fragments.map((fragment) => {
      const text = buildTextFragment(fragment, originX, baseline, options);
      text.tag = 'tspan';
      // tspan 的 transform 在浏览器中不能可靠参与文字布局；用原始字形位置与
      // textLength 表达横向压缩，保持同一行是一棵可跨样式搜索的 text 子树。
      if (fragment.style.scale !== 100 && fragment.style.scale !== 0) {
        delete text.attrs.transform;
        text.attrs.x = fmtList(fragment.glyphX.map((x) => twipsToPt(originX + x)));
        text.attrs.textLength = fmt(twipsToPt(fragment.width));
        text.attrs.lengthAdjust = 'spacingAndGlyphs';
      }
      text.attrs.fill = 'transparent';
      text.attrs.style = 'user-select:text;-webkit-user-select:text;pointer-events:all;cursor:text';
      text.attrs['data-content-index'] = String(fragment.contentIndex);
      text.attrs['data-offset'] = String(fragment.offset);
      return text;
    }),
  );
}

/**
 * 文本框里的行。它们不在布局索引里（不可编辑、没有 `DocPosition` 能指过去），
 * 所以不走索引那条路，直接照摞好的块现摊；表格里的字没收（文本框里放表格罕见）
 */
function textBoxLines(box: PlacedTextBox, options: RenderOptions): RElement[] {
  const out: RElement[] = [];
  for (const block of box.blocks) {
    if (block.kind !== 'paragraph') continue;
    for (const placed of block.lines) {
      const text = lineText(placed.line, box.x, box.y + placed.y + placed.line.baseline, options);
      if (text !== undefined) out.push(text);
    }
  }
  return out;
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
      const text = lineText(line.line, line.originX, line.top + line.line.baseline, options);
      if (text !== undefined) children.push(text);
      // 绘制层是 inert，图片的替代说明必须在常驻层保留；不用文本节点，避免混入复制。
      for (const object of line.line.objects ?? []) {
        // 内嵌文本框的字在下面随 `page.floats` 收，不是图片
        if (object.textBox !== undefined) continue;
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
      // 文本框的字进常驻层（查找、复制、读屏都要它），框本身不再补「图片」的替代说明
      if (object.textBox !== undefined) {
        children.push(...textBoxLines(object.textBox, options));
        continue;
      }
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
