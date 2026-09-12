/**
 * 原生 SVG 选区不保证在行间插入换行。只在选区完全属于本视图时接管纯文本复制，
 * 保留同一行内片段的拼接，跨行 / 跨单元格 / 跨页以换行分开。
 * 不从绘制层取字：那里包含编号和重复表头，而且已经设为 inert。
 */
export function selectedText(root: Element, selection: Selection): string | undefined {
  if (selection.isCollapsed || selection.rangeCount !== 1) return undefined;
  const range = selection.getRangeAt(0);
  if (!root.contains(range.startContainer) || !root.contains(range.endContainer)) return undefined;
  if (
    range.startContainer.parentElement?.closest('[data-overlay]') ||
    range.endContainer.parentElement?.closest('[data-overlay]')
  )
    return undefined;
  const lines: string[] = [];
  for (const line of root.querySelectorAll('[data-copy-line]')) {
    let text = '';
    for (const element of line.querySelectorAll('tspan')) {
      const node = element.firstChild;
      if (node === null || node.nodeType !== 3 || !range.intersectsNode(node)) continue;
      const content = node.textContent ?? '';
      const start = range.startContainer === node ? range.startOffset : 0;
      const end = range.endContainer === node ? range.endOffset : content.length;
      text += content.slice(start, end);
    }
    if (text !== '') lines.push(text);
  }
  return lines.length === 0 ? undefined : lines.join('\n');
}
