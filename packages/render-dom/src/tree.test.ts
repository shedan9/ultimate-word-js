import { describe, expect, it } from 'vitest';
import { el, fmt, fmtList, serialize, textEl } from './tree.ts';

describe('fmt', () => {
  it('保留 3 位小数，且不吐出浮点尾巴', () => {
    // 0.1 + 0.2 直接 String() 会是 0.30000000000000004，快照会在不同平台上飘
    expect(fmt(0.1 + 0.2)).toBe('0.3');
    expect(fmt(119.0512345)).toBe('119.051');
    expect(fmt(12)).toBe('12');
  });

  it('-0 与 0 是同一个字符串', () => {
    // 居中对齐算出来的偏移常常是 -0：`String(-0)` 是 "-0"，golden file 会因此不稳
    expect(fmt(-0)).toBe('0');
    expect(fmt(-0.0001)).toBe('0');
  });

  it('fmtList 用空格分隔 —— SVG 的 x 列表就吃这个格式', () => {
    expect(fmtList([0, 10.5, 21])).toBe('0 10.5 21');
  });
});

describe('serialize', () => {
  it('无子节点的元素收成自闭合', () => {
    expect(serialize(el('rect', { x: '0', width: '10' }))).toBe('<rect x="0" width="10"/>');
  });

  it('文本与属性各按各的规则转义', () => {
    const node = textEl('text', { 'data-run': 'a<b', 'xml:space': 'preserve' }, 'a & b <c>');
    expect(serialize(node)).toBe('<text data-run="a&lt;b" xml:space="preserve">a &amp; b &lt;c&gt;</text>');
  });

  it('不加缩进 —— SVG 不折叠空白，缩进会变成真的空格', () => {
    const svg = el('svg', {}, [el('g', {}, [textEl('text', {}, ' 甲 ')])]);
    expect(serialize(svg)).toBe('<svg><g><text> 甲 </text></g></svg>');
  });
});
