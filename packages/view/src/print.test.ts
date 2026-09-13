import { describe, expect, it } from 'vitest';
import { PRINT_SHEET_ATTR, paperOf, printStyles } from './print.ts';

const A4 = { width: 11906, height: 16838 };
const A4_LANDSCAPE = { width: 16838, height: 11906 };

describe('paperOf', () => {
  it('第一种纸张是默认 @page，之后每种新纸张各起一个命名页', () => {
    const { names, papers } = paperOf([A4, A4, A4_LANDSCAPE, A4], 'uw');
    expect(names).toEqual([undefined, undefined, 'uw-print-1', undefined]);
    expect(papers.map((p) => p.name)).toEqual([undefined, 'uw-print-1']);
  });

  it('宽高相同才算同一种纸张 —— 横向 A4 与纵向 A4 是两种', () => {
    expect(paperOf([A4_LANDSCAPE, A4], 'x').papers).toHaveLength(2);
  });
});

describe('printStyles', () => {
  const css = printStyles([A4, A4_LANDSCAPE], 'uw');

  it('@page 的尺寸是 pt 且边距为 0 —— 页边距早已算进版心，打印机再加一遍就是双份', () => {
    expect(css).toContain('@page {size:595.3pt 841.9pt;margin:0}');
    expect(css).toContain('@page uw-print-1{size:841.9pt 595.3pt;margin:0}');
    expect(css).toContain(`[${PRINT_SHEET_ATTR}]>.uw-print-1{page:uw-print-1}`);
  });

  it('屏幕上整个藏起来、打印时藏起 body 的其余直接子元素', () => {
    expect(css).toMatch(/@media screen\{\[data-uw-print\]\{display:none\}\}/);
    expect(css).toContain(`body>:not([${PRINT_SHEET_ATTR}]){display:none!important}`);
    expect(css).toContain('overflow:visible!important');
  });

  it('分页用 break-before 而不是 break-after，最后一页后面不多印白纸', () => {
    expect(css).toContain('>div+div{break-before:page');
    expect(css).not.toContain('break-after');
  });
});
