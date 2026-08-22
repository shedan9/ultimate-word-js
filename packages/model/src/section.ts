/**
 * `w:sectPr` → `SectionProps`。
 *
 * 版心尺寸由这里决定，而版心是**所有**布局坐标的原点，所以这个文件错一个字段，
 * 后面每一行的 x/y 全错。字段都是 twips（`w:pgSz` / `w:pgMar` 本来就是 twips，不用换算）。
 */
import type { Twips } from '@uw/core';
import type { XmlElement } from '@uw/ooxml';
import { attr, child, children } from '@uw/ooxml';
import type { DocGrid, HeaderFooterRef, SectionProps, SectionStart } from './nodes.ts';
import { attrInt, attrOf, enumVal, onOff, put } from './xml-values.ts';

const GRID_TYPES = ['default', 'lines', 'linesAndChars', 'snapToChars'] as const;
const HF_TYPES = ['default', 'first', 'even'] as const;
const SECTION_TYPES: readonly SectionStart[] = [
  'nextPage',
  'continuous',
  'nextColumn',
  'evenPage',
  'oddPage',
];

/**
 * `w:sectPr` 缺席时的兜底 —— A4 纵向 + 中文版 Word 默认页边距。
 *
 * 真实文档必有 `sectPr`，走到这里说明文件不正常，因此 `parseSectionProps` 的调用方
 * 会同时发一条诊断。这些值只是让引擎能继续画出东西来，不是「合理的默认」。
 */
export const DEFAULT_SECTION_PROPS: SectionProps = {
  page: { width: 11906, height: 16838, orientation: 'portrait' },
  type: 'nextPage',
  margin: { top: 1440, right: 1800, bottom: 1440, left: 1800, header: 851, footer: 992, gutter: 0 },
  docGrid: { type: 'default', linePitch: 0, charSpace: 0 },
  columns: 1,
  titlePage: false,
  headers: [],
  footers: [],
};

export function parseSectionProps(sectPr: XmlElement | undefined): SectionProps {
  if (sectPr === undefined) return structuredClone(DEFAULT_SECTION_PROPS);

  const d = DEFAULT_SECTION_PROPS;
  const pgSz = child(sectPr, 'w:pgSz');
  const pgMar = child(sectPr, 'w:pgMar');
  const grid = child(sectPr, 'w:docGrid');

  const out: SectionProps = {
    page: {
      width: attrInt(pgSz, 'w:w') ?? d.page.width,
      height: attrInt(pgSz, 'w:h') ?? d.page.height,
      // 注意 orient 只是个声明：**Word 不会因为它去交换 w/h**，w:w 已经是横放后的宽度。
      // 拿它去转置尺寸会把横向页面转回竖的
      orientation: attrOf(pgSz, 'w:orient') === 'landscape' ? 'landscape' : 'portrait',
    },
    // 缺席按 nextPage（规范默认）。注意它说的是**本节**从哪儿开始，见 SectionStart
    type: enumVal(attrOf(child(sectPr, 'w:type'), 'w:val'), SECTION_TYPES) ?? d.type,
    margin: {
      top: attrInt(pgMar, 'w:top') ?? d.margin.top,
      right: attrInt(pgMar, 'w:right') ?? d.margin.right,
      bottom: attrInt(pgMar, 'w:bottom') ?? d.margin.bottom,
      left: attrInt(pgMar, 'w:left') ?? d.margin.left,
      header: attrInt(pgMar, 'w:header') ?? d.margin.header,
      footer: attrInt(pgMar, 'w:footer') ?? d.margin.footer,
      gutter: attrInt(pgMar, 'w:gutter') ?? d.margin.gutter,
    },
    docGrid: parseDocGrid(grid),
    columns: attrInt(child(sectPr, 'w:cols'), 'w:num') ?? 1,
    titlePage: onOff(sectPr, 'w:titlePg') ?? false,
    headers: hfRefs(sectPr, 'w:headerReference'),
    footers: hfRefs(sectPr, 'w:footerReference'),
  };

  const pgNumType = child(sectPr, 'w:pgNumType');
  put(out, 'pageNumStart', attrInt(pgNumType, 'w:start'));
  put(out, 'pageNumFormat', attrOf(pgNumType, 'w:fmt'));
  return out;
}

function parseDocGrid(grid: XmlElement | undefined): DocGrid {
  if (grid === undefined) return { type: 'default', linePitch: 0, charSpace: 0 };
  return {
    type: enumVal(attrOf(grid, 'w:type'), GRID_TYPES) ?? 'default',
    linePitch: (attrInt(grid, 'w:linePitch') ?? 0) as Twips,
    charSpace: attrInt(grid, 'w:charSpace') ?? 0,
  };
}

function hfRefs(sectPr: XmlElement, name: string): HeaderFooterRef[] {
  const out: HeaderFooterRef[] = [];
  for (const el of children(sectPr, name)) {
    const relId = attr(el, 'r:id');
    // 没有 r:id 的引用指不到任何部件，留着只会让 Phase 3 空跑一趟
    if (relId === undefined) continue;
    out.push({ type: enumVal(attr(el, 'w:type'), HF_TYPES) ?? 'default', relId });
  }
  return out;
}
