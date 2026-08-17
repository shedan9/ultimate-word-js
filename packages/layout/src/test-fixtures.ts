/**
 * 测试脚手架（不从 index.ts 导出，产库时也不会被打进去）。
 *
 * 度量器是**合成**的，不碰任何系统字体：布局测试要在 macOS / Linux / CI 上和 Windows 上
 * 给出同一个结果，读系统字体会让它变成「在我机器上是绿的」。真实字体的度量正确性由
 * `@uw/fonts` 的 metrics.test.ts（硬编码 Word 实测值）兜着，这里只测**排版算法**。
 *
 * 合成字体故意取最好算的形状：东亚字 1 em、ASCII 半角 0.5 em、win 跨度 1 em。
 * 于是「一行 28 个字」这种公文级的期望值可以手算出来，测试失败时一眼看得出差了几个字。
 */
import type { Twips } from '@uw/core';
import type { LineMetrics, TextMeasurer } from '@uw/fonts';
import { isEastAsianCodePoint } from '@uw/fonts';
import type {
  DocGrid,
  NumberLabel,
  ResolvedBlock,
  ResolvedCellProps,
  ResolvedParagraph,
  ResolvedParaProps,
  ResolvedRowProps,
  ResolvedRun,
  ResolvedRunProps,
  ResolvedTable,
  ResolvedTableCell,
  ResolvedTableProps,
  ResolvedTableRow,
  RunContent,
} from '@uw/model';

/** 五号字（10.5pt）= 210 twips，公文正文的默认字号 */
export const SIZE_5 = 210;

/** 合成度量器：东亚全角、其余半角；行高按 `@uw/fonts` 的两条规则同构地造 */
export function fakeMeasurer(): TextMeasurer {
  const advanceOf = (fontSize: Twips, cp: number): Twips =>
    isEastAsianCodePoint(cp) ? fontSize : fontSize / 2;
  return {
    status: () => 'metrics',
    lineMetrics(_family, fontSize, o = {}): LineMetrics {
      const ascent = fontSize * 0.8;
      const descent = fontSize * 0.2;
      // 东亚行的 1.3 倍系数（metrics.ts 里 13 个样本标定的那个），拉丁行没有外部行距。
      // 合成字体刻意让拉丁的外部行距为 0，于是 coreAbove 恒等于 ascent —— 期望值能手算，
      // 而「外部行距整块在基线以上」那条规则由 @uw/fonts 的单测用真字体的度量兜着。
      const lineGap = o.eastAsian === true ? (ascent + descent) * 0.3 : 0;
      return { ascent, descent, lineGap, lineHeight: ascent + descent + lineGap, coreAbove: ascent };
    },
    advances(_family, fontSize, codePoints, out, count) {
      const n = count ?? codePoints.length;
      for (let i = 0; i < n; i++) out[i] = advanceOf(fontSize, codePoints[i] as number);
    },
    advance: (_family, fontSize, cp) => advanceOf(fontSize, cp),
  };
}

export const NO_GRID: DocGrid = { type: 'default', linePitch: 0, charSpace: 0 };

export function runProps(over: Partial<ResolvedRunProps> = {}): ResolvedRunProps {
  return {
    fonts: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '仿宋', cs: '', hint: 'eastAsia' },
    bold: false,
    boldCs: false,
    italic: false,
    italicCs: false,
    caps: false,
    smallCaps: false,
    strike: false,
    doubleStrike: false,
    hidden: false,
    size: SIZE_5,
    sizeCs: SIZE_5,
    underline: 'none',
    color: 'auto',
    vertAlign: 'baseline',
    charSpacing: 0,
    scale: 100,
    position: 0,
    kerning: 0,
    snapToGrid: true,
    langEastAsia: 'zh-CN',
    ...over,
  };
}

export function paraProps(over: Partial<ResolvedParaProps> = {}): ResolvedParaProps {
  return {
    styleId: 'Normal',
    justification: 'left',
    indent: {
      left: 0,
      right: 0,
      firstLine: 0,
      hanging: 0,
      leftChars: 0,
      rightChars: 0,
      firstLineChars: 0,
      hangingChars: 0,
    },
    spacing: {
      before: 0,
      after: 0,
      beforeLines: 0,
      afterLines: 0,
      line: 240,
      lineRule: 'auto',
      beforeAutospacing: false,
      afterAutospacing: false,
    },
    tabs: [],
    keepNext: false,
    keepLines: false,
    pageBreakBefore: false,
    widowControl: true,
    outlineLevel: 9,
    numbering: { numId: 0, level: 0 },
    snapToGrid: true,
    autoSpaceDE: true,
    autoSpaceDN: true,
    overflowPunct: true,
    markRunProps: runProps(),
    ...over,
  };
}

/**
 * 一个算好的列表编号。真实来源是 `@uw/model` 的级联（`resolveParaProps` + 计数器），
 * 布局层只吃结果 —— 这里直接造结果，免得布局测试被样式表与 numbering.xml 绑住。
 */
export function numberLabel(text: string, over: Partial<NumberLabel> = {}): NumberLabel {
  return { text, value: 1, suffix: 'tab', justification: 'left', runProps: runProps(), ...over };
}

let seq = 0;

/** 一个纯文本 run */
export function run(text: string, over: Partial<ResolvedRunProps> = {}): ResolvedRun {
  return { kind: 'run', id: `r${seq++}`, props: runProps(over), content: [{ kind: 'text', text }] };
}

/** 内容片段任意的 run（制表位、换行、图片） */
export function runOf(content: RunContent[], over: Partial<ResolvedRunProps> = {}): ResolvedRun {
  return { kind: 'run', id: `r${seq++}`, props: runProps(over), content };
}

export function para(runs: ResolvedRun[], over: Partial<ResolvedParaProps> = {}): ResolvedParagraph {
  return { kind: 'paragraph', id: `p${seq++}`, props: paraProps(over), runs };
}

// ── 表格 ──────────────────────────────────────────────────────────────────────
// 同样直接造**级联后**的属性：表格样式与条件格式那套的正确性归 `@uw/model` 管
// （table-cascade.test.ts），布局测试只关心「给定这些属性，格子摆在哪儿」。

/** 默认边距取 Word 模板的左右各 108 twips，这样测出来的可用宽是真实文档的样子 */
export function cellProps(over: Partial<ResolvedCellProps> = {}): ResolvedCellProps {
  return {
    width: { value: 0, type: 'auto' },
    borders: {},
    shading: undefined,
    margins: {
      top: { value: 0, type: 'dxa' },
      left: { value: 108, type: 'dxa' },
      bottom: { value: 0, type: 'dxa' },
      right: { value: 108, type: 'dxa' },
    },
    verticalAlign: 'top',
    noWrap: false,
    fitText: false,
    textDirection: '',
    ...over,
  };
}

export function rowProps(over: Partial<ResolvedRowProps> = {}): ResolvedRowProps {
  return {
    height: { value: 0, rule: 'auto' },
    cantSplit: false,
    header: false,
    justification: undefined,
    cellSpacing: { value: 0, type: 'nil' },
    gridBefore: 0,
    gridAfter: 0,
    widthBefore: { value: 0, type: 'nil' },
    widthAfter: { value: 0, type: 'nil' },
    ...over,
  };
}

export function tableProps(over: Partial<ResolvedTableProps> = {}): ResolvedTableProps {
  return {
    styleId: '',
    width: { value: 0, type: 'auto' },
    justification: 'left',
    indent: { value: 0, type: 'dxa' },
    borders: {},
    shading: undefined,
    cellMargins: {
      top: { value: 0, type: 'dxa' },
      left: { value: 108, type: 'dxa' },
      bottom: { value: 0, type: 'dxa' },
      right: { value: 108, type: 'dxa' },
    },
    cellSpacing: { value: 0, type: 'nil' },
    layout: 'autofit',
    look: {
      firstRow: false,
      lastRow: false,
      firstColumn: false,
      lastColumn: false,
      noHBand: false,
      noVBand: false,
    },
    rowBandSize: 1,
    colBandSize: 1,
    ...over,
  };
}

export function cell(blocks: ResolvedBlock[], over: Partial<ResolvedTableCell> = {}): ResolvedTableCell {
  return {
    kind: 'cell',
    id: `tc${seq++}`,
    props: cellProps(),
    blocks,
    gridSpan: 1,
    vMerge: 'none',
    ...over,
  };
}

export function row(cells: ResolvedTableCell[], over: Partial<ResolvedTableRow> = {}): ResolvedTableRow {
  return { kind: 'row', id: `tr${seq++}`, props: rowProps(), cells, ...over };
}

export function table(
  grid: number[],
  rows: ResolvedTableRow[],
  over: Partial<ResolvedTable> = {},
): ResolvedTable {
  return { kind: 'table', id: `tbl${seq++}`, props: tableProps(), grid, rows, ...over };
}
