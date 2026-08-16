/**
 * `w:tblPr` / `w:tblGrid` / `w:trPr` / `w:tcPr` → 部分属性对象。
 *
 * 与 parse-props.ts 同一条纪律：**只翻译，不合并、不填默认值**。缺席就是缺席 ——
 * 表格的级联层数比段落还多（表格样式 → 条件格式 → 表 → 行 → 格），
 * 在这里填一个默认值，等于让最外层凭空压过里面每一层。
 */
import { ptToTwips } from '@uw/core';
import type { XmlElement } from '@uw/ooxml';
import { attr, child, children } from '@uw/ooxml';
import { parseJustification } from './parse-props.ts';
import type {
  Border,
  CellBorders,
  CellMargins,
  CellProps,
  RowHeight,
  RowProps,
  Shading,
  TableBorders,
  TableLook,
  TableProps,
  TableWidth,
} from './table-props.ts';
import { attrInt, attrOf, attrOnOff, enumVal, intVal, onOff, put, valOf } from './xml-values.ts';

const WIDTH_TYPES = ['auto', 'dxa', 'pct', 'nil'] as const;
const H_RULES = ['auto', 'atLeast', 'exact'] as const;
const V_ALIGNS = ['top', 'center', 'bottom'] as const;
const TABLE_BORDER_SIDES = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'] as const;
const MARGIN_SIDES = ['top', 'left', 'bottom', 'right'] as const;

/** `w:sz` 在边框里的刻度是 **1/8 磅**（`w:sz="4"` 就是最常见的 0.5pt 细线） */
const eighthPtToTwips = (eighth: number): number => ptToTwips(eighth / 8);

/**
 * `<w:tblW w:w="5000" w:type="pct"/>` 这类宽度。
 *
 * `w:w` 缺席按 0 处理，但 `w:type` 缺席**不能**当 `dxa`：规范默认是 `auto`，
 * 而 `auto` 与 `dxa 0`（宽度零）是天差地别的两件事。
 */
function parseWidth(el: XmlElement | undefined, attrName = 'w:w'): TableWidth | undefined {
  if (el === undefined) return undefined;
  const value = attrInt(el, attrName);
  const type = enumVal(attr(el, 'w:type'), WIDTH_TYPES);
  if (value === undefined && type === undefined) return undefined;
  return { value: value ?? 0, type: type ?? 'auto' };
}

function parseBorder(el: XmlElement | undefined): Border | undefined {
  if (el === undefined) return undefined;
  const size = attrInt(el, 'w:sz');
  const space = attrInt(el, 'w:space');
  const out: Border = {
    // 缺 w:val 的边框元素在实文件里见过，按「未指定」处理而不是丢掉整条
    style: attrOf(el, 'w:val') ?? 'none',
    size: size === undefined ? 0 : eighthPtToTwips(size),
    // w:space 的单位是磅，不是 1/8 磅 —— 同一个元素上两种刻度，抄错一处就偏
    space: space === undefined ? 0 : ptToTwips(space),
    color: attrOf(el, 'w:color') ?? 'auto',
  };
  const shadow = attrOnOff(el, 'w:shadow');
  if (shadow !== undefined) out.shadow = shadow;
  return out;
}

function parseTableBorders(el: XmlElement | undefined): TableBorders | undefined {
  if (el === undefined) return undefined;
  const out: TableBorders = {};
  for (const side of TABLE_BORDER_SIDES) {
    // w:start / w:end 是 w:left / w:right 的新名字，与 w:ind 那边同理
    const alias = side === 'left' ? 'w:start' : side === 'right' ? 'w:end' : undefined;
    const found = child(el, `w:${side}`) ?? (alias === undefined ? undefined : child(el, alias));
    put(out, side, parseBorder(found));
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

function parseCellBorders(el: XmlElement | undefined): CellBorders | undefined {
  const base = parseTableBorders(el);
  if (el === undefined) return undefined;
  const out: CellBorders = base ?? {};
  put(out, 'tl2br', parseBorder(child(el, 'w:tl2br')));
  put(out, 'tr2bl', parseBorder(child(el, 'w:tr2bl')));
  return Object.keys(out).length > 0 ? out : undefined;
}

function parseShading(el: XmlElement | undefined): Shading | undefined {
  if (el === undefined) return undefined;
  return {
    pattern: attrOf(el, 'w:val') ?? 'clear',
    color: attrOf(el, 'w:color') ?? 'auto',
    // fill 才是背景色。缺席时是 auto（跟随主题），不是白色
    fill: attrOf(el, 'w:fill') ?? 'auto',
  };
}

function parseCellMargins(el: XmlElement | undefined): CellMargins | undefined {
  if (el === undefined) return undefined;
  const out: CellMargins = {};
  for (const side of MARGIN_SIDES) {
    const alias = side === 'left' ? 'w:start' : side === 'right' ? 'w:end' : undefined;
    const found = child(el, `w:${side}`) ?? (alias === undefined ? undefined : child(el, alias));
    put(out, side, parseWidth(found));
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

/**
 * `w:tblLook`。
 *
 * 两种写法都在真实文件里出现：新的是六个独立属性，旧的（Word 2007）是
 * `w:val="04A0"` 一个**十六进制位掩码**。只认新的，那些 2007 存的文档会丢掉整套
 * 条件格式开关 —— 表现为「表格样式的表头行格式莫名其妙全没了」。
 */
const LOOK_BITS: Record<keyof TableLook, number> = {
  firstRow: 0x0020,
  lastRow: 0x0040,
  firstColumn: 0x0080,
  lastColumn: 0x0100,
  noHBand: 0x0200,
  noVBand: 0x0400,
};

function parseTableLook(el: XmlElement | undefined): TableLook | undefined {
  if (el === undefined) return undefined;
  // w:val 是十六进制字符串（"04A0"），按十进制读会得到一组毫不相干的位
  const raw = attr(el, 'w:val');
  const bits = raw === undefined ? Number.NaN : Number.parseInt(raw, 16);
  const of = (name: string, key: keyof TableLook): boolean =>
    attrOnOff(el, name) ?? (Number.isNaN(bits) ? false : (bits & LOOK_BITS[key]) !== 0);

  return {
    firstRow: of('w:firstRow', 'firstRow'),
    lastRow: of('w:lastRow', 'lastRow'),
    firstColumn: of('w:firstColumn', 'firstColumn'),
    lastColumn: of('w:lastColumn', 'lastColumn'),
    noHBand: of('w:noHBand', 'noHBand'),
    noVBand: of('w:noVBand', 'noVBand'),
  };
}

export function parseTableProps(tblPr: XmlElement | undefined): TableProps {
  const out: TableProps = {};
  if (tblPr === undefined) return out;

  put(out, 'styleId', valOf(tblPr, 'w:tblStyle'));
  put(out, 'width', parseWidth(child(tblPr, 'w:tblW')));
  put(out, 'justification', parseJustification(valOf(tblPr, 'w:jc')));
  put(out, 'indent', parseWidth(child(tblPr, 'w:tblInd')));
  put(out, 'borders', parseTableBorders(child(tblPr, 'w:tblBorders')));
  put(out, 'shading', parseShading(child(tblPr, 'w:shd')));
  put(out, 'cellMargins', parseCellMargins(child(tblPr, 'w:tblCellMar')));
  put(out, 'cellSpacing', parseWidth(child(tblPr, 'w:tblCellSpacing')));
  put(out, 'layout', enumVal(attrOf(child(tblPr, 'w:tblLayout'), 'w:type'), ['fixed', 'autofit']));
  put(out, 'look', parseTableLook(child(tblPr, 'w:tblLook')));
  put(out, 'rowBandSize', intVal(tblPr, 'w:tblStyleRowBandSize'));
  put(out, 'colBandSize', intVal(tblPr, 'w:tblStyleColBandSize'));

  return out;
}

/**
 * `w:tblGrid` → 每列的基准宽度（twips）。
 *
 * `w:gridCol` 缺 `w:w` 时补 0 而不是跳过：**列的个数**本身是有意义的
 * （`gridSpan` 按列数算），少一列会让后面所有格子的列号错位。
 */
export function parseTableGrid(tblGrid: XmlElement | undefined): number[] {
  if (tblGrid === undefined) return [];
  return [...children(tblGrid, 'w:gridCol')].map((c) => attrInt(c, 'w:w') ?? 0);
}

export function parseRowProps(trPr: XmlElement | undefined): RowProps {
  const out: RowProps = {};
  if (trPr === undefined) return out;

  const h = child(trPr, 'w:trHeight');
  if (h !== undefined) {
    const value = attrInt(h, 'w:val');
    const height: RowHeight = {
      value: value ?? 0,
      // 缺席是 atLeast（最小值），认成 exact 会把内容压扁
      rule: enumVal(attr(h, 'w:hRule'), H_RULES) ?? 'atLeast',
    };
    out.height = height;
  }
  put(out, 'cantSplit', onOff(trPr, 'w:cantSplit'));
  put(out, 'header', onOff(trPr, 'w:tblHeader'));
  put(out, 'justification', parseJustification(valOf(trPr, 'w:jc')));
  put(out, 'cellSpacing', parseWidth(child(trPr, 'w:tblCellSpacing')));
  put(out, 'gridBefore', intVal(trPr, 'w:gridBefore'));
  put(out, 'gridAfter', intVal(trPr, 'w:gridAfter'));
  put(out, 'widthBefore', parseWidth(child(trPr, 'w:wBefore')));
  put(out, 'widthAfter', parseWidth(child(trPr, 'w:wAfter')));

  return out;
}

/**
 * `w:tcPr` → 单元格属性。
 *
 * `w:gridSpan` / `w:vMerge` **不在这里**：它们是结构而不是格式（决定这个格子占几列、
 * 内容画不画），存在节点上（`nodes.ts` 的 `TableCellNode`），级联时才并进结果。
 * 混在一起的话，表格样式的条件格式就能把「合并了几列」给覆盖掉。
 */
export function parseCellProps(tcPr: XmlElement | undefined): CellProps {
  const out: CellProps = {};
  if (tcPr === undefined) return out;

  put(out, 'width', parseWidth(child(tcPr, 'w:tcW')));
  put(out, 'borders', parseCellBorders(child(tcPr, 'w:tcBorders')));
  put(out, 'shading', parseShading(child(tcPr, 'w:shd')));
  put(out, 'margins', parseCellMargins(child(tcPr, 'w:tcMar')));
  put(out, 'verticalAlign', enumVal(valOf(tcPr, 'w:vAlign'), V_ALIGNS));
  put(out, 'noWrap', onOff(tcPr, 'w:noWrap'));
  put(out, 'fitText', onOff(tcPr, 'w:tcFitText'));
  put(out, 'textDirection', valOf(tcPr, 'w:textDirection'));

  return out;
}
