/**
 * 表格属性的解析与级联（Phase 4 的 model 侧）。
 *
 * `gongwen-01.docx` 里没有表格（只有一份默认表格样式），所以这里全是手写样本。
 * 语料库进了带表格的真实公文之后，这些期望值要拿真值复核一遍 —— 尤其是
 * **隔行带的序号**，它现在只有规范做依据（见 cascade-table.ts 文件头）。
 */
import { createDiagnosticSink, ptToTwips } from '@uw/core';
import { child, parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import type { CascadeContext } from './cascade.ts';
import type { CellPosition } from './cascade-table.ts';
import { conditionsAt, resolveCellProps, resolveRowProps, resolveTableProps } from './cascade-table.ts';
import type { Body, ResolvedTable } from './nodes.ts';
import { EMPTY_NUMBERING } from './numbering.ts';
import { parseBody } from './parse-body.ts';
import { parseCellProps, parseRowProps, parseTableProps } from './parse-table-props.ts';
import { resolveBody } from './resolve-body.ts';
import { DEFAULT_SETTINGS } from './settings.ts';
import { parseStyles } from './styles.ts';
import type { TableLook } from './table-props.ts';
import { EMPTY_THEME } from './theme.ts';

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const wrap = (name: string, xml: string) =>
  child(parseXml(`<w:root ${W_NS}><${name}>${xml}</${name}></w:root>`).root, name);
const tblPr = (xml: string) => parseTableProps(wrap('w:tblPr', xml));
const trPr = (xml: string) => parseRowProps(wrap('w:trPr', xml));
const tcPr = (xml: string) => parseCellProps(wrap('w:tcPr', xml));

function ctxFrom(stylesXml: string): CascadeContext {
  return {
    styles: parseStyles(parseXml(`<w:styles ${W_NS}>${stylesXml}</w:styles>`), createDiagnosticSink()),
    theme: EMPTY_THEME,
    settings: DEFAULT_SETTINGS,
    numbering: EMPTY_NUMBERING,
  };
}

function bodyOf(xml: string): Body {
  return parseBody(
    parseXml(`<w:document ${W_NS}><w:body>${xml}<w:sectPr/></w:body></w:document>`, 'document.xml'),
    createDiagnosticSink(),
  );
}

/** Word 模板里那份默认表格样式，单元格左右 108 twips 就写在这儿 */
const NORMAL_TABLE = `
  <w:style w:type="table" w:default="1" w:styleId="a1"><w:name w:val="Normal Table"/>
    <w:tblPr><w:tblInd w:w="0" w:type="dxa"/>
      <w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>
        <w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>
    </w:tblPr>
  </w:style>`;

const ALL_LOOK: TableLook = {
  firstRow: true,
  lastRow: true,
  firstColumn: true,
  lastColumn: true,
  noHBand: false,
  noVBand: false,
};

describe('解析', () => {
  it('宽度的 pct 是 1/50 个百分点 —— 5000 才是 100%', () => {
    expect(tblPr('<w:tblW w:w="5000" w:type="pct"/>').width).toEqual({ value: 5000, type: 'pct' });
    // w:type 缺席按 auto，**不是** dxa：auto 与「宽度 0」是两回事
    expect(tblPr('<w:tblW w:w="1440"/>').width).toEqual({ value: 1440, type: 'auto' });
  });

  it('边框 w:sz 是 1/8 磅、w:space 是磅 —— 同一个元素上两种刻度', () => {
    const b = tblPr(
      '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="1" w:color="000000"/></w:tblBorders>',
    );
    expect(b.borders?.top).toEqual({
      style: 'single',
      size: ptToTwips(0.5), // 4/8 磅 = 0.5pt = 10 twips
      space: ptToTwips(1),
      color: '000000',
    });
  });

  it('w:tblLook 的两种写法都认 —— 旧文档写的是十六进制位掩码', () => {
    const modern = tblPr('<w:tblLook w:firstRow="1" w:lastRow="0" w:noVBand="1"/>').look;
    expect(modern).toMatchObject({ firstRow: true, lastRow: false, noVBand: true });

    // 04A0 = 0x0020(firstRow) | 0x0080(firstColumn) | 0x0400(noVBand)
    const legacy = tblPr('<w:tblLook w:val="04A0"/>').look;
    expect(legacy).toEqual({
      firstRow: true,
      lastRow: false,
      firstColumn: true,
      lastColumn: false,
      noHBand: false,
      noVBand: true,
    });
  });

  it('w:trHeight 缺 hRule 时是 atLeast，不是 exact', () => {
    expect(trPr('<w:trHeight w:val="567"/>').height).toEqual({ value: 567, rule: 'atLeast' });
    expect(trPr('<w:trHeight w:val="567" w:hRule="exact"/>').height?.rule).toBe('exact');
  });

  it('gridSpan / vMerge 不进 tcPr 的属性对象 —— 它们是结构', () => {
    const props = tcPr('<w:gridSpan w:val="3"/><w:vMerge w:val="restart"/><w:tcW w:w="2000" w:type="dxa"/>');
    expect(props).toEqual({ width: { value: 2000, type: 'dxa' } });
  });

  it('w:tblGrid 的列数就是列数 —— 缺 w:w 的列补 0 而不是丢掉', () => {
    const body = bodyOf(
      `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol/><w:gridCol w:w="3000"/></w:tblGrid>
        <w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`,
    );
    const t = body.sections[0]?.blocks[0];
    expect(t?.kind).toBe('table');
    expect(t?.kind === 'table' ? t.grid : []).toEqual([2000, 0, 3000]);
  });
});

describe('表级级联', () => {
  it('没写 w:tblStyle 也吃默认表格样式 —— 108 的边距来自它，不是硬编码', () => {
    const ctx = ctxFrom(NORMAL_TABLE);
    const t = resolveTableProps(ctx, {});
    expect(t.styleId).toBe('a1');
    expect(t.cellMargins.left).toEqual({ value: 108, type: 'dxa' });
    expect(t.cellMargins.top).toEqual({ value: 0, type: 'dxa' });
  });

  it('边距逐边合并：直接格式只改左边，其余三边留着样式的', () => {
    const ctx = ctxFrom(NORMAL_TABLE);
    const t = resolveTableProps(ctx, tblPr('<w:tblCellMar><w:left w:w="360" w:type="dxa"/></w:tblCellMar>'));
    expect(t.cellMargins.left).toEqual({ value: 360, type: 'dxa' });
    expect(t.cellMargins.right).toEqual({ value: 108, type: 'dxa' });
  });

  it('basedOn 链上祖先先铺、后代覆盖', () => {
    const ctx = ctxFrom(`${NORMAL_TABLE}
      <w:style w:type="table" w:styleId="base"><w:name w:val="base"/>
        <w:tblPr><w:tblLayout w:type="fixed"/><w:tblStyleRowBandSize w:val="2"/></w:tblPr></w:style>
      <w:style w:type="table" w:styleId="derived"><w:name w:val="derived"/><w:basedOn w:val="base"/>
        <w:tblPr><w:tblStyleRowBandSize w:val="3"/></w:tblPr></w:style>`);
    const t = resolveTableProps(ctx, { styleId: 'derived' });
    expect(t.layout).toBe('fixed'); // 只有祖先写了
    expect(t.rowBandSize).toBe(3); // 后代赢
  });

  it('layout 缺席是 autofit —— 规范默认，不是我们挑的', () => {
    expect(resolveTableProps(ctxFrom(''), {}).layout).toBe('autofit');
  });
});

describe('条件格式的命中', () => {
  const at = (look: TableLook, pos: Partial<CellPosition>, bands = { row: 1, col: 1 }) =>
    conditionsAt(look, bands, { row: 0, rowCount: 3, col: 0, span: 1, colCount: 3, ...pos });

  it('左上角同时命中首行、首列与 nwCell，且按应用顺序排好', () => {
    // 顺序是「列 → 行 → 角」：firstRow 排在 firstCol 之后，所以表头行赢
    expect(at(ALL_LOOK, { row: 0, col: 0 })).toEqual(['firstCol', 'firstRow', 'nwCell']);
  });

  it('tblLook 关掉的那一项整个不命中 —— 它是开关不是格式', () => {
    const look: TableLook = { ...ALL_LOOK, firstRow: false, noVBand: true };
    // 首行不再是首行，于是它落回隔行带里去了
    expect(at(look, { row: 0, col: 1 })).toEqual(['band1Horz']);
  });

  it('跨列的格子盖到最后一列就算末列', () => {
    expect(at(ALL_LOOK, { row: 1, col: 1, span: 2, colCount: 3 })).toContain('lastCol');
  });

  it('首末行不计入隔行带 —— 否则「表头 + 隔行底纹」会整体错一行', () => {
    const look: TableLook = { ...ALL_LOOK, firstColumn: false, lastColumn: false, noVBand: true };
    const rows = (n: number) => at(look, { row: n, rowCount: 6, col: 1, colCount: 3 });
    expect(rows(0)).toEqual(['firstRow']); // 表头，不是带
    expect(rows(1)).toEqual(['band1Horz']); // 排除表头后的第 0 条带
    expect(rows(2)).toEqual(['band2Horz']);
    expect(rows(3)).toEqual(['band1Horz']);
    expect(rows(5)).toEqual(['lastRow']);
  });

  it('rowBandSize=2 时两行一条带', () => {
    const look: TableLook = { ...ALL_LOOK, firstRow: false, lastRow: false, noVBand: true };
    const rows = (n: number) =>
      conditionsAt(look, { row: 2, col: 1 }, { row: n, rowCount: 8, col: 1, span: 1, colCount: 3 });
    expect(rows(0)).toEqual(['band1Horz']);
    expect(rows(1)).toEqual(['band1Horz']);
    expect(rows(2)).toEqual(['band2Horz']);
    expect(rows(3)).toEqual(['band2Horz']);
    expect(rows(4)).toEqual(['band1Horz']);
  });
});

describe('格级级联与格内段落', () => {
  const STYLED = `${NORMAL_TABLE}
    <w:style w:type="table" w:styleId="grid"><w:name w:val="Table Grid"/>
      <w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>
        <w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>
      <w:tblStylePr w:type="firstRow">
        <w:rPr><w:b/><w:sz w:val="32"/></w:rPr>
        <w:tcPr><w:shd w:val="clear" w:fill="D9D9D9"/><w:vAlign w:val="center"/></w:tcPr>
      </w:tblStylePr>
      <w:tblStylePr w:type="firstCol"><w:rPr><w:i/><w:sz w:val="18"/></w:rPr></w:tblStylePr>
      <w:tblStylePr w:type="nwCell"><w:rPr><w:sz w:val="44"/></w:rPr></w:tblStylePr>
    </w:style>`;

  const table = (look: string, cells: string) =>
    bodyOf(`<w:tbl>
      <w:tblPr><w:tblStyle w:val="grid"/><w:tblLook ${look}/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
      ${cells}</w:tbl>`);

  const row = (...texts: string[]) =>
    `<w:tr>${texts.map((t) => `<w:tc><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`).join('')}</w:tr>`;

  function resolved(look: string, cells: string): ResolvedTable {
    const out = resolveBody(ctxFrom(STYLED), table(look, cells));
    const t = out.sections[0]?.blocks[0];
    if (t?.kind !== 'table') throw new Error('第一个块不是表格');
    return t;
  }

  const runProps = (t: ResolvedTable, r: number, c: number) => {
    const block = t.rows[r]?.cells[c]?.blocks[0];
    if (block?.kind !== 'paragraph') throw new Error('单元格里第一个块不是段落');
    return block.runs[0]?.props;
  };

  it('表头行的 rPr 穿透到格内 run 上', () => {
    const t = resolved('w:firstRow="1"', row('表头', 'B') + row('正文', 'D'));
    expect(runProps(t, 0, 0)?.bold).toBe(true);
    expect(runProps(t, 1, 0)?.bold).toBe(false);
  });

  it('tblLook 说不要表头行，那份格式就不应用', () => {
    const t = resolved('w:firstRow="0"', row('表头', 'B') + row('正文', 'D'));
    expect(runProps(t, 0, 0)?.bold).toBe(false);
  });

  it('角单元格最后裁决：nwCell 的字号压过 firstRow 与 firstCol', () => {
    const t = resolved('w:firstRow="1" w:firstColumn="1"', row('角', 'B') + row('C', 'D'));
    expect(runProps(t, 0, 0)?.size).toBe(440); // 22pt = w:sz 44 半磅 → twips
    // 同一行的非首列格子只吃 firstRow
    expect(runProps(t, 0, 1)?.size).toBe(320);
    // 同一列的非首行格子只吃 firstCol
    expect(runProps(t, 1, 0)?.size).toBe(180);
  });

  it('表格样式排在段落样式之前 —— 段落自己的样式能盖掉表头行的加粗', () => {
    const ctx = ctxFrom(`${STYLED}
      <w:style w:type="paragraph" w:styleId="plain"><w:name w:val="plain"/>
        <w:rPr><w:b w:val="0"/></w:rPr></w:style>`);
    const body = table(
      'w:firstRow="1"',
      `<w:tr><w:tc><w:p><w:pPr><w:pStyle w:val="plain"/></w:pPr>
        <w:r><w:t>表头</w:t></w:r></w:p></w:tc></w:tr>`,
    );
    const t = resolveBody(ctx, body).sections[0]?.blocks[0];
    if (t?.kind !== 'table') throw new Error('第一个块不是表格');
    const p = t.rows[0]?.cells[0]?.blocks[0];
    expect(p?.kind === 'paragraph' ? p.runs[0]?.props.bold : undefined).toBe(false);
  });

  it('条件格式的 tcPr 落到单元格属性上', () => {
    const t = resolved('w:firstRow="1"', row('表头') + row('正文'));
    expect(t.rows[0]?.cells[0]?.props.shading?.fill).toBe('D9D9D9');
    expect(t.rows[0]?.cells[0]?.props.verticalAlign).toBe('center');
    expect(t.rows[1]?.cells[0]?.props.verticalAlign).toBe('top');
  });

  it('w:tcMar 缺席时退到表级 w:tblCellMar，不是退到 0', () => {
    const t = resolved('', row('A'));
    expect(t.rows[0]?.cells[0]?.props.margins.left).toEqual({ value: 108, type: 'dxa' });
  });

  it('列号按 gridSpan 累加 —— 跨列格子之后的列号不是「第几个 tc」', () => {
    const ctx = ctxFrom(STYLED);
    const t = resolveTableProps(ctx, {
      styleId: 'grid',
      look: { ...ALL_LOOK, noHBand: true, noVBand: true },
    });
    const pos = (col: number, span: number): CellPosition => ({
      row: 1,
      rowCount: 3,
      col,
      span,
      colCount: 4,
    });
    // 第一格跨 3 列，第二格的起始列是 3 —— 于是它才是末列
    expect(resolveCellProps(ctx, t, {}, {}, pos(0, 3)).layers.length).toBeGreaterThan(0);
    expect(conditionsAt(t.look, { row: 1, col: 1 }, pos(3, 1))).toContain('lastCol');
    expect(conditionsAt(t.look, { row: 1, col: 1 }, pos(1, 1))).not.toContain('lastCol');
  });

  it('w:gridBefore 把本行的列号整体后移', () => {
    const t = resolved(
      'w:firstColumn="1"',
      `<w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr>
        <w:tc><w:p><w:r><w:t>右边那格</w:t></w:r></w:p></w:tc></w:tr>`,
    );
    expect(t.rows[0]?.props.gridBefore).toBe(1);
    // 跳过了第 0 列，所以这一格不是首列，不该吃 firstCol 的斜体
    expect(runProps(t, 0, 0)?.italic).toBe(false);
  });

  it('行属性只吃与行有关的条件格式', () => {
    const ctx = ctxFrom(`${NORMAL_TABLE}
      <w:style w:type="table" w:styleId="hdr"><w:name w:val="hdr"/>
        <w:tblStylePr w:type="firstRow"><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr></w:tblStylePr>
        <w:tblStylePr w:type="firstCol"><w:trPr><w:trHeight w:val="9999"/></w:trPr></w:tblStylePr>
      </w:style>`);
    const direct = { styleId: 'hdr', look: ALL_LOOK };
    const t = resolveTableProps(ctx, direct);
    const first = resolveRowProps(ctx, t, direct, {}, { row: 0, rowCount: 3 });
    expect(first.header).toBe(true);
    expect(first.cantSplit).toBe(true);
    // firstCol 是列上的条件，行级不该命中它 —— 「col 填 0」曾让它无条件命中
    expect(first.height).toEqual({ value: 0, rule: 'auto' });
    expect(resolveRowProps(ctx, t, direct, {}, { row: 1, rowCount: 3 }).header).toBe(false);
  });

  it('嵌套表格用内层自己的表格样式，不叠加外层', () => {
    const ctx = ctxFrom(`${STYLED}
      <w:style w:type="table" w:styleId="inner"><w:name w:val="inner"/>
        <w:tblStylePr w:type="firstRow"><w:rPr><w:i/></w:rPr></w:tblStylePr></w:style>`);
    const body = bodyOf(`<w:tbl>
      <w:tblPr><w:tblStyle w:val="grid"/><w:tblLook w:firstRow="1"/></w:tblPr>
      <w:tr><w:tc>
        <w:tbl><w:tblPr><w:tblStyle w:val="inner"/><w:tblLook w:firstRow="1"/></w:tblPr>
          <w:tr><w:tc><w:p><w:r><w:t>内层</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
      </w:tc></w:tr></w:tbl>`);
    const outer = resolveBody(ctx, body).sections[0]?.blocks[0];
    if (outer?.kind !== 'table') throw new Error('第一个块不是表格');
    const inner = outer.rows[0]?.cells[0]?.blocks[0];
    if (inner?.kind !== 'table') throw new Error('内层不是表格');
    const p = inner.rows[0]?.cells[0]?.blocks[0];
    const props = p?.kind === 'paragraph' ? p.runs[0]?.props : undefined;
    expect(props?.italic).toBe(true); // 内层 firstRow 的斜体
    expect(props?.bold).toBe(false); // 外层 firstRow 的加粗没有跟进来
  });
});

describe('w:tblPrEx（行级表格属性例外）', () => {
  /** 表级 insideH 半磅、左边距 108；第二行用例外改成 3 磅 + 400 */
  const TBL = `<w:tbl>
    <w:tblPr>
      <w:tblBorders><w:insideH w:val="single" w:sz="4"/></w:tblBorders>
      <w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>
    </w:tblPr>
    <w:tr><w:tc><w:p/></w:tc></w:tr>
    <w:tr>
      <w:tblPrEx>
        <w:tblBorders><w:insideH w:val="single" w:sz="24"/></w:tblBorders>
        <w:tblCellMar><w:left w:w="400" w:type="dxa"/></w:tblCellMar>
      </w:tblPrEx>
      <w:tc><w:p/></w:tc>
    </w:tr>
  </w:tbl>`;

  function rows(): ResolvedTable['rows'] {
    const t = resolveBody(ctxFrom(NORMAL_TABLE), bodyOf(TBL)).sections[0]?.blocks[0];
    if (t?.kind !== 'table') throw new Error('第一个块不是表格');
    return t.rows;
  }

  it('例外只作用于本行，别的行照旧用整表那一份', () => {
    const [first, second] = rows();
    expect(first?.cells[0]?.props.margins.left).toEqual({ value: 108, type: 'dxa' });
    expect(second?.cells[0]?.props.margins.left).toEqual({ value: 400, type: 'dxa' });
  });

  it('例外里没写的项逐边留着 —— 只改了 left，right 仍是整表的 108', () => {
    expect(rows()[1]?.cells[0]?.props.margins.right).toEqual({ value: 108, type: 'dxa' });
  });

  it('例外改过的表级边框带到布局层去 —— 冲突解析的「退到表级」对这一行说的是它', () => {
    const [first, second] = rows();
    expect(first?.props.tableBorders).toBeUndefined(); // 缺席 = 用整表那一份
    expect(second?.props.tableBorders?.insideH?.size).toBe(ptToTwips(3)); // 24/8 磅
  });

  it('不再报 unknown-element —— 漏解析时这一行的格线会沿用整表的宽度', () => {
    const sink = createDiagnosticSink();
    parseBody(
      parseXml(`<w:document ${W_NS}><w:body>${TBL}<w:sectPr/></w:body></w:document>`, 'document.xml'),
      sink,
    );
    expect(sink.list()).toEqual([]);
  });
});
