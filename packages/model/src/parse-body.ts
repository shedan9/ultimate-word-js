/**
 * `document.xml` → 正文节点树。
 *
 * 三条纪律：
 *
 * 1. **不认识的东西要出声。** 未知元素记一条 `Diagnostic` 然后跳过（原则 1.5），
 *    绝不静默丢内容 —— 静默丢的后果是「文档少了一段」这种最难查的 bug。
 *    同名元素只报一次，否则一份带修订的文档能刷出上千条诊断
 * 2. **容器一律压平。** `w:hyperlink` / `w:ins` / `w:sdt` / `w:smartTag` / `w:fldSimple`
 *    在排版上都是透明的，只是包着 run 而已。压平之后段落的子节点是一列扁平的 run，
 *    断行算法不必递归下钻（见 nodes.ts 的 `RunNode.hyperlink`）
 * 3. **这里不做级联。** 产出的是直接格式（可缺席），级联在 resolve-body.ts
 */
import type { DiagnosticSink } from '@uw/core';
import type { XmlDocument, XmlElement } from '@uw/ooxml';
import { attr, child, children, textContent } from '@uw/ooxml';
import type {
  Block,
  Body,
  NodeId,
  Paragraph,
  Run,
  RunContent,
  Section,
  Table,
  TableCell,
  TableRow,
} from './nodes.ts';
import { parseDrawing, parsePict } from './parse-drawing.ts';
import { parseParaProps, parseRunProps } from './parse-props.ts';
import { parseCellProps, parseRowProps, parseTableGrid, parseTableProps } from './parse-table-props.ts';
import type { ParaProps } from './props.ts';
import { parseSectionProps } from './section.ts';
import { attrOf, enumVal } from './xml-values.ts';

/**
 * 排版上完全无关、见到就跳过的元素 —— 书签、拼写检查标记、批注范围、编辑权限范围。
 *
 * 它们**不该**进诊断：每份 Word 文档都有一堆，报出来只会淹掉真正的未知元素。
 */
const IGNORED = new Set([
  'w:bookmarkStart',
  'w:bookmarkEnd',
  'w:commentRangeStart',
  'w:commentRangeEnd',
  'w:permStart',
  'w:permEnd',
  'w:proofErr',
  'w:customXmlInsRangeStart',
  'w:customXmlInsRangeEnd',
  'w:customXmlDelRangeStart',
  'w:customXmlDelRangeEnd',
  'w:moveFromRangeStart',
  'w:moveFromRangeEnd',
  'w:moveToRangeStart',
  'w:moveToRangeEnd',
]);

/** run 内部同样跳过的：脚注/尾注/批注的引用标记（Phase 3 之后才接），以及 Word 缓存的分页提示 */
const IGNORED_IN_RUN = new Set([
  'w:rPr',
  'w:footnoteReference',
  'w:endnoteReference',
  'w:commentReference',
  'w:annotationRef',
  'w:footnoteRef',
  'w:endnoteRef',
  // Word 自己排完版后写回来的「这里分了页」。**绝不能采信** ——
  // 采信它等于让 Word 替我们排版，而这个引擎的立身之本就是自己算
  'w:lastRenderedPageBreak',
  'w:ptab',
]);

/** 透明容器：内容照收，容器本身不产生节点 */
const TRANSPARENT = new Set(['w:smartTag', 'w:bdo', 'w:dir', 'w:ins', 'w:moveTo']);

/** 删除的内容：`w:del` / `w:moveFrom` 里的文字是「已删除」，最终版式里不出现 */
const DELETED = new Set(['w:del', 'w:moveFrom']);

interface Ctx {
  diagnostics: DiagnosticSink;
  part: string;
  /** 已经报过的元素名，去重用 */
  reported: Set<string>;
  counters: Map<string, number>;
  /**
   * id 前缀。主文档是空串，页眉页脚的每个部件各带一个 —— **不带前缀就会撞车**：
   * 计数器是按部件新建的，页眉里的第一个 run 与正文里的第一个 run 都会叫 `r0`，
   * 而域求值那张「run id → 显示的文字」的表是全文档一张，撞了就会把页脚的页码
   * 画到正文里去。
   */
  idPrefix: string;
}

function nextId(ctx: Ctx, prefix: string): NodeId {
  const n = ctx.counters.get(prefix) ?? 0;
  ctx.counters.set(prefix, n + 1);
  return `${ctx.idPrefix}${prefix}${n}`;
}

function unknown(ctx: Ctx, el: XmlElement, where: string): void {
  const key = `${where}/${el.name}`;
  if (ctx.reported.has(key)) return;
  ctx.reported.add(key);
  ctx.diagnostics.warn('unknown-element', `${where} 里遇到不认识的元素 <${el.name}>，已跳过`, {
    part: ctx.part,
    path: el.name,
  });
}

// ── 入口 ──────────────────────────────────────────────────────────────────────

export function parseBody(doc: XmlDocument, diagnostics: DiagnosticSink, part = 'document.xml'): Body {
  const ctx: Ctx = { diagnostics, part, reported: new Set(), counters: new Map(), idPrefix: '' };
  const body = child(doc.root, 'w:body');
  if (body === undefined) {
    // 结构性问题，但不抛 —— `OpcPackage.requirePart` 已经保证部件在，
    // 走到这里只是文件内容怪，画一个空文档比白屏强（原则 1.5）
    diagnostics.warn('missing-body', '<w:document> 里没有 <w:body>，按空文档处理', { part });
    return { sections: [] };
  }

  const sections: Section[] = [];
  let pending: Block[] = [];

  for (const el of children(body)) {
    switch (el.name) {
      case 'w:p': {
        const props = parseParaProps(child(el, 'w:pPr'));
        pending.push(parseParagraph(ctx, el, props));
        // 段落属性里的 sectPr 表示「这一节到此为止」，**这个段落属于本节**
        const sectPr = sectPrOf(el);
        if (sectPr !== undefined) {
          sections.push({ id: nextId(ctx, 'sec'), props: parseSectionProps(sectPr), blocks: pending });
          pending = [];
        }
        break;
      }
      case 'w:tbl':
        pending.push(parseTable(ctx, el));
        break;
      case 'w:sdt':
        pending.push(...sdtBlocks(ctx, el));
        break;
      case 'w:sectPr':
        // body 末尾这个是**最后一节**的属性
        sections.push({ id: nextId(ctx, 'sec'), props: parseSectionProps(el), blocks: pending });
        pending = [];
        break;
      default:
        if (!IGNORED.has(el.name)) unknown(ctx, el, 'w:body');
    }
  }

  // 没有 body 级 sectPr（不合规，但见过）：剩下的块也得有节可归
  if (pending.length > 0 || sections.length === 0) {
    diagnostics.warn('missing-sectPr', '文档末尾没有 <w:sectPr>，页面尺寸用兜底值', { part });
    sections.push({ id: nextId(ctx, 'sec'), props: parseSectionProps(undefined), blocks: pending });
  }
  return { sections };
}

/**
 * `header*.xml` / `footer*.xml` → 一列块。
 *
 * 与 `parseBody` 的差别只有两处，但都藏得深：
 *
 * 1. 根元素是 `w:hdr` / `w:ftr`，块**直接**挂在根下，中间没有 `w:body`；
 *    也没有 `w:sectPr` —— 页眉不分节，它的版心是引用它的那一节给的
 * 2. `idPrefix` 必须给一个与主文档不同的值（见 `Ctx.idPrefix`）。
 *    传关系 id 是最省事的一个选择：它在主文档的关系表里唯一，且看 id 就知道来自哪个部件
 */
export function parseHeaderFooter(
  doc: XmlDocument,
  diagnostics: DiagnosticSink,
  part: string,
  idPrefix: string,
): Block[] {
  const ctx: Ctx = { diagnostics, part, reported: new Set(), counters: new Map(), idPrefix };
  const out: Block[] = [];
  for (const el of children(doc.root)) {
    switch (el.name) {
      case 'w:p':
        out.push(parseParagraph(ctx, el, parseParaProps(child(el, 'w:pPr'))));
        break;
      case 'w:tbl':
        out.push(parseTable(ctx, el));
        break;
      case 'w:sdt':
        out.push(...sdtBlocks(ctx, el));
        break;
      default:
        if (!IGNORED.has(el.name)) unknown(ctx, el, doc.root.name);
    }
  }
  return out;
}

function sectPrOf(p: XmlElement): XmlElement | undefined {
  const pPr = child(p, 'w:pPr');
  return pPr === undefined ? undefined : child(pPr, 'w:sectPr');
}

/** `w:sdt`（内容控件）在块级也是透明的，内容在 `w:sdtContent` 里 */
function sdtBlocks(ctx: Ctx, sdt: XmlElement): Block[] {
  const content = child(sdt, 'w:sdtContent');
  if (content === undefined) return [];
  const out: Block[] = [];
  for (const el of children(content)) {
    if (el.name === 'w:p') out.push(parseParagraph(ctx, el, parseParaProps(child(el, 'w:pPr'))));
    else if (el.name === 'w:tbl') out.push(parseTable(ctx, el));
    else if (el.name === 'w:sdt') out.push(...sdtBlocks(ctx, el));
    else if (!IGNORED.has(el.name)) unknown(ctx, el, 'w:sdtContent');
  }
  return out;
}

// ── 段落 ──────────────────────────────────────────────────────────────────────

function parseParagraph(ctx: Ctx, p: XmlElement, props: ParaProps): Paragraph {
  const runs: Run[] = [];
  collectRuns(ctx, p, runs, {});
  return { kind: 'paragraph', id: nextId(ctx, 'p'), props, runs };
}

/**
 * 从外层容器往下带的标记。
 *
 * `w:hyperlink` 与 `w:fldSimple` 都是**包着 run 的容器**，压平后信息只能挂在 run 上。
 * 合成一个对象而不是排开成两个参数：这类容器还会再加（`w:customXml` 之类），
 * 每加一种就改遍递归签名的写法撑不住。
 */
interface RunMarks {
  link?: Run['hyperlink'];
  field?: Run['fieldSimple'];
}

/**
 * 把 `parent` 底下所有 run 收进 `out`，沿途拆掉透明容器。
 *
 * `marks` 是从外层容器带下来的，会盖到每个 run 上 —— 嵌套超链接在 Word 里不合法，
 * 所以内层直接覆盖外层，不必合并。
 */
function collectRuns(ctx: Ctx, parent: XmlElement, out: Run[], marks: RunMarks): void {
  for (const el of children(parent)) {
    if (el.name === 'w:r') {
      out.push(parseRun(ctx, el, marks));
    } else if (el.name === 'w:hyperlink') {
      const next: NonNullable<Run['hyperlink']> = {};
      const relId = attr(el, 'r:id');
      const anchor = attr(el, 'w:anchor');
      if (relId !== undefined) next.relId = relId;
      if (anchor !== undefined) next.anchor = anchor;
      collectRuns(ctx, el, out, { ...marks, link: next });
    } else if (el.name === 'w:fldSimple') {
      // 简单域：整个域压缩成一个元素，域代码在 `w:instr` 属性里、结果就是它的子 run。
      // 内容照旧压平（结果文字与普通文字排版上毫无区别），域代码挂在 run 上给 fields.ts 收。
      // **必须给个 id**：相邻两个 `w:fldSimple w:instr="PAGE"` 是两个域，只比指令文字会并成一个
      const field = { id: nextId(ctx, 'fld'), instr: attr(el, 'w:instr') ?? '' };
      collectRuns(ctx, el, out, { ...marks, field });
    } else if (TRANSPARENT.has(el.name)) {
      collectRuns(ctx, el, out, marks);
    } else if (el.name === 'w:sdt') {
      const content = child(el, 'w:sdtContent');
      if (content !== undefined) collectRuns(ctx, content, out, marks);
    } else if (DELETED.has(el.name)) {
      // 修订只做显示、不做编辑（非目标），显示的是**接受后**的版式：删掉的字不占位。
      // 记一条 info 是因为「文档里有字没画出来」必须留痕，否则查起来无从下手
      ctx.diagnostics.info('revision-deleted', `跳过了 <${el.name}> 里被删除的内容`, { part: ctx.part });
    } else if (el.name === 'w:pPr') {
      // 段落属性已在别处解析
    } else if (!IGNORED.has(el.name)) {
      unknown(ctx, el, parent.name);
    }
  }
}

function parseRun(ctx: Ctx, r: XmlElement, marks: RunMarks): Run {
  const props: Run['props'] = parseRunProps(child(r, 'w:rPr'));
  const content: RunContent[] = [];
  collectRunContent(ctx, r, content);
  const run: Run = { kind: 'run', id: nextId(ctx, 'r'), props, content };
  if (marks.link !== undefined) run.hyperlink = marks.link;
  if (marks.field !== undefined) run.fieldSimple = marks.field;
  return run;
}

const BREAK_TYPES = ['page', 'column'] as const;

function collectRunContent(ctx: Ctx, parent: XmlElement, out: RunContent[]): void {
  for (const el of children(parent)) {
    switch (el.name) {
      case 'w:t':
        out.push({ kind: 'text', text: xmlText(el) });
        break;
      case 'w:delText':
        break; // 同 DELETED：已删除的字不占位，外层已经记过诊断
      case 'w:instrText':
        // 域代码**不去首尾空白**（与 `w:t` 相反）：这段文字不显示，空白是词与词的分隔符。
        // 一条指令常被切成好几段，去掉空白再拼就成了 ` IF ` + ` = 1 ` → `IF= 1`，
        // 域名当场变成 `IF=`。Word 自己总写 xml:space="preserve"，
        // 这一条挡的是第三方生成器（见 fields.ts 的「切碎的 instrText」用例）
        out.push({ kind: 'fieldInstruction', text: textContent(el) });
        break;
      case 'w:tab':
        out.push({ kind: 'tab' });
        break;
      case 'w:br':
        // w:type 缺省是 textWrapping，也就是普通换行
        out.push({ kind: 'break', breakType: enumVal(attr(el, 'w:type'), BREAK_TYPES) ?? 'line' });
        break;
      case 'w:cr':
        out.push({ kind: 'break', breakType: 'line' });
        break;
      case 'w:noBreakHyphen':
        out.push({ kind: 'noBreakHyphen' });
        break;
      case 'w:softHyphen':
        out.push({ kind: 'softHyphen' });
        break;
      case 'w:sym':
        out.push(parseSymbol(el));
        break;
      case 'w:fldChar': {
        const t = enumVal(attr(el, 'w:fldCharType'), ['begin', 'separate', 'end'] as const);
        if (t !== undefined) out.push({ kind: 'fieldChar', charType: t });
        break;
      }
      case 'w:drawing':
        out.push(parseDrawing(el, ctx.idPrefix));
        break;
      case 'w:pict':
        out.push(parsePict(el, ctx.idPrefix));
        break;
      case 'w:object':
        // OLE 对象（嵌入的 Excel 表、公式编辑器）在文件里也是一个 VML 形状加一张预览图，
        // 走同一条路 —— 我们画的就是那张预览图，与 Word 不激活时显示的东西一致
        out.push({ ...parsePict(el, ctx.idPrefix), objectKind: 'object' });
        break;
      case 'mc:AlternateContent': {
        // Word 2010 之后把图形包在这里：Choice 是新格式，Fallback 是给老版本的 VML。
        // 优先 Choice，两个都收会画出两份
        const pick = child(el, 'mc:Choice') ?? child(el, 'mc:Fallback');
        if (pick !== undefined) collectRunContent(ctx, pick, out);
        break;
      }
      default:
        if (!IGNORED_IN_RUN.has(el.name) && !IGNORED.has(el.name)) unknown(ctx, el, 'w:r');
    }
  }
}

/**
 * `w:t` / `w:instrText` 的文字。
 *
 * 没有 `xml:space="preserve"` 时，**首尾空白按 XML 规矩要去掉**。Word 自己在有首尾空白时
 * 总会写 preserve，所以这条只对第三方生成的文件起作用 —— 但不做的话那些文件每段都会多出空格。
 */
function xmlText(el: XmlElement): string {
  const raw = textContent(el);
  return attr(el, 'xml:space') === 'preserve' ? raw : raw.trim();
}

function parseSymbol(el: XmlElement): RunContent {
  const font = attr(el, 'w:font') ?? '';
  const hex = attr(el, 'w:char') ?? '';
  const code = Number.parseInt(hex, 16);
  // 符号字体的码位通常落在 F000–F0FF 私用区；解不出来就当空字符，别让 NaN 传下去
  return { kind: 'symbol', font, char: Number.isNaN(code) ? '' : String.fromCodePoint(code) };
}

// ── 表格 ──────────────────────────────────────────────────────────────────────
// 属性（宽度 / 边框 / 单元格边距 / 表格样式）见 parse-table-props.ts。
// gridSpan / vMerge 不走那条路：它们是**结构**，见 nodes.ts 上的注释。

function parseTable(ctx: Ctx, tbl: XmlElement): Table {
  const rows: TableRow[] = [];
  for (const el of children(tbl)) {
    if (el.name === 'w:tr') rows.push(parseRow(ctx, el));
    else if (el.name === 'w:sdt') {
      const content = child(el, 'w:sdtContent');
      if (content !== undefined) {
        for (const tr of children(content, 'w:tr')) rows.push(parseRow(ctx, tr));
      }
    } else if (el.name !== 'w:tblPr' && el.name !== 'w:tblGrid' && !IGNORED.has(el.name)) {
      unknown(ctx, el, 'w:tbl');
    }
  }
  return {
    kind: 'table',
    id: nextId(ctx, 'tbl'),
    props: parseTableProps(child(tbl, 'w:tblPr')),
    grid: parseTableGrid(child(tbl, 'w:tblGrid')),
    rows,
  };
}

function parseRow(ctx: Ctx, tr: XmlElement): TableRow {
  const cells: TableCell[] = [];
  for (const el of children(tr)) {
    if (el.name === 'w:tc') cells.push(parseCell(ctx, el));
    else if (el.name === 'w:sdt') {
      const content = child(el, 'w:sdtContent');
      if (content !== undefined) for (const tc of children(content, 'w:tc')) cells.push(parseCell(ctx, tc));
    } else if (el.name !== 'w:trPr' && el.name !== 'w:tblPrEx' && !IGNORED.has(el.name)) {
      unknown(ctx, el, 'w:tr');
    }
  }
  const row: TableRow = {
    kind: 'row',
    id: nextId(ctx, 'tr'),
    props: parseRowProps(child(tr, 'w:trPr')),
    cells,
  };
  // `w:tblPrEx` 是 `w:tblPr` 的子集（边框 / 边距 / 底纹 / look / 宽度），同一个解析器就够；
  // 里面的 `w:tblPrExChange`（修订痕迹）不在解析表里，自动被忽略
  const ex = child(tr, 'w:tblPrEx');
  if (ex !== undefined) row.propsEx = parseTableProps(ex);
  return row;
}

const V_MERGE = ['restart', 'continue'] as const;

function parseCell(ctx: Ctx, tc: XmlElement): TableCell {
  const tcPr = child(tc, 'w:tcPr');
  const blocks: Block[] = [];
  for (const el of children(tc)) {
    if (el.name === 'w:p') blocks.push(parseParagraph(ctx, el, parseParaProps(child(el, 'w:pPr'))));
    else if (el.name === 'w:tbl') blocks.push(parseTable(ctx, el));
    else if (el.name === 'w:sdt') blocks.push(...sdtBlocks(ctx, el));
    else if (el.name !== 'w:tcPr' && !IGNORED.has(el.name)) unknown(ctx, el, 'w:tc');
  }

  const vMergeEl = tcPr === undefined ? undefined : child(tcPr, 'w:vMerge');
  // <w:vMerge/> 不带 w:val 时默认是 continue，**不是** restart —— 反了会多画一份内容
  const vMerge = vMergeEl === undefined ? 'none' : (enumVal(attr(vMergeEl, 'w:val'), V_MERGE) ?? 'continue');

  const gridSpan = Number.parseInt(attrOf(tcPr && child(tcPr, 'w:gridSpan'), 'w:val') ?? '', 10);
  return {
    kind: 'cell',
    id: nextId(ctx, 'tc'),
    props: parseCellProps(tcPr),
    blocks,
    gridSpan: Number.isNaN(gridSpan) || gridSpan < 1 ? 1 : gridSpan,
    vMerge,
  };
}
