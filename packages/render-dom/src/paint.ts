/**
 * `DocumentLayout` → 元素树。整条流水线的**出口**：px 只在这一步出现（原则 1.3）。
 *
 * ## 一页一个 `<svg>`，viewBox 的单位是 **pt**
 *
 * 三个理由，都不是审美：
 *
 * 1. **逐字 x 只有 SVG 给得了。** `<text x="x1 x2 x3 …">` 原生支持逐字形定位，
 *    两端对齐、标点挤压、中西文间距造成的所有偏移直接喂进去就行。HTML 要做到同样精度
 *    只能一字一 span，DOM 会爆（开发计划 §0.3）
 * 2. **缩放不重排。** 内部坐标全是 pt，缩放只改 `<svg>` 的 width / height 属性，
 *    viewBox 一个字不动 —— 架构 §4.1 说的「缩放是 O(1) 的，永不触发重排」在这里落地
 * 3. **pt 而不是 twips，是为了能直接和真值对眼。** `fixtures/*.truth.json` 的单位就是 pt，
 *    原点也同样是纸的左上角、y 向下。于是 SVG 里读到的 `y="119.05"` 与真值里的
 *    `y: 119.05` 是同一个数，肉眼比对不需要换算
 *
 * ## 层序：底纹 → 文字 → 线
 *
 * 表格的三样东西分三遍画（`uw-table-shading` / `-content` / `-borders`），不是逐格画完
 * 再画下一格。原因是**格线是共享的**：相邻两格各画一次同一条线，后画的那格的底纹
 * 会把先画的线盖掉半条。分遍画之后线永远在最上层，也顺便让「同一条线只画一次」
 * 的去重（`seen`）能跨格生效。
 *
 * ## 已知的洞（写下来免得以为已经画了）
 *
 * - **图片 / 内嵌对象**：`ObjectItem` 在布局里只占位，这里连占位框都不画 ——
 *   画一个空框会让人以为「图加载失败」，比什么都没有更误导
 * - **run 级底纹与高亮**（`w:highlight` / `w:shd`）：`ResolvedRunProps` 里就没有这两项，
 *   要先在 model 侧补
 * - **纵向合并区的内容裁剪**：`vMerge="restart"` 那一格的内容整个算在起始行里
 *   （见 `table.ts` 的 `rowHeight` 注释），这里照画，不裁到合并区
 * - **粗体 / 斜体是「让浏览器合成」**：度量走的是常规字重的度量包（`@uw/fonts` 的
 *   包里一款字体只有一份度量），所以加粗后的字**画得比量得宽**。这不是渲染的锅，
 *   是度量包还没有按字重分档 —— 真值里加粗行的断行点对不上时先怀疑这里
 */
import type { Twips } from '@uw/core';
import { twipsToPt, twipsToPx } from '@uw/core';
import type {
  BlockLayout,
  CellLayout,
  DocumentLayout,
  LineFragment,
  LineLayout,
  PageLayout,
  PlacedBlock,
  PlacedHeaderFooter,
  PlacedParagraph,
  PlacedTable,
  RowLayout,
  TabLeader,
} from '@uw/layout';
import { contentHeightOf } from '@uw/layout';
import { defaultFontFamily } from './font-stack.ts';
import type { RElement } from './tree.ts';
import { el, fmt, fmtList, textEl } from './tree.ts';
import {
  DOUBLE_STRIKE_GAP_EM,
  LEADER_DOT_PITCH_EM,
  LEADER_FALLBACK_SIZE,
  LEADER_THICKNESS_EM,
  STRIKE_POSITION_EM,
  SUBSCRIPT_DROP_EM,
  SUPERSCRIPT_RAISE_EM,
  UNDERLINE_OFFSET_EM,
  UNDERLINE_THICKNESS_EM,
} from './uncalibrated.ts';

export interface RenderOptions {
  /**
   * 屏幕缩放。**只作用在 `<svg>` 的 width / height 上** —— 布局结果一个字节都不动，
   * 所以缩放是 O(1) 的，而且同一份 `DocumentLayout` 能同时挂在主视图和缩略图上
   */
  zoom?: number;
  /** Word 字体名 → CSS font-family，默认见 `defaultFontFamily` */
  fontFamily?: (family: string) => string;
  /** class 前缀。默认 `uw`，改它是为了同一页面上挂两份互不干扰的样式 */
  classPrefix?: string;
  /** 画出版心框与每一行的行盒。调试用，不进产物路径 */
  debug?: boolean;
}

interface Ctx {
  zoom: number;
  fontFamily: (family: string) => string;
  cls: (name: string) => string;
  debug: boolean;
}

const SVG_NS = 'http://www.w3.org/2000/svg';

function context(opts: RenderOptions): Ctx {
  const prefix = opts.classPrefix ?? 'uw';
  return {
    zoom: opts.zoom ?? 1,
    fontFamily: opts.fontFamily ?? defaultFontFamily,
    cls: (name) => `${prefix}-${name}`,
    debug: opts.debug === true,
  };
}

/** twips → pt。本文件里所有出现在属性上的坐标都过这一道 */
const pt = (tw: Twips): number => twipsToPt(tw);

// ── 文档与页 ──────────────────────────────────────────────────────────────────

/**
 * 整份文档 → 一个 `<div>`，每页一个 `<svg>`。
 *
 * 页与页之间的间距、滚动、视口虚拟化**都不在这里** —— 那是 `@uw/view` 的事
 * （架构 §3.1：view 是渲染器的调度者）。这里只保证「把第 n 页画出来」这一件事，
 * 页容器带着 `data-page`，调度者按它挂载 / 卸载。
 */
export function buildDocument(layout: DocumentLayout, opts: RenderOptions = {}): RElement {
  const ctx = context(opts);
  return el(
    'div',
    { class: ctx.cls('doc') },
    layout.pages.map((p) => buildPageWith(p, ctx)),
  );
}

/** 单页 → `<svg>`。视口虚拟化按页装卸时走这个入口 */
export function buildPage(page: PageLayout, opts: RenderOptions = {}): RElement {
  return buildPageWith(page, context(opts));
}

function buildPageWith(page: PageLayout, ctx: Ctx): RElement {
  const g = page.geometry;
  const children: RElement[] = [
    el('rect', {
      class: ctx.cls('page-bg'),
      x: '0',
      y: '0',
      width: fmt(pt(g.width)),
      height: fmt(pt(g.height)),
      fill: '#ffffff',
    }),
  ];
  if (ctx.debug) {
    children.push(
      el('rect', {
        class: ctx.cls('debug-content'),
        x: fmt(pt(g.content.x)),
        y: fmt(pt(g.content.y)),
        width: fmt(pt(g.content.width)),
        height: fmt(pt(g.content.height)),
        fill: 'none',
        stroke: '#2da44e',
        'stroke-width': '0.4',
        'stroke-dasharray': '3 2',
      }),
    );
  }

  // 页眉在正文**之前**画、页脚在之后：三者的框在 Word 里就不该重叠（版心是让开了的），
  // 万一重叠了（页眉长到把版心吃光），正文压在页眉上比反过来更容易看出是哪儿排错了
  if (page.header !== undefined) children.push(paintFrame(page.header, ctx));

  const inner: RElement[] = [];
  for (const block of page.blocks) paintBlock(block, ctx, inner);
  children.push(
    el(
      'g',
      { class: ctx.cls('content'), transform: `translate(${fmt(pt(g.content.x))} ${fmt(pt(g.content.y))})` },
      inner,
    ),
  );

  if (page.footer !== undefined) children.push(paintFrame(page.footer, ctx));

  const attrs: Record<string, string> = {
    xmlns: SVG_NS,
    class: page.filler === true ? `${ctx.cls('page')} ${ctx.cls('page-filler')}` : ctx.cls('page'),
    'data-page': String(page.index),
    // 显示页码与物理页序不是一回事（`w:pgNumType w:start` 会让某一节重新起算）
    'data-page-number': String(page.number),
    width: `${fmt(twipsToPx(g.width, ctx.zoom))}px`,
    height: `${fmt(twipsToPx(g.height, ctx.zoom))}px`,
    viewBox: `0 0 ${fmt(pt(g.width))} ${fmt(pt(g.height))}`,
  };
  return el('svg', attrs, children);
}

/**
 * 页眉 / 页脚。**坐标相对纸左上角**，与版心那个 `<g>` 平级 ——
 * 它不在版心里（版心是被它挤出来的），套进去会平白多偏一个上边距。
 */
function paintFrame(frame: PlacedHeaderFooter, ctx: Ctx): RElement {
  const inner: RElement[] = [];
  for (const block of frame.blocks) paintBlock(block, ctx, inner);
  if (ctx.debug) {
    inner.unshift(
      el('rect', {
        class: ctx.cls(`debug-${frame.kind}`),
        x: '0',
        y: '0',
        width: fmt(pt(frame.width)),
        height: fmt(pt(frame.height)),
        fill: 'none',
        stroke: '#bf8700',
        'stroke-width': '0.4',
        'stroke-dasharray': '3 2',
      }),
    );
  }
  return el(
    'g',
    {
      class: ctx.cls(frame.kind),
      'data-rel': frame.relId,
      transform: `translate(${fmt(pt(frame.x))} ${fmt(pt(frame.y))})`,
    },
    inner,
  );
}

// ── 块 ────────────────────────────────────────────────────────────────────────

function paintBlock(block: PlacedBlock, ctx: Ctx, out: RElement[]): void {
  if (block.kind === 'paragraph') out.push(paintPlacedParagraph(block, ctx));
  else out.push(paintPlacedTable(block, ctx));
}

function paintPlacedParagraph(p: PlacedParagraph, ctx: Ctx): RElement {
  const children: RElement[] = [];
  for (const placed of p.lines) paintLine(placed.line, 0, placed.y, ctx, children);
  return el('g', { class: ctx.cls('para'), 'data-id': p.id }, children);
}

/**
 * 一行。`x0` / `y0` 是这一行所在**容器**的原点（版心 or 单元格内容区），
 * 行自己的 `x` 是相对容器左边的 —— 两者相加才是页面坐标。
 */
function paintLine(line: LineLayout, x0: Twips, y0: Twips, ctx: Ctx, out: RElement[]): void {
  const baseline = y0 + line.baseline;
  if (ctx.debug) {
    out.push(
      el('rect', {
        class: ctx.cls('debug-line'),
        x: fmt(pt(x0 + line.x)),
        y: fmt(pt(y0)),
        width: fmt(pt(line.width)),
        height: fmt(pt(line.height)),
        fill: 'none',
        stroke: '#1f6feb',
        'stroke-width': '0.3',
      }),
    );
  }
  for (const leader of line.leaders) out.push(paintLeader(leader, line, x0, baseline, ctx));
  for (const frag of line.fragments) {
    // 装饰先画：下划线在文字之下，画在后面会盖住字的下半截
    for (const d of decorations(frag, x0, baseline, ctx)) out.push(d);
    out.push(paintFragment(frag, x0, baseline, ctx));
  }
}

/**
 * 一个渲染片段 → 一个 `<text>`。
 *
 * `w:w`（横向缩放）用 `transform="scale(s 1)"` 落地，同时把每个字的 x 除以 s ——
 * 缩放会把坐标系一起压扁，不除回去整段文字会往左缩成一堆。宽度那一侧早就
 * 折进 `CharItem.width` 了（`items.ts` 的 `scaledWidth`），这里改的**只是字形**。
 */
function paintFragment(frag: LineFragment, x0: Twips, baseline: Twips, ctx: Ctx): RElement {
  const style = frag.style;
  const scale = style.scale / 100;
  const xs = frag.glyphX.map((x) => pt(x0 + x) / (scale === 0 ? 1 : scale));
  const attrs: Record<string, string> = {
    class: ctx.cls('frag'),
    'data-run': frag.runId,
    x: fmtList(xs),
    y: fmt(baselineOf(frag, baseline)),
    'font-family': ctx.fontFamily(frag.font),
    'font-size': fmt(pt(frag.fontSize)),
    fill: cssColor(style.color),
    // SVG 不折叠空白，但也不会自动保留 —— 不写这一条，行首行尾的空格会被吃掉
    'xml:space': 'preserve',
  };
  if (style.bold) attrs['font-weight'] = 'bold';
  if (style.italic) attrs['font-style'] = 'italic';
  if (scale !== 1 && scale !== 0) attrs.transform = `scale(${fmt(scale)} 1)`;
  // 编号文字不在 document.xml 里：可选文本层与复制要跳过它（见 CharItem.numbering）
  if (frag.numbering === true) attrs['data-numbering'] = '1';
  // 域求值的结果：与编号相反，它**要**能被复制与 Ctrl+F 搜到，标出来是因为它不可编辑 ——
  // 文件里存的是上次算出来的旧值，这串字反查不到 DocPosition（见 CharItem.field）
  if (frag.field === true) attrs['data-field'] = '1';
  return textEl('text', attrs, frag.text);
}

/** 基线的最终 y：`w:position` 的升降 + 上下标的升降都落在这里，两者可叠加 */
function baselineOf(frag: LineFragment, baseline: Twips): number {
  const style = frag.style;
  let y = pt(baseline - style.position);
  if (style.vertAlign === 'superscript') y -= pt(frag.fontSize) * SUPERSCRIPT_RAISE_EM;
  else if (style.vertAlign === 'subscript') y += pt(frag.fontSize) * SUBSCRIPT_DROP_EM;
  return y;
}

/**
 * 下划线与删除线。
 *
 * 一段一画，不跨片段合并 —— 同一个 run 被切成几个片段只可能因为字体 / 字号 / 脚本变了
 * （见 `fragmentsOf`），那时下划线本来就该跟着变粗细，合并反而是错的。
 */
function decorations(frag: LineFragment, x0: Twips, baseline: Twips, ctx: Ctx): RElement[] {
  const style = frag.style;
  const out: RElement[] = [];
  const size = pt(frag.fontSize);
  const y = baselineOf(frag, baseline);
  const x1 = pt(x0 + frag.x);
  const x2 = x1 + pt(frag.width);
  const color = cssColor(style.color);

  if (style.underline !== 'none' && style.underline !== '') {
    const top = y + size * UNDERLINE_OFFSET_EM;
    const w = size * UNDERLINE_THICKNESS_EM;
    const dash = underlineDash(style.underline, size);
    out.push(rule(ctx, 'underline', x1, x2, top, w, color, dash));
    if (style.underline === 'double' || style.underline === 'dottedHeavy') {
      out.push(rule(ctx, 'underline', x1, x2, top + w * 2, w, color, dash));
    }
  }
  if (style.strike || style.doubleStrike) {
    const w = size * UNDERLINE_THICKNESS_EM;
    const mid = y - size * STRIKE_POSITION_EM;
    if (style.doubleStrike) {
      const gap = size * DOUBLE_STRIKE_GAP_EM;
      out.push(rule(ctx, 'strike', x1, x2, mid - gap / 2, w, color, undefined));
      out.push(rule(ctx, 'strike', x1, x2, mid + gap / 2, w, color, undefined));
    } else {
      out.push(rule(ctx, 'strike', x1, x2, mid, w, color, undefined));
    }
  }
  return out;
}

/** 一条横线。用 `<rect>` 而不是 `<line>`：线宽以中心线为准，改成矩形省一次心算 */
function rule(
  ctx: Ctx,
  name: string,
  x1: number,
  x2: number,
  y: number,
  thickness: number,
  color: string,
  dash: string | undefined,
): RElement {
  const attrs: Record<string, string> = {
    class: ctx.cls(name),
    x: fmt(x1),
    y: fmt(y),
    width: fmt(Math.max(0, x2 - x1)),
    height: fmt(thickness),
    fill: color,
  };
  if (dash !== undefined) {
    // 虚线下划线用 rect 画不出来，退回描边：把矩形当成一条居中的线
    attrs.fill = 'none';
    attrs.stroke = color;
    attrs['stroke-width'] = fmt(thickness);
    attrs['stroke-dasharray'] = dash;
    attrs.height = '0';
    attrs.y = fmt(y + thickness / 2);
  }
  return el('rect', attrs);
}

/**
 * `w:u/@w:val` 里认得出的虚线族 → dash 图案。
 *
 * 认不出的一律画实线：`w:u` 有二十来种取值（`wavyHeavy` `dashLongDotDotHeavy` …），
 * 都是「有线」的花样，退成不画会让文档看起来少了东西。波浪线画成实线是**已知的将就** ——
 * SVG 要画波浪得用 path，等有人真的抱怨再说。
 */
function underlineDash(val: string, size: number): string | undefined {
  if (val.startsWith('dotted') || val.startsWith('dotDash') || val.startsWith('dotDotDash')) {
    return `${fmt(size * 0.06)} ${fmt(size * 0.1)}`;
  }
  if (val.startsWith('dash')) return `${fmt(size * 0.24)} ${fmt(size * 0.16)}`;
  return undefined;
}

/**
 * 制表位前导符。真值里 Word 画的是一个个字符，我们画的是一条带图案的线 ——
 * 差别与钉死办法写在 `uncalibrated.ts` 的 `LEADER_DOT_PITCH_EM`。
 */
function paintLeader(leader: TabLeader, line: LineLayout, x0: Twips, baseline: Twips, ctx: Ctx): RElement {
  // 前导符自己没有字号，取本行最大的那个 —— 目录行里前导点与正文同高
  let size: Twips = 0;
  for (const f of line.fragments) if (f.fontSize > size) size = f.fontSize;
  const em = pt(size === 0 ? LEADER_FALLBACK_SIZE : size);
  const thickness = em * LEADER_THICKNESS_EM;
  const y = pt(baseline) + em * UNDERLINE_OFFSET_EM;
  const attrs: Record<string, string> = {
    class: ctx.cls('leader'),
    x1: fmt(pt(x0 + leader.x1)),
    x2: fmt(pt(x0 + leader.x2)),
    y1: fmt(y),
    y2: fmt(y),
    stroke: '#000000',
    'stroke-width': fmt(thickness),
  };
  if (leader.leader === 'dot' || leader.leader === 'middleDot') {
    attrs['stroke-dasharray'] = `${fmt(thickness)} ${fmt(em * LEADER_DOT_PITCH_EM - thickness)}`;
    attrs['stroke-linecap'] = 'round';
  } else if (leader.leader === 'hyphen') {
    attrs['stroke-dasharray'] = `${fmt(em * 0.3)} ${fmt(em * 0.2)}`;
  }
  // underscore / heavy 就是实线，只有粗细不同
  if (leader.leader === 'heavy') attrs['stroke-width'] = fmt(thickness * 2);
  return el('line', attrs);
}

// ── 表格 ──────────────────────────────────────────────────────────────────────

function paintPlacedTable(t: PlacedTable, ctx: Ctx): RElement {
  const shading: RElement[] = [];
  const content: RElement[] = [];
  const borders: RElement[] = [];
  // 同一条格线两边的格子会各解析出一份**相同**的结果（table-borders.ts 的保证），
  // 按几何位置去重，画一遍就够
  const seen = new Set<string>();

  for (const placed of t.rows) {
    const rowTop = placed.y;
    for (const cell of placed.row.cells) {
      if (cell.vMerge !== 'continue') {
        paintCellShading(cell, rowTop, placed.height, ctx, shading);
        paintCellContent(cell, rowTop, placed.height, ctx, content);
      }
      paintCellBorders(cell, t.columns, rowTop, placed.height, ctx, seen, borders);
    }
  }

  return el('g', { class: ctx.cls('table'), 'data-id': t.id, transform: `translate(${fmt(pt(t.x))} 0)` }, [
    el('g', { class: ctx.cls('table-shading') }, shading),
    el('g', { class: ctx.cls('table-content') }, content),
    el('g', { class: ctx.cls('table-borders') }, borders),
  ]);
}

function paintCellShading(
  cell: CellLayout,
  rowTop: Twips,
  rowHeight: Twips,
  ctx: Ctx,
  out: RElement[],
): void {
  const fill = shadingFill(cell.shading);
  if (fill === undefined) return;
  out.push(
    el('rect', {
      class: ctx.cls('cell-bg'),
      x: fmt(pt(cell.x)),
      y: fmt(pt(rowTop)),
      width: fmt(pt(cell.width)),
      height: fmt(pt(rowHeight)),
      fill,
    }),
  );
}

/**
 * `w:shd` → 填充色。
 *
 * `fill` 才是底色，`color` 是网点图案的前景色 —— 取反了「浅色底纹」会变成实心块
 * （model 的 `Shading` 注释里也写着这一条）。网点图案（`pct25` / `diagStripe` …）
 * **不画图案**，只按 `fill` 铺纯色：画图案要 SVG pattern，而公文里的底纹几乎全是 `clear`。
 */
function shadingFill(shading: CellLayout['shading']): string | undefined {
  if (shading === undefined) return undefined;
  if (shading.pattern === 'nil') return undefined;
  if (shading.fill === '' || shading.fill === 'auto') return undefined;
  return cssColor(shading.fill);
}

/**
 * 格内内容。格子里的块**自己不带 y**（与段落同理），所以这里现摞一遍：
 * 先按 `w:vAlign` 求出这一摞的起始 y，再逐块累加。
 */
function paintCellContent(
  cell: CellLayout,
  rowTop: Twips,
  rowHeight: Twips,
  ctx: Ctx,
  out: RElement[],
): void {
  const inner = contentHeightOf(cell.blocks);
  const avail = rowHeight - cell.paddingTop - cell.paddingBottom;
  let y = rowTop + cell.paddingTop;
  if (cell.verticalAlign === 'center') y += Math.max(0, (avail - inner) / 2);
  else if (cell.verticalAlign === 'bottom') y += Math.max(0, avail - inner);

  const x = cell.x + cell.paddingLeft;
  const children: RElement[] = [];
  paintBlockStack(cell.blocks, x, y, ctx, children);
  out.push(el('g', { class: ctx.cls('cell'), 'data-id': cell.cellId }, children));
}

/** 一摞块（段落 / 嵌套表格）从 `y0` 往下排。累加规则与 `contentHeightOf` 必须一致 */
function paintBlockStack(
  blocks: readonly BlockLayout[],
  x0: Twips,
  y0: Twips,
  ctx: Ctx,
  out: RElement[],
): void {
  let y = y0;
  for (const b of blocks) {
    if (b.kind === 'table') {
      for (const row of b.layout.rows) {
        paintNestedRow(row, x0 + b.layout.x, y, ctx, out);
        y += row.height;
      }
      continue;
    }
    y += b.layout.spaceBefore;
    for (const line of b.layout.lines) {
      paintLine(line, x0, y, ctx, out);
      y += line.height;
    }
    y += b.layout.spaceAfter;
  }
}

/**
 * 嵌套表格的一行。没有走 `paintPlacedTable` 是因为那条路要的是分页产物
 * （`PlacedRow` 带 y），而嵌套表格的行还没被分页碰过 —— 它的 y 是这里现摞出来的。
 * 代价是嵌套表格的格线不参与外层的去重，同一条线会画两遍（幂等，只是多几个元素）。
 */
function paintNestedRow(row: RowLayout, x0: Twips, y: Twips, ctx: Ctx, out: RElement[]): void {
  const seen = new Set<string>();
  const children: RElement[] = [];
  for (const cell of row.cells) {
    if (cell.vMerge === 'continue') continue;
    paintCellShading(cell, 0, row.height, ctx, children);
    paintCellContent(cell, 0, row.height, ctx, children);
    paintCellBorders(cell, [], 0, row.height, ctx, seen, children);
  }
  out.push(
    el('g', { class: ctx.cls('nested-row'), transform: `translate(${fmt(pt(x0))} ${fmt(pt(y))})` }, children),
  );
}

/**
 * 一格的四条边。
 *
 * 水平边**按列分段**（表头一格跨 3 列、下面 3 格，那条线就分 3 段各比各的，
 * 见 `table-borders.ts`），所以要拿网格列宽把段号换算回 x。`columns` 为空时
 * 退成整格一条 —— 嵌套表格走的就是这条（那里没有外层网格）。
 */
function paintCellBorders(
  cell: CellLayout,
  columns: readonly Twips[],
  rowTop: Twips,
  rowHeight: Twips,
  ctx: Ctx,
  seen: Set<string>,
  out: RElement[],
): void {
  const top = rowTop;
  const bottom = rowTop + rowHeight;
  for (const seg of cell.borders.top) {
    const [x1, x2] = segmentRange(cell, columns, seg.col, seg.span);
    pushLine(ctx, seen, out, x1, top, x2, top, seg.border);
  }
  for (const seg of cell.borders.bottom) {
    const [x1, x2] = segmentRange(cell, columns, seg.col, seg.span);
    pushLine(ctx, seen, out, x1, bottom, x2, bottom, seg.border);
  }
  pushLine(ctx, seen, out, cell.x, top, cell.x, bottom, cell.borders.left);
  const right = cell.x + cell.width;
  pushLine(ctx, seen, out, right, top, right, bottom, cell.borders.right);
  // 对角线不与谁共享，也不去重
  if (cell.borders.tl2br !== undefined) {
    pushLine(ctx, undefined, out, cell.x, top, right, bottom, cell.borders.tl2br);
  }
  if (cell.borders.tr2bl !== undefined) {
    pushLine(ctx, undefined, out, right, top, cell.x, bottom, cell.borders.tr2bl);
  }
}

function segmentRange(
  cell: CellLayout,
  columns: readonly Twips[],
  col: number,
  span: number,
): [Twips, Twips] {
  if (columns.length === 0) return [cell.x, cell.x + cell.width];
  let x1: Twips = 0;
  for (let i = 0; i < col && i < columns.length; i++) x1 += columns[i] as Twips;
  let x2 = x1;
  for (let i = col; i < col + span && i < columns.length; i++) x2 += columns[i] as Twips;
  return [x1, x2];
}

function pushLine(
  ctx: Ctx,
  seen: Set<string> | undefined,
  out: RElement[],
  x1: Twips,
  y1: Twips,
  x2: Twips,
  y2: Twips,
  border: CellLayout['borders']['left'],
): void {
  if (border === undefined || border.size <= 0) return;
  if (border.style === 'nil' || border.style === 'none') return;
  const key = `${fmt(x1)},${fmt(y1)},${fmt(x2)},${fmt(y2)}`;
  if (seen !== undefined) {
    if (seen.has(key)) return;
    seen.add(key);
  }
  const width = pt(border.size);
  const attrs: Record<string, string> = {
    class: ctx.cls('border'),
    x1: fmt(pt(x1)),
    y1: fmt(pt(y1)),
    x2: fmt(pt(x2)),
    y2: fmt(pt(y2)),
    stroke: cssColor(border.color),
    'stroke-width': fmt(width),
  };
  const dash = borderDash(border.style, width);
  if (dash !== undefined) attrs['stroke-dasharray'] = dash;
  if (border.style !== 'double') {
    out.push(el('line', attrs));
    return;
  }
  // 双线 = 两条 1/3 宽的线，中间空 1/3。总宽仍是 `w:sz`——它说的是整条线的视觉宽度，
  // 画成两条各 size 宽会让「双细线」变成两条粗线
  const thin = fmt(width / 3);
  // 垂直边往右让、水平边往下让：两条线之间空出 1/3
  const [dx, dy] = x1 === x2 ? [(width * 2) / 3, 0] : [0, (width * 2) / 3];
  out.push(el('line', { ...attrs, 'stroke-width': thin }));
  out.push(
    el('line', {
      ...attrs,
      'stroke-width': thin,
      x1: fmt(pt(x1) + dx),
      x2: fmt(pt(x2) + dx),
      y1: fmt(pt(y1) + dy),
      y2: fmt(pt(y2) + dy),
    }),
  );
}

/** `w:val` 的虚线族 → dash 图案。与下划线同理，认不出的画实线 */
function borderDash(style: string, width: number): string | undefined {
  if (style === 'dotted') return `${fmt(width)} ${fmt(width * 2)}`;
  if (style.startsWith('dash') || style.startsWith('dotDash') || style.startsWith('dotDotDash')) {
    return `${fmt(width * 3)} ${fmt(width * 2)}`;
  }
  return undefined;
}

// ── 颜色 ──────────────────────────────────────────────────────────────────────

const HEX6 = /^[0-9A-Fa-f]{6}$/;

/**
 * OOXML 的颜色 → CSS。
 *
 * `auto` 是「由渲染器挑一个与背景对比的颜色」，Word 在白底上给的就是黑 ——
 * 这里不做真正的对比度推断（底纹是任意色时才需要，公文里没有这种用法）。
 * 主题色（`w:themeColor`）在 model 那一层就没解析成 RGB，到这儿仍是原样的字符串，
 * 认不出的一律退到黑：退到透明或粉色都会让人以为是渲染 bug，黑至少是可读的。
 */
export function cssColor(value: string): string {
  if (HEX6.test(value)) return `#${value.toLowerCase()}`;
  return '#000000';
}
