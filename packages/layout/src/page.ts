/**
 * 分页 —— 把段落与表格摞进一页页的版心里，给每一行补上 y。
 *
 * 这是整条流水线的**最后一个坐标环节**：在此之前段落只知道「我有几行、每行多高、
 * 行顶到基线多远」，坐标原点是它自己的左上角；在此之后每一行都有了页号与页内的 y，
 * 渲染、域求值（PAGE / NUMPAGES）、表格跨页才谈得上。
 *
 * ## 为什么 y 是在这里补而不是在段落里算
 *
 * 段落的布局结果要能缓存与复用：改第一段不该让第五十段的坐标全部失效。所以
 * `ParagraphLayout` 一直不带 y，行的 y 靠**把前面各行的 height 累加**得到 ——
 * 这个累加只有在知道「前面还有谁、这一页放不放得下」的地方才做得了，也就是这里。
 *
 * ## 累加得准不准：gongwen-01 的 18 行全部实测过
 *
 * 「版心顶 + Σ行高 + 段间距 + 行内基线」直接就是 Word 的基线 y，18 行最大误差 **0.06pt**
 * （L3 的判据是 0.5pt）。也就是说 Phase 0 与基线穿刺定下的两条公式**可以逐行累加**，
 * 不需要任何「每页重新对齐到网格」的修正 —— 网格吸附已经吸在行高上了。
 * 断言在 `fixture.test.ts` 的 L3 一节。
 *
 * ## 三件**没做**的（写下来免得以为已经做了）
 *
 * - **页眉页脚**：`headers` / `footers` 指向的部件还没解析，版心顶固定取 `w:top`。
 *   Word 里页眉内容过深会把正文往下顶，补页眉时 `content.y` 要改成 max(top, 页眉底)
 * - **表格拆行**：行是原子的（一行放不下就整行挪到下一页）。Word 默认会把一行**内部**
 *   拆开，`w:cantSplit` 才禁止 —— 所以现在的行为等价于「全表 cantSplit」。
 *   只有单行高过剩余版心时才看得出差别，公文表格基本不会，但它是个洞
 * - **脚注 / 尾注 / 浮动对象**：完全不参与占位
 */
import type { DiagnosticSink, Twips } from '@uw/core';
import type { TextMeasurer } from '@uw/fonts';
import type {
  DocumentSettings,
  NodeId,
  ResolvedBlock,
  ResolvedBody,
  ResolvedParaProps,
  ResolvedTableRow,
  SectionProps,
} from '@uw/model';
import { layoutParagraph } from './paragraph.ts';
import type { RowLayout, TableLayout } from './table.ts';
import { layoutTable } from './table.ts';
import type { LineLayout, ParagraphLayout } from './types.ts';

// ── 输出的数据形状 ────────────────────────────────────────────────────────────

/** 一页的纸张与版心，全部相对**纸的左上角**，twips */
export interface PageGeometry {
  width: Twips;
  height: Twips;
  /** 版心。页上所有块的 x / y 都相对它的左上角 */
  content: { x: Twips; y: Twips; width: Twips; height: Twips };
}

/** 一行落到页上的位置 */
export interface PlacedLine {
  /** 在 `ParagraphLayout.lines` 里的下标 —— 跨页时后半段不是从 0 起 */
  index: number;
  /** 行顶相对**版心顶**。基线在页面里的绝对 y = `content.y + y + line.baseline` */
  y: Twips;
  line: LineLayout;
}

export interface PlacedParagraph {
  kind: 'paragraph';
  id: NodeId;
  /** 这一片的顶（第一行的行顶），相对版心顶。段前间距**不含**在内 */
  y: Twips;
  lines: PlacedLine[];
  /**
   * 这一片是不是段落的开头 / 结尾。跨页的段落会在两页上各出现一次，
   * 两个标记都为 false 的是中段（三页以上的长段落才有）。
   */
  first: boolean;
  last: boolean;
}

export interface PlacedRow {
  /** 在 `TableLayout.rows` 里的下标 */
  index: number;
  y: Twips;
  height: Twips;
  row: RowLayout;
  /**
   * 跨页时重复出来的表头行（`w:tblHeader`）。文档里它只有一份 ——
   * 命中测试与可选文本层必须跳过重复的那些，否则复制出来会多一遍表头。
   */
  repeated?: true;
}

export interface PlacedTable {
  kind: 'table';
  id: NodeId;
  /** 表格左边相对**版心左边**（`w:tblInd` 与 `w:jc` 已经算进去了） */
  x: Twips;
  y: Twips;
  width: Twips;
  columns: Twips[];
  rows: PlacedRow[];
  first: boolean;
  last: boolean;
}

export type PlacedBlock = PlacedParagraph | PlacedTable;

export interface PageLayout {
  /** 物理页序，0 起 */
  index: number;
  /** 显示页码：`w:pgNumType w:start` 会让它在某一节重新起算，所以与 `index` 不是一回事 */
  number: number;
  sectionIndex: number;
  geometry: PageGeometry;
  blocks: PlacedBlock[];
  /** `evenPage` / `oddPage` 为了凑奇偶补出来的空页 */
  filler?: true;
}

export interface DocumentLayout {
  pages: PageLayout[];
}

export interface LayoutDocumentOptions {
  measurer: TextMeasurer;
  settings: DocumentSettings;
  /** 四个字体桶全空时用哪款字体 */
  defaultFont?: string;
  diagnostics?: DiagnosticSink;
  /**
   * 分页规则。**标定用的接缝** —— 正常调用不要传，默认值就是实测出来的那一套
   * （见 `PAGINATION_RULES`）。`apps/fidelity` 的 `spike:page` 靠它把几种假设各跑一遍，
   * 证明「代码里实现的这一套」是唯一能复现 Word 的那一套。
   */
  rules?: Partial<PaginationRules>;
}

export interface PaginationRules {
  /** 孤行寡行的保底行数 */
  widowMinLines: number;
  /** 段前间距落在页首算不算 */
  spaceBeforeAtPageTop: boolean;
  /** keepNext 的接缝要给下一块留出多少 */
  keepNextJoin: 'min-chunk' | 'first-line' | 'whole-block';
}

/**
 * 分页规则的实测值。样本 `spike-page-01/02`（`pnpm --filter @uw/fidelity spike:page`），
 * 版心刻意做成「一页恰好 11 行、一行 18 个汉字、固定行距 20pt」，于是行高与字宽都不依赖
 * 任何待标定的度量，阶梯靠垫行的条数移动断页点。
 *
 * ① **孤行寡行保底 2 行**（`spike-page-01`，7 级阶梯 + 5 级关掉 widowControl 的对照）：
 *    垫 7 行时目标段落的自然断点是 4|1，Word 给的是 **3|2** —— 不做控制会是 4|1、
 *    下限若是 3 则应当整段推走（0|5）。垫 10 行（自然 1|4）Word 整段推走，与下限 2 吻合。
 *    对照组 5 级全部与「老实排」逐页一致，这同时验证了「一页 11 行」这个前提本身。
 *
 * ② **段前间距落在页首不算**（`spike-page-02` 的 C 组）：24pt 段前的段落被自动分页顶到页首时，
 *    首行基线 y = **72.74pt**，与其余每一页的首行基线**一模一样**；靠硬分页符顶上去的那一份
 *    也是 72.74pt。同一段排在页中间时，它与上一行的基线差是 43.95pt ≈ 20 + 24 ——
 *    也就是段前间距**本身没问题，只是在页首被丢掉**。原先按规范里
 *    `w:suppressSpBfAfterPgBrk` 的存在推断「默认要加」，推反了。
 *
 * ③ **keepNext 要留出的是下一块「最少能放多少」**（`spike-page-02` 的 A / D 两组）：
 *    A 组是 3 行的 keepNext 段落 P 后面跟 2 行的 Q（Q 的 widowControl 是默认的开）。垫 7 行时
 *    版心还剩 4 行，按「留一行」算 P 的三行加 Q 的首行正好 4 行放得下，Word 却在 P 的第 2 行
 *    就断了 —— 因为 Q 只有 2 行，孤行寡行不许它拆，Q **整块**都得下去，P 的末行跟着走。
 *    D 组把「最少能放多少」与「整块」分开：Q 换成 5 行（孤行寡行只要求它留 2 行），
 *    Word 按 2 行留 —— 也就是说它算的是「Q 至少要占的那一截」，不是「整个 Q」。
 *
 * 三条规则各自的分辨力（`spike:page` 把 3 × 2 × 3 种组合逐页对了一遍，两份样本共 50 页）：
 * 寡行下限 1 / 2 / 3 → 42 / **50** / 40 页；页首段前 加 / 不加 → 49 / **50**；
 * 接缝 first-line / whole-block / min-chunk → 46 / 46 / **50**。实现的这一组是唯一的满分，
 * 没有并列 —— 并列就说明样本分不开，那时该加密阶梯而不是随便挑一个。
 */
export const PAGINATION_RULES: PaginationRules = {
  widowMinLines: 2,
  spaceBeforeAtPageTop: false,
  keepNextJoin: 'min-chunk',
};

// ── 分页过程中的状态 ──────────────────────────────────────────────────────────

/**
 * `page` 为 undefined 表示「下一次放东西时才开页」—— **页是惰性开的**。
 *
 * 这样文档末尾的硬分页符不会凭空多出一张空页（它后面没有内容了），节与节之间也不会
 * 因为「先开页再发现这一节是空的」留下垃圾页。真正需要空页的只有 `evenPage` /
 * `oddPage` 补出来的那种，那一种是显式造的，带 `filler` 标记。
 */
interface Flow {
  opts: LayoutDocumentOptions;
  rules: PaginationRules;
  pages: PageLayout[];
  page: PageLayout | undefined;
  /** 游标：下一块内容的顶，相对版心顶 */
  y: Twips;
  geometry: PageGeometry;
  sectionIndex: number;
  /** 下一张开出来的页拿到的页码 */
  nextNumber: number;
}

/** 排完行、还没分页的中间形态。分页只关心高度，所以两种块在这里被拉平成同一个层级 */
type Prepared =
  | { kind: 'paragraph'; id: NodeId; layout: ParagraphLayout; props: ResolvedParaProps }
  | { kind: 'table'; id: NodeId; layout: TableLayout; rows: ResolvedTableRow[] };

export function layoutDocument(body: ResolvedBody, opts: LayoutDocumentOptions): DocumentLayout {
  const first = body.sections[0];
  const flow: Flow = {
    opts,
    rules: { ...PAGINATION_RULES, ...opts.rules },
    pages: [],
    page: undefined,
    y: 0,
    geometry: pageGeometry(first?.props ?? FALLBACK_SECTION, opts),
    sectionIndex: 0,
    nextNumber: 1,
  };

  body.sections.forEach((section, index) => {
    flow.sectionIndex = index;
    startSection(flow, section.props, index);

    const blocks = section.blocks.map((b) => prepare(b, section.props, opts));
    // keepNext 把相邻的块串成「接缝不许跨页」的链，接缝高度要在排**上一块**时就知道
    blocks.forEach((b, i) => {
      place(flow, b, joinHeight(blocks, i, flow.rules));
    });

    // 空节也要占一页：Word 里一个分节符至少产生一页，否则页码序列就断了
    if (section.blocks.length === 0) currentPage(flow);
  });

  // 空文档也得有一页 —— 渲染层拿到 pages: [] 只能画白屏，那与「文档是空的」不是一回事
  if (flow.pages.length === 0) currentPage(flow);
  return { pages: flow.pages };
}

/** 一份节都没有的文档（不合法但解析得出来）拿它兜底，A4 纵向 */
const FALLBACK_SECTION: SectionProps = {
  page: { width: 11906, height: 16838, orientation: 'portrait' },
  type: 'nextPage',
  margin: { top: 1440, right: 1800, bottom: 1440, left: 1800, header: 851, footer: 992, gutter: 0 },
  docGrid: { type: 'default', linePitch: 0, charSpace: 0 },
  columns: 1,
  titlePage: false,
  headers: [],
  footers: [],
};

// ── 页面几何 ──────────────────────────────────────────────────────────────────

/**
 * 版心 = 纸张减页边距。装订线（`w:gutter`）默认加在**左边**，`w:gutterAtTop` 时加在上边 ——
 * 两种情形下它都是从版心里扣掉的，不是往纸外扩，忘了扣会让每一行都偏出装订线那么多。
 *
 * 页眉页脚还没做，所以 `content.y` 直接取 `w:top`（见文件头）。
 * 对称页边距（`w:mirrorMargins`）也没做：那要按页码奇偶交换左右，得等页码定下来之后再算。
 */
export function pageGeometry(props: SectionProps, opts?: { settings?: DocumentSettings }): PageGeometry {
  const { page, margin } = props;
  const atTop = opts?.settings?.gutterAtTop === true;
  const left = margin.left + (atTop ? 0 : margin.gutter);
  const top = margin.top + (atTop ? margin.gutter : 0);
  return {
    width: page.width,
    height: page.height,
    content: {
      x: left,
      y: top,
      width: page.width - left - margin.right,
      height: page.height - top - margin.bottom,
    },
  };
}

// ── 分节 ──────────────────────────────────────────────────────────────────────

function startSection(flow: Flow, props: SectionProps, index: number): void {
  const geometry = pageGeometry(props, flow.opts);

  if (props.columns > 1) {
    // 多栏是明确的非目标。按单栏排出来的**行长是错的**，所以必须说一声，
    // 而不是安静地画一份看起来没毛病的东西
    flow.opts.diagnostics?.warn(
      'multi-column-unsupported',
      `第 ${index + 1} 节是 ${props.columns} 栏排版，按单栏排 —— 行长与断行点都会与 Word 不同`,
    );
  }

  if (index > 0) {
    const continuous = props.type === 'continuous';
    // continuous 的本意是「接着上一节往下排，不换页」。但版心一换，同一页上就要有两个
    // 不同的版心框，而 PageLayout 只有一个 —— 与其画错，不如换页并说明
    if (continuous && sameGeometry(flow.geometry, geometry)) {
      flow.geometry = geometry;
      return;
    }
    if (continuous) {
      flow.opts.diagnostics?.warn(
        'continuous-section-geometry-changed',
        `第 ${index + 1} 节是连续分节符但页面设置变了，按换页处理 —— 一页只能有一个版心`,
      );
    }
    flow.page = undefined;
  }

  flow.geometry = geometry;
  flow.y = 0;

  if (props.pageNumStart !== undefined) flow.nextNumber = props.pageNumStart;

  // 奇偶起始：页码对不上就补一张空页。补出来的那张用的是**新节**的版心（上一节的页
  // 已经关掉了），页码也占掉一个号 —— 与 `w:pgNumType` 同时出现时哪个先算没有真值，
  // 拿一份「evenPage + pgNumStart=5」的样本就能钉死
  if (props.type === 'evenPage' || props.type === 'oddPage') {
    const wantEven = props.type === 'evenPage';
    if ((flow.nextNumber % 2 === 0) !== wantEven) {
      currentPage(flow).filler = true;
      flow.page = undefined;
      flow.y = 0;
    }
  }
}

function sameGeometry(a: PageGeometry, b: PageGeometry): boolean {
  return (
    a.width === b.width &&
    a.height === b.height &&
    a.content.x === b.content.x &&
    a.content.y === b.content.y &&
    a.content.width === b.content.width &&
    a.content.height === b.content.height
  );
}

// ── 页 ────────────────────────────────────────────────────────────────────────

function currentPage(flow: Flow): PageLayout {
  if (flow.page !== undefined) return flow.page;
  const page: PageLayout = {
    index: flow.pages.length,
    number: flow.nextNumber,
    sectionIndex: flow.sectionIndex,
    geometry: flow.geometry,
    blocks: [],
  };
  flow.nextNumber += 1;
  flow.pages.push(page);
  flow.page = page;
  flow.y = 0;
  return page;
}

function breakPage(flow: Flow): void {
  flow.page = undefined;
  flow.y = 0;
}

/** 当前页上已经放过东西了吗 —— 「挪到下一页」只有这时候才有意义，否则会空转出一串空页 */
function pageHasContent(flow: Flow): boolean {
  return flow.page !== undefined && flow.page.blocks.length > 0;
}

function availHeight(flow: Flow): Twips {
  return flow.geometry.content.height - flow.y;
}

// ── 块的准备 ──────────────────────────────────────────────────────────────────

function prepare(b: ResolvedBlock, section: SectionProps, opts: LayoutDocumentOptions): Prepared {
  const shared = {
    measurer: opts.measurer,
    settings: opts.settings,
    docGrid: section.docGrid,
    ...(opts.defaultFont === undefined ? {} : { defaultFont: opts.defaultFont }),
  };
  const width = pageGeometry(section, opts).content.width;
  if (b.kind === 'paragraph') {
    return {
      kind: 'paragraph',
      id: b.id,
      layout: layoutParagraph(b, { ...shared, contentWidth: width }),
      props: b.props,
    };
  }
  return {
    kind: 'table',
    id: b.id,
    layout: layoutTable(b, { ...shared, availWidth: width }),
    rows: b.rows,
  };
}

/**
 * `w:keepNext` 的接缝高度：本块的**末行**与下一块**必须留在本页的那一截**要一起放得下。
 *
 * 「必须留在本页的那一截」不等于「下一块的第一行」—— 这是被真值推翻的第一版实现
 * （`spike-page-02` 的 A 组）：下一段只有 2 行时，孤行寡行不许它拆，于是它**整块**
 * 都得跟着走，接缝要按 2 行算。按一行算的话本页会多收一行，往后每一页都跟着错位。
 *
 * keepNext 也不是「整段跟着走」：一段十行的 keepNext 段落照样能拆到两页，
 * 只要它的最后一行与下一块该留的那一截还在一起（同一批样本的四级阶梯都是这么断的）。
 *
 * 下一块整块都得走时，接缝要接着往下串（A 的末行 + 整个 B + C 该留的那一截）——
 * 标题、副标题、正文首行这种三连在公文里很常见。
 */
function joinHeight(blocks: readonly Prepared[], i: number, rules: PaginationRules): Twips {
  const self = blocks[i];
  if (self === undefined || self.kind !== 'paragraph' || !self.props.keepNext) return 0;
  const next = blocks[i + 1];
  if (next === undefined) return 0;
  if (rules.keepNextJoin === 'whole-block') return blockHeight(next) + gapBetween(self, next);

  if (next.kind === 'table') {
    // 表格按行分页，最少能放的就是第一行（`w:cantSplit` 与拆行都还没做，见文件头）
    return next.layout.rows[0]?.height ?? 0;
  }
  const lines = next.layout.lines;
  const min = rules.keepNextJoin === 'first-line' ? 1 : minChunkLines(next.props, lines.length, rules);
  let h = gapBetween(self, next);
  for (let k = 0; k < Math.min(min, lines.length); k++) h += lines[k]?.height ?? 0;
  // 下一块整块都留不下来时，它的接缝也串上来
  return min >= lines.length ? h + joinHeight(blocks, i + 1, rules) : h;
}

/** 两块之间的空当。表格没有段前段后间距 */
function gapBetween(self: Prepared, next: Prepared): Twips {
  const after = self.kind === 'paragraph' ? self.layout.spaceAfter : 0;
  const before = next.kind === 'paragraph' ? next.layout.spaceBefore : 0;
  return after + before;
}

/**
 * 一个段落**最少**要在本页留下几行 —— 也就是「它开始排了，就至少占这么多」。
 *
 * `keepLines` 与「拆开会违反孤行寡行」两种情形下答案都是整段（n）：
 * n < 2×min 时任何一刀都会让某一边不足 min，只能整段放或者整段不放。
 */
function minChunkLines(props: ResolvedParaProps, n: number, rules: PaginationRules): number {
  if (props.keepLines) return n;
  if (!props.widowControl) return 1;
  return n >= 2 * rules.widowMinLines ? rules.widowMinLines : n;
}

function blockHeight(b: Prepared): Twips {
  if (b.kind === 'table') return b.layout.rows.reduce((sum, r) => sum + r.height, 0);
  return b.layout.lines.reduce((sum, l) => sum + l.height, 0);
}

function place(flow: Flow, b: Prepared, join: Twips): void {
  if (b.kind === 'paragraph') placeParagraph(flow, b, join);
  else placeTable(flow, b, join);
}

// ── 段落 ──────────────────────────────────────────────────────────────────────

function placeParagraph(flow: Flow, b: Extract<Prepared, { kind: 'paragraph' }>, join: Twips): void {
  const lines = b.layout.lines;
  const total = lines.length;
  if (total === 0) return;

  // 页首不再为它空跑一页：`w:pageBreakBefore` 说的是「本段从新的一页开始」，
  // 已经在新的一页上就已经满足了
  if (b.props.pageBreakBefore && pageHasContent(flow)) breakPage(flow);

  // 段前间距落在页首**不算**（实测，见 PAGINATION_RULES ②）。判断必须在开页**之前**做：
  // `currentPage()` 一开页，`pageHasContent()` 看的就是新页了。
  // 段落整段被推到下一页的那条路不用管：`flow.y` 会在换页时清零，加过的段前间距自然就没了
  const atPageTop = !pageHasContent(flow);
  currentPage(flow);
  if (!atPageTop || flow.rules.spaceBeforeAtPageTop) flow.y += b.layout.spaceBefore;

  let i = 0;
  while (i < total) {
    const raw = fitLines(lines, i, availHeight(flow), join);
    let count = adjust(raw, lines, i, b.props, flow.rules);

    if (count === 0) {
      // 空页上都放不下就只能硬塞（溢出版心），否则这个循环永远换页换不完
      if (pageHasContent(flow)) {
        breakPage(flow);
        currentPage(flow);
        continue;
      }
      count = Math.max(1, raw);
    }

    // 硬分页符（`w:br w:type="page"`）在哪一行就在哪一行截断。单栏文档里 column 等同于 page
    const hard = hardBreakAt(lines, i, count);
    if (hard >= 0) count = hard - i + 1;

    emitLines(flow, b, i, count, total);
    i += count;

    if (i < total || hard >= 0) breakPage(flow);
  }

  flow.y += b.layout.spaceAfter;
}

/** 从 `from` 起，`avail` 的高度里**老实**装得下几行（可能是 0）。末行要连接缝一起量 */
function fitLines(lines: readonly LineLayout[], from: number, avail: Twips, join: Twips): number {
  let used = 0;
  let count = 0;
  for (let k = from; k < lines.length; k++) {
    const line = lines[k];
    if (line === undefined) break;
    const need = used + line.height + (k === lines.length - 1 ? join : 0);
    if (need > avail) break;
    used += line.height;
    count += 1;
  }
  return count;
}

/**
 * 孤行寡行与 `w:keepLines` 的调整。三条都只在「这一段要被拆开」时才谈。
 *
 * 顺序不能反：先按寡行把行往下一页赶，赶完再看本页留下的够不够孤行的下限 ——
 * 反过来会得到「上面留 2 行、下面只剩 1 行」这种两头都不合规的结果。
 */
function adjust(
  count: number,
  lines: readonly LineLayout[],
  from: number,
  props: ResolvedParaProps,
  rules: PaginationRules,
): number {
  const rest = lines.length - from;
  if (count >= rest) return count;

  // 整段不许拆：放不下就整段挪到下一页。只在段落的**开头**判，
  // 已经拆过一次的段落（from > 0）再往回讲没有意义
  if (from === 0 && props.keepLines) return 0;
  if (!props.widowControl) return count;

  const min = rules.widowMinLines;
  let n = count;
  if (n > 0 && rest - n < min) n = Math.max(0, rest - min); // 寡行：下一页至少接走 min 行
  if (n > 0 && n < min) n = 0; // 孤行：本页底至少留下 min 行
  return n;
}

/** `[from, from+count)` 里第一个带硬分页符的行下标，没有则 -1 */
function hardBreakAt(lines: readonly LineLayout[], from: number, count: number): number {
  for (let k = from; k < from + count; k++) {
    const after = lines[k]?.breakAfter;
    if (after === 'page' || after === 'column') return k;
  }
  return -1;
}

function emitLines(
  flow: Flow,
  b: Extract<Prepared, { kind: 'paragraph' }>,
  from: number,
  count: number,
  total: number,
): void {
  const page = currentPage(flow);
  const placed: PlacedLine[] = [];
  const top = flow.y;
  for (let k = from; k < from + count; k++) {
    const line = b.layout.lines[k];
    if (line === undefined) break;
    placed.push({ index: k, y: flow.y, line });
    flow.y += line.height;
  }
  page.blocks.push({
    kind: 'paragraph',
    id: b.id,
    y: top,
    lines: placed,
    first: from === 0,
    last: from + count >= total,
  });
}

// ── 表格 ──────────────────────────────────────────────────────────────────────

/**
 * 表格按**行**分页：一行放不下就整行挪到下一页（差别见文件头）。
 *
 * `w:tblHeader` 的重复表头是真的要在每一续页顶部再画一遍，所以它**占高度**：
 * 判断「这一页还能放几行」时先把重复表头的高度扣掉，否则续页会多收一行、溢出版心。
 */
function placeTable(flow: Flow, b: Extract<Prepared, { kind: 'table' }>, join: Twips): void {
  const rows = b.layout.rows;
  if (rows.length === 0) return;

  // 表头只认**开头连续**的那几行：中间某行写了 tblHeader 是无效的（Word 也这么处理）
  let headerCount = 0;
  while (headerCount < rows.length && b.rows[headerCount]?.props.header === true) headerCount += 1;
  // 整张表都是表头行就当没有表头，否则续页会先重复一遍表头、再重复一遍…… 永远排不完
  if (headerCount >= rows.length) headerCount = 0;
  const headerHeight = rows.slice(0, headerCount).reduce((sum, r) => sum + r.height, 0);

  currentPage(flow);

  let i = 0;
  while (i < rows.length) {
    const repeat = i >= headerCount && i > 0 ? headerCount : 0;
    const avail = availHeight(flow) - (repeat > 0 ? headerHeight : 0);
    let count = fitRows(rows, i, avail, join);

    if (count === 0) {
      if (pageHasContent(flow)) {
        breakPage(flow);
        currentPage(flow);
        continue;
      }
      count = 1; // 一行高过一整页：先硬塞，等表格拆行做了再谈
    }

    emitRows(flow, b, i, count, repeat);
    i += count;
    if (i < rows.length) breakPage(flow);
  }
}

function fitRows(rows: readonly RowLayout[], from: number, avail: Twips, join: Twips): number {
  let used = 0;
  let count = 0;
  for (let k = from; k < rows.length; k++) {
    const row = rows[k];
    if (row === undefined) break;
    const need = used + row.height + (k === rows.length - 1 ? join : 0);
    if (need > avail) break;
    used += row.height;
    count += 1;
  }
  return count;
}

function emitRows(
  flow: Flow,
  b: Extract<Prepared, { kind: 'table' }>,
  from: number,
  count: number,
  repeat: number,
): void {
  const page = currentPage(flow);
  const placed: PlacedRow[] = [];
  const top = flow.y;

  for (let k = 0; k < repeat; k++) {
    const row = b.layout.rows[k];
    if (row === undefined) continue;
    placed.push({ index: k, y: flow.y, height: row.height, row, repeated: true });
    flow.y += row.height;
  }
  for (let k = from; k < from + count; k++) {
    const row = b.layout.rows[k];
    if (row === undefined) break;
    placed.push({ index: k, y: flow.y, height: row.height, row });
    flow.y += row.height;
  }

  page.blocks.push({
    kind: 'table',
    id: b.id,
    x: b.layout.x,
    y: top,
    width: b.layout.width,
    columns: [...b.layout.columns],
    rows: placed,
    first: from === 0,
    last: from + count >= b.layout.rows.length,
  });
}
