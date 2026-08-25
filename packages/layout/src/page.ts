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
 * ## 版心不再是「纸减页边距」
 *
 * 页眉页脚进来之后，**页边距是最小值不是固定值**：版心顶 = max(`w:top`, 页眉底)，
 * 版心底 = min(纸高 − `w:bottom`, 页脚顶)。所以每一页的版心要等它自己的页眉页脚排完
 * 才知道，`currentPage()` 是唯一有资格算它的地方（首页 / 偶数页用的页眉长度可以不同，
 * 同一节里各页的版心因此可以不一样高）。实测见 `header-footer.ts` 的 `HEADER_RULES`。
 *
 * ## 两件**没做**的（写下来免得以为已经做了）
 *
 * - **表格拆行**：行是原子的（一行放不下就整行挪到下一页）。Word 默认会把一行**内部**
 *   拆开，`w:cantSplit` 才禁止 —— 所以现在的行为等价于「全表 cantSplit」。
 *   只有单行高过剩余版心时才看得出差别，公文表格基本不会，但它是个洞
 * - **脚注 / 尾注 / 浮动对象**：完全不参与占位
 */
import type { DiagnosticSink, Twips } from '@uw/core';
import type { TextMeasurer } from '@uw/fonts';
import type {
  AnchorPos,
  DocumentSettings,
  NodeId,
  ResolvedBlock,
  ResolvedBody,
  ResolvedParaProps,
  ResolvedTableRow,
  SectionProps,
} from '@uw/model';
import { formatNumber, walkBlocks } from '@uw/model';
import type { HeaderFooterSource, HeaderRules, PlacedHeaderFooter, StackResult } from './header-footer.ts';
import {
  contentWithHeaderFooter,
  frameOf,
  HEADER_RULES,
  pickHeaderFooter,
  stackBlocks,
} from './header-footer.ts';
import type { ObjectRules } from './line-height.ts';
import { OBJECT_RULES } from './line-height.ts';
import { layoutParagraph } from './paragraph.ts';
import type { RowLayout, TableLayout } from './table.ts';
import { layoutTable } from './table.ts';
import type { LineFloat, LineLayout, LineObject, ParagraphLayout } from './types.ts';

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
  /**
   * 这一页的版心。**注意它不等于「纸减页边距」** —— 页眉页脚长过页边距时会把它挤窄，
   * 所以同一节里各页的 `content` 可以不同（首页页眉与偶数页页眉长度不一样就够了）
   */
  geometry: PageGeometry;
  blocks: PlacedBlock[];
  /** 这一页的页眉 / 页脚。没有引用、或者内容是空的时候缺席 */
  header?: PlacedHeaderFooter;
  footer?: PlacedHeaderFooter;
  /**
   * 这一页上**不参与文字流**的浮动对象（印章、水印、衬在文字下的红头），
   * 坐标相对**纸左上角**、已按 z 序排好。渲染层画的是这一份，不是 `LineLayout.floats`
   * （那一份是输入：与页无关、可缓存，见 types.ts）。
   */
  floats?: PlacedFloat[];
  /** `evenPage` / `oddPage` 为了凑奇偶补出来的空页 */
  filler?: true;
}

/**
 * 一个浮动对象在纸上的位置。
 *
 * 与 `PlacedBlock` 不同，它的坐标**不相对版心**：`wp:anchor` 的参照物可以是纸、页边距、
 * 段落…… 换算完统一落到纸坐标上，渲染层才不必认识那六种参照物（它也不该认识 ——
 * 那是布局）。
 */
export interface PlacedFloat {
  runId: NodeId;
  contentIndex: number;
  /** 相对**纸左上角** */
  x: Twips;
  y: Twips;
  width: Twips;
  height: Twips;
  objectKind: LineObject['objectKind'];
  image?: LineObject['image'];
  alt?: string;
  graphic?: string;
  /** 衬于文字下方。渲染层据此决定画在正文之前还是之后 */
  behindDoc: boolean;
  /** z 序（`wp:anchor@relativeHeight`），同一页内已按它升序排好 */
  z: number;
}

export interface DocumentLayout {
  pages: PageLayout[];
}

/**
 * 页眉页脚里一个域该怎么算。`clear` 是结果区里除第一个 run 以外的那些 ——
 * Word 常把一个数字切成好几个 `w:t`，不清掉旧的会留在页面上（与正文域同理）。
 */
export interface HeaderFieldSpec {
  type: 'PAGE' | 'NUMPAGES' | 'SECTIONPAGES' | 'clear';
  /** `\*` 开关解析出来的数字格式。缺席时 PAGE 跟着本节的 `w:pgNumType w:fmt` */
  format?: string;
}

export interface HeaderFieldPlan {
  fields: ReadonlyMap<NodeId, HeaderFieldSpec>;
  /**
   * 上一趟排出来的总页数 / 每节页数。**第一趟没有**（页数还不知道），
   * 那时 NUMPAGES 保留文件里存着的旧值 —— 迭代由 `layoutDocumentWithFields` 转，
   * 收敛之后这两个数就是自洽的
   */
  totalPages?: number;
  sectionPages?: readonly number[];
}

export interface LayoutDocumentOptions {
  measurer: TextMeasurer;
  settings: DocumentSettings;
  /** 四个字体桶全空时用哪款字体 */
  defaultFont?: string;
  /**
   * 域求值的结果（run id → 显示的文字）。**这里只是照着用**，算它是
   * `layoutDocumentWithFields()` 的事（fields.ts）—— 页码依赖分页、分页又依赖页码的宽度，
   * 那个循环必须在分页**外面**转，否则这一趟排版就得递归调用自己。
   */
  fieldValues?: ReadonlyMap<NodeId, string>;
  /**
   * 页眉页脚的内容，关系 id → 级联完的块。直接把 `LoadedDocument.headerFooters` 传进来即可。
   * 不传 = 不画页眉页脚，版心也就退回「纸减页边距」。
   */
  headerFooters?: HeaderFooterSource;
  /**
   * 页眉页脚里那些**要算**的域。与正文的 `fieldValues` 分开走，是因为同一个 run
   * 在每一页显示的**不是同一串字**（`{ PAGE }` 每页都不同），一张全局的
   * 「run id → 文字」表按定义就装不下它。这里存的是「怎么算」，算在开页那一刻做。
   */
  headerFields?: HeaderFieldPlan;
  diagnostics?: DiagnosticSink;
  /**
   * 分页规则。**标定用的接缝** —— 正常调用不要传，默认值就是实测出来的那一套
   * （见 `PAGINATION_RULES`）。`apps/fidelity` 的 `spike:page` 靠它把几种假设各跑一遍，
   * 证明「代码里实现的这一套」是唯一能复现 Word 的那一套。
   */
  rules?: Partial<PaginationRules>;
  /**
   * 页眉页脚的几何规则。同样是**标定用的接缝**，正常调用不要传 ——
   * `apps/fidelity` 的 `spike:header` 靠它把 8 组假设各跑一遍，见 `HEADER_RULES`
   */
  headerRules?: Partial<HeaderRules>;
  /**
   * 内嵌对象的行盒规则。同样是**标定用的接缝**，正常调用不要传 ——
   * `apps/fidelity` 的 `spike:image` 靠它把 4 组假设各跑一遍，见 `OBJECT_RULES`
   */
  objectRules?: Partial<ObjectRules>;
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
  headerRules: HeaderRules;
  pages: PageLayout[];
  page: PageLayout | undefined;
  /** 游标：下一块内容的顶，相对版心顶 */
  y: Twips;
  /**
   * 本节的**纸面**几何（只减了页边距与装订线）。每一页的版心从它出发再减页眉页脚，
   * 所以这一份与 `PageLayout.geometry` **不是同一个东西** —— 拿它去算「还剩多高」会多算
   */
  geometry: PageGeometry;
  /** 全部节的属性，`pickHeaderFooter` 要往回找「上一节的那一份」 */
  sections: SectionProps[];
  sectionIndex: number;
  /** 下一张开出来的页是不是本节的第一页 —— `w:titlePg` 靠它 */
  sectionFirstPage: boolean;
  /** 下一张开出来的页拿到的页码 */
  nextNumber: number;
  /**
   * 排好的页眉页脚，按「内容 + 这一页的域文字」缓存。纯静态的页眉（绝大多数）
   * 全文档只排一次，几百页共用同一份数据 —— 只读，且 `structuredClone` 保得住这种共享
   */
  hf: Map<string, StackResult>;
  /** 哪些部件里有要算的域（关系 id）。没有的那些缓存键里就不必带页码 */
  hfDynamic: Set<string>;
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
    headerRules: { ...HEADER_RULES, ...opts.headerRules },
    pages: [],
    page: undefined,
    y: 0,
    geometry: pageGeometry(first?.props ?? FALLBACK_SECTION, opts),
    sections: body.sections.map((s) => s.props),
    sectionIndex: 0,
    sectionFirstPage: true,
    nextNumber: 1,
    hf: new Map(),
    hfDynamic: dynamicParts(opts.headerFooters, opts.headerFields),
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
  // 浮动对象等**整页排完**再算：它的参照物可以是「纸」「页边距」，那两样要等这一页的
  // 版心定下来（页眉长度会挤窄版心，见文件头）。反过来它不影响任何一行的位置，
  // 所以放在最后一趟是安全的
  for (const page of flow.pages) placeFloats(page);
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
      // 没换页就谈不上「本节的第一页」：这一页是上一节开的，页眉早画好了
      flow.sectionFirstPage = flow.page === undefined;
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
  flow.sectionFirstPage = true;

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
  const number = flow.nextNumber;
  // 页眉页脚要在版心之前排完：版心顶取 max(w:top, 页眉底)，不知道页眉多高就算不出来
  const header = buildFrame(flow, 'header', number);
  const footer = buildFrame(flow, 'footer', number);
  const margin = (flow.sections[flow.sectionIndex] ?? FALLBACK_SECTION).margin;

  const page: PageLayout = {
    index: flow.pages.length,
    number,
    sectionIndex: flow.sectionIndex,
    geometry: contentWithHeaderFooter(
      flow.geometry,
      margin,
      header?.height ?? 0,
      footer?.height ?? 0,
      flow.headerRules,
    ),
    blocks: [],
    ...(header === undefined ? {} : { header }),
    ...(footer === undefined ? {} : { footer }),
  };
  flow.nextNumber += 1;
  flow.sectionFirstPage = false;
  flow.pages.push(page);
  flow.page = page;
  flow.y = 0;
  return page;
}

// ── 页眉页脚 ──────────────────────────────────────────────────────────────────

/**
 * 哪些部件里有要算的域。
 *
 * 按 run **是不是真的在这份内容里**判断，而不是拿 id 的前缀去猜 —— 前缀（`rId7:`）是
 * `parseHeaderFooter` 的约定，一旦有人换个 id 方案，靠前缀的写法会安静地退化成
 * 「所有页共用第一页的页码」，而那种错在小样本上看不出来。
 */
function dynamicParts(
  source: HeaderFooterSource | undefined,
  plan: HeaderFieldPlan | undefined,
): Set<string> {
  const out = new Set<string>();
  if (source === undefined || plan === undefined || plan.fields.size === 0) return out;
  for (const [relId, content] of Object.entries(source)) {
    for (const b of walkBlocks(content.resolved as ResolvedBlock[])) {
      if (b.kind !== 'paragraph') continue;
      if (b.runs.some((r) => plan.fields.has(r.id))) {
        out.add(relId);
        break;
      }
    }
  }
  return out;
}

function buildFrame(flow: Flow, kind: 'header' | 'footer', number: number): PlacedHeaderFooter | undefined {
  const source = flow.opts.headerFooters;
  if (source === undefined) return undefined;
  const props = flow.sections[flow.sectionIndex];
  if (props === undefined) return undefined;

  const ref = pickHeaderFooter(
    flow.sections,
    flow.sectionIndex,
    kind,
    flow.sectionFirstPage,
    number,
    flow.opts.settings,
  );
  if (ref === undefined) return undefined;
  const content = source[ref.relId];
  if (content === undefined || content.resolved.length === 0) return undefined;

  // 静态页眉全文档共用一份；带页码的那种同一页码也能共用（奇偶页眉在偶数页之间就是同一份）
  const key = flow.hfDynamic.has(ref.relId) ? `${ref.relId}|${number}|${flow.sectionIndex}` : ref.relId;
  let stacked = flow.hf.get(key);
  if (stacked === undefined) {
    const values = headerFieldValues(flow.opts.headerFields, number, flow.sectionIndex, props);
    stacked = stackBlocks(content.resolved, {
      measurer: flow.opts.measurer,
      settings: flow.opts.settings,
      docGrid: props.docGrid,
      contentWidth: flow.geometry.content.width,
      ...(flow.opts.defaultFont === undefined ? {} : { defaultFont: flow.opts.defaultFont }),
      ...(values === undefined ? {} : { fieldValues: values }),
    });
    flow.hf.set(key, stacked);
  }

  return {
    kind,
    relId: ref.relId,
    ...frameOf(kind, flow.geometry, props.margin, stacked.height, flow.headerRules),
    blocks: stacked.blocks,
  };
}

/**
 * 这一页的页眉页脚里，每个域该显示什么。
 *
 * PAGE 在这里是**精确**的 —— 页码在开页那一刻就定了，不像正文的域要靠迭代把上一趟的
 * 结果喂回来。要迭代的只有 NUMPAGES / SECTIONPAGES：它们要的是「一共几页」，
 * 而那个数得整份排完才知道，所以第一趟拿不到（`plan.totalPages` 是 undefined），
 * 那一趟就让文件里存着的旧值先顶着。
 */
function headerFieldValues(
  plan: HeaderFieldPlan | undefined,
  number: number,
  sectionIndex: number,
  props: SectionProps,
): ReadonlyMap<NodeId, string> | undefined {
  if (plan === undefined || plan.fields.size === 0) return undefined;
  const out = new Map<NodeId, string>();
  for (const [id, spec] of plan.fields) {
    switch (spec.type) {
      case 'clear':
        out.set(id, '');
        break;
      case 'PAGE':
        out.set(id, formatNumber(number, spec.format ?? props.pageNumFormat ?? 'decimal'));
        break;
      case 'NUMPAGES':
        if (plan.totalPages !== undefined)
          out.set(id, formatNumber(plan.totalPages, spec.format ?? 'decimal'));
        break;
      default:
        if (plan.sectionPages !== undefined) {
          out.set(id, formatNumber(plan.sectionPages[sectionIndex] ?? 0, spec.format ?? 'decimal'));
        }
    }
  }
  return out;
}

function breakPage(flow: Flow): void {
  flow.page = undefined;
  flow.y = 0;
}

/** 当前页上已经放过东西了吗 —— 「挪到下一页」只有这时候才有意义，否则会空转出一串空页 */
function pageHasContent(flow: Flow): boolean {
  return flow.page !== undefined && flow.page.blocks.length > 0;
}

/**
 * 这一页还剩多高。
 *
 * **必须看开出来的那一页**而不是 `flow.geometry`：页眉页脚已经把版心挤过一道，
 * 用节的纸面几何算会平白多出页眉那么多空间。所以这里顺手把页开出来 ——
 * 段落 / 表格的分页循环里，`breakPage()` 之后紧跟的就是「下一页还剩多高」，
 * 那时页还没开，取不到它自己的版心（页眉进来之前两者恰好相等，这个洞才一直没露出来）。
 * 惰性开页因此只剩这一处例外，也正是它该有的样子：
 * **问了「还剩多高」就说明真的要往里放东西了**。
 */
function availHeight(flow: Flow): Twips {
  return currentPage(flow).geometry.content.height - flow.y;
}

// ── 块的准备 ──────────────────────────────────────────────────────────────────

function prepare(b: ResolvedBlock, section: SectionProps, opts: LayoutDocumentOptions): Prepared {
  const shared = {
    measurer: opts.measurer,
    settings: opts.settings,
    docGrid: section.docGrid,
    ...(opts.defaultFont === undefined ? {} : { defaultFont: opts.defaultFont }),
    ...(opts.fieldValues === undefined ? {} : { fieldValues: opts.fieldValues }),
    ...(opts.objectRules === undefined ? {} : { objectRules: { ...OBJECT_RULES, ...opts.objectRules } }),
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

// ── 浮动对象 ──────────────────────────────────────────────────────────────────

/**
 * 把锚在各行上的浮动对象换算成**纸坐标**。
 *
 * 只处理 `wrap="none"`（items.ts 已经把别的环绕方式退化成内嵌了），也就是
 * 「衬于文字下方 / 浮于文字上方」这一类：印章、水印、红头的花纹。它们不参与文字流，
 * 所以整页排完再算，算错也只是它自己歪，一行文字都不会跟着动。
 *
 * **表格单元格里的浮动对象没做**：格子的纸坐标要连着行高与合并区一起算，
 * 而公文里浮动对象几乎都锚在正文段落上。漏掉的那些不会消失得无声无息 ——
 * 它们仍然在 `LineLayout.floats` 里，将来接上就是多走一层遍历。
 *
 * 参照物的对应关系已经实测，见 `FLOAT_ORIGIN_RULES`。
 */
/**
 * 浮动对象的参照框（`wp:positionH/V @relativeFrom`）各是哪个框 —— **全部实测**。
 *
 * 样本 `spike-image-02`（`pnpm --filter @uw/fidelity spike:image`）：同一张图复制十三份，
 * 八种横向 × 八种纵向配对着取，**偏移一律写 0**，于是量到的 x / y 就是那个框的起点本身，
 * 不用反推。页边距四边故意各不相同（上 25 / 下 20 / 左 30 / 右 20mm）—— 四边一样的话
 * 「版心」「上页边距框」「纸」三个答案会撞在一起，样本就分不开了。十三张图的高各不相同
 * （11…23pt），因为 PDF 只按绘制顺序给图，靠尺寸认人才不依赖那个顺序。
 *
 * ① `page` → 纸的左上角（实测 0, 0）
 * ② `margin` / `column` → 版心左上角（实测 85.05, 70.85；版心算出来是 85.04, 70.87）。
 *    单栏文档里分栏框就是版心，这一份分不开这两个 —— 也不需要分开
 * ③ `leftMargin` → x = 0、`topMargin` → y = 0：**左 / 上页边距框是从纸边起算的**
 * ④ `rightMargin` → x = 版心右边（538.60）、`bottomMargin` → y = 版心底（785.20）：
 *    右 / 下页边距框从版心的边起算，与 ③ 对称
 * ⑤ `insideMargin` / `outsideMargin` 按**显示页码**的奇偶镜像。横向镜像左右页边距
 *    （奇数页 inside = 左、偶数页 inside = 右），**纵向镜像的是上下页边距**：
 *    奇数页 inside = 上（y=0）、偶数页 inside = 下（y=785.20）。
 *    原来纵向退到「版心」，差着整整一个上页边距（这份样本里 70.87pt）
 * ⑥ `character` → 锚点**前一个**字的左边缘。三级阶梯（锚在第 1 / 5 / 9 个字之后）
 *    量到的 x 分别落在第 0 / 4 / 8 个字上，所以不是「锚点那个字」。
 *    ⚠️ 锚在段首（前面一个字都没有）没有样本，那时退到行的左边缘
 * ⑦ `line` → 行顶、`paragraph` → 段顶（各差 0.11 与 0.22pt，是 Word 自己的行位置抖动）
 *
 * 没有样本的只剩「同一根轴上 align（left/center/right）与 offset 并存时谁赢」——
 * 规范里那是个 choice，Word 也只写一个，所以造不出样本。
 */
function placeFloats(page: PageLayout): void {
  const out: PlacedFloat[] = [];
  const g = page.geometry;
  collectFloats(page.blocks, g.content.x, g.content.y, page, out);
  // 页眉页脚里的浮动对象锚在**框**上（框自己的坐标已经是纸坐标了）
  if (page.header !== undefined) collectFloats(page.header.blocks, page.header.x, page.header.y, page, out);
  if (page.footer !== undefined) collectFloats(page.footer.blocks, page.footer.x, page.footer.y, page, out);
  if (out.length === 0) return;
  // 稳定排序：z 相同的按文档顺序，与 Word 「后插入的盖在上面」一致
  out.sort((a, b) => a.z - b.z);
  page.floats = out;
}

function collectFloats(
  blocks: readonly PlacedBlock[],
  originX: Twips,
  originY: Twips,
  page: PageLayout,
  out: PlacedFloat[],
): void {
  for (const block of blocks) {
    if (block.kind !== 'paragraph') continue;
    for (const placed of block.lines) {
      const floats = placed.line.floats;
      if (floats === undefined) continue;
      for (const f of floats) {
        out.push(
          resolveFloat(f, page, {
            originX,
            originY,
            paraTop: originY + block.y,
            lineTop: originY + placed.y,
            lineHeight: placed.line.height,
          }),
        );
      }
    }
  }
}

/** 锚点周围那几个「参照物」的纸坐标。`page` / `margin` 之外的四种都要它 */
interface FloatContext {
  originX: Twips;
  originY: Twips;
  paraTop: Twips;
  lineTop: Twips;
  lineHeight: Twips;
}

function resolveFloat(f: LineFloat, page: PageLayout, ctx: FloatContext): PlacedFloat {
  const out: PlacedFloat = {
    runId: f.runId,
    contentIndex: f.contentIndex,
    x: axisPosition(hBox(f, page, ctx), f.anchor.h, f.width, page.number),
    y: axisPosition(vBox(f, page, ctx), f.anchor.v, f.height, page.number),
    width: f.width,
    height: f.height,
    objectKind: f.objectKind,
    behindDoc: f.anchor.behindDoc,
    z: f.anchor.z,
  };
  if (f.image !== undefined) out.image = f.image;
  if (f.alt !== undefined) out.alt = f.alt;
  if (f.graphic !== undefined) out.graphic = f.graphic;
  return out;
}

/** 一个方向上的参照框：从哪儿起、有多长 */
interface AxisBox {
  start: Twips;
  size: Twips;
}

function hBox(f: LineFloat, page: PageLayout, ctx: FloatContext): AxisBox {
  const g = page.geometry;
  const content = g.content;
  const rightMargin: AxisBox = {
    start: content.x + content.width,
    size: g.width - content.x - content.width,
  };
  const leftMargin: AxisBox = { start: 0, size: content.x };
  const odd = page.number % 2 === 1;
  switch (f.anchor.h.relativeFrom) {
    case 'page':
      return { start: 0, size: g.width };
    case 'leftMargin':
      return leftMargin;
    case 'rightMargin':
      return rightMargin;
    case 'insideMargin':
      return odd ? leftMargin : rightMargin;
    case 'outsideMargin':
      return odd ? rightMargin : leftMargin;
    case 'character':
      // 参照的是锚点**前一个**字的左边缘（实测，见 `FLOAT_ORIGIN_RULES` ⑥ 与 `LineFloat.anchorX`）。
      // 宽度为 0，所以 align 的居中 / 右对齐在这里退化成同一个点
      return { start: ctx.originX + (f.anchorX ?? f.x), size: 0 };
    default:
      // margin / column / 认不出的：单栏文档里分栏框就是版心
      return { start: content.x, size: content.width };
  }
}

function vBox(f: LineFloat, page: PageLayout, ctx: FloatContext): AxisBox {
  const g = page.geometry;
  const content = g.content;
  const topMargin: AxisBox = { start: 0, size: content.y };
  const bottomMargin: AxisBox = {
    start: content.y + content.height,
    size: g.height - content.y - content.height,
  };
  const odd = page.number % 2 === 1;
  switch (f.anchor.v.relativeFrom) {
    case 'page':
      return { start: 0, size: g.height };
    case 'topMargin':
      return topMargin;
    case 'bottomMargin':
      return bottomMargin;
    // 纵向的 inside / outside 镜像的是**上下**页边距，不是版心（实测，见 `FLOAT_ORIGIN_RULES` ⑤）。
    // 原来退到版心，差着整整一个上页边距（样本里 70.87pt）
    case 'insideMargin':
      return odd ? topMargin : bottomMargin;
    case 'outsideMargin':
      return odd ? bottomMargin : topMargin;
    case 'paragraph':
      return { start: ctx.paraTop, size: 0 };
    case 'line':
      return { start: ctx.lineTop, size: ctx.lineHeight };
    default:
      // margin / 认不出的
      return { start: content.y, size: content.height };
  }
}

/**
 * 偏移与对齐**二选一**（规范里是 choice）：两个都没有时落在参照框的起点。
 *
 * `inside` / `outside` 在**奇数页**是「左 / 右」，偶数页反过来 —— 这是装订成册的书页方向，
 * 与页边距的镜像同源。用的是**显示页码**而不是物理页序：`w:pgNumType w:start` 让某一节
 * 从 5 起算时，Word 的奇偶是跟着页码走的（与页眉的奇偶同一条判据，见 header-footer.ts）。
 */
function axisPosition(box: AxisBox, pos: AnchorPos, size: Twips, pageNumber: number): Twips {
  if (pos.offset !== undefined) return box.start + pos.offset;
  const odd = pageNumber % 2 === 1;
  switch (pos.align) {
    case 'center':
      return box.start + (box.size - size) / 2;
    case 'right':
    case 'bottom':
      return box.start + box.size - size;
    case 'inside':
      return odd ? box.start : box.start + box.size - size;
    case 'outside':
      return odd ? box.start + box.size - size : box.start;
    default:
      // left / top / 认不出的对齐值
      return box.start;
  }
}
