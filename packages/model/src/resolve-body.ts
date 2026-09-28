/**
 * 节点树（直接格式）→ 节点树（级联完的属性）。
 *
 * 这一步就是 **Worker 边界**：输入这边还需要 `CascadeContext`（`StyleSheet` 带方法，
 * 不可结构化克隆），输出那边是一棵纯数据树，可以整个 `postMessage` 给布局（原则 1.1）。
 * 所以级联**必须**在这里做完，不能留给 `@uw/layout` 边走边算。
 *
 * 形状一比一保留：id 不变、顺序不变、层级不变。只有 `props` 换了类型。
 * 这样 `DocPosition{nodeId}` 在两棵树上都能查，布局结果也能反查回可编辑的那棵。
 */
import type { CascadeContext } from './cascade.ts';
import { resolveParaProps, resolveRunProps } from './cascade.ts';
import type { CellPosition } from './cascade-table.ts';
import {
  applyRowExceptions,
  gridColumnCount,
  resolveCellProps,
  resolveRowProps,
  resolveTableProps,
} from './cascade-table.ts';
import type { FieldHyperlink } from './fields.ts';
import type {
  Block,
  Body,
  NodeId,
  Paragraph,
  ResolvedBlock,
  ResolvedBody,
  ResolvedParagraph,
  ResolvedRun,
  ResolvedSection,
  ResolvedTable,
  ResolvedTableCell,
  ResolvedTableRow,
  Table,
} from './nodes.ts';
import type { NumberingCounters } from './numbering-counter.ts';
import { createNumberingCounters } from './numbering-counter.ts';
import type { ResolvedParaProps, ResolvedRunProps } from './props.ts';
import { extendStyleSheet } from './styles.ts';

/**
 * 一趟级联的全部状态。
 *
 * `counters` 是**有状态**的那一半：编号「第几」只有按文档顺序走一遍才知道。
 * 它在这里创建、随遍历往下传，因此一次 `resolveBody` 就是一份干净的计数 ——
 * 同一份文档重解析两次结果相同（原则：级联结果是派生量，不许跨调用残留）。
 */
interface Pass {
  ctx: CascadeContext;
  counters: NumberingCounters;
  hyperlinks: ReadonlyMap<NodeId, FieldHyperlink> | undefined;
  /** 跨调用的备忘；缺席 = 每段都重算（加载那一趟与低层调用） */
  memo: WeakMap<Paragraph, MemoEntry> | undefined;
  /** 本段所在单元格命中的表格样式层，序列化成串供备忘比对；正文里是空串 */
  layersKey: string;
}

interface MemoEntry {
  resolved: ResolvedParagraph;
  layersKey: string;
}

/**
 * 跨 `resolveBody` 调用复用段落的级联结果（`createResolveCache()` 造，调用方只管传）。
 *
 * 为什么需要它：打一个字只改一个段落，但级联要整份重跑，结果又是一棵全新的树 ——
 * 下游的段落缓存只能按**值**比，于是每一趟都把全文每段约 3.6KB 的级联属性序列化一遍
 * 当键（54 页的文档每次按键 9ms，比排版本身还贵）。复用了对象，下游就能按**身份**认出
 * 「这一段没变」。
 *
 * 复用的判据，缺一不可：
 * 1. **源段落已冻结且是同一个对象** —— 编辑器的快照深冻结、按修改路径共享子树，没改的
 *    段落跨事务就是同一个对象；可写的段落可能被原地改过，一律重算
 * 2. **上下文的来源没变**：级联上下文、树上的编号定义与补充样式（新建列表 / 套标题会换掉它们，
 *    换了就整份作废 —— 它们能改任何一段的结果）、所在单元格命中的样式层
 * 3. **上次没有编号**（numId = 0）：编号「第几」取决于前文，前面插一条列表项，
 *    后面每一项的文字都变；而 numId 由样式链与直接格式定，前两条成立时它也不会变，
 *    所以「上次 numId = 0」就说明这次也不推进计数器、跳过它不会让后文数错
 * 4. **HYPERLINK 域给每个 run 的链接没变**：域跨段落，别处改了界桩也能改这一段的链接
 *
 * 复用出去的结果**深冻结**：它同时挂在前后两棵级联树上，谁改了都会串到另一份。
 */
export interface ResolveCache {
  readonly __brand: 'ResolveCache';
}

interface ResolveCacheState extends ResolveCache {
  context: CascadeContext | undefined;
  numbering: Body['numbering'];
  styles: Body['styles'];
  paragraphs: WeakMap<Paragraph, MemoEntry>;
}

export function createResolveCache(): ResolveCache {
  const state: ResolveCacheState = {
    __brand: 'ResolveCache',
    context: undefined,
    numbering: undefined,
    styles: undefined,
    paragraphs: new WeakMap(),
  };
  return state;
}

/** 上下文的来源换了就整份作废，返回这一趟可用的备忘 */
function memoFor(
  cache: ResolveCache | undefined,
  context: CascadeContext,
  body: Body,
): WeakMap<Paragraph, MemoEntry> | undefined {
  if (cache === undefined) return undefined;
  const state = cache as ResolveCacheState;
  if (state.context !== context || state.numbering !== body.numbering || state.styles !== body.styles) {
    state.context = context;
    state.numbering = body.numbering;
    state.styles = body.styles;
    state.paragraphs = new WeakMap();
  }
  return state.paragraphs;
}

export interface ResolveBodyOptions {
  /**
   * HYPERLINK 域算出来的链接（`fieldHyperlinks(scanFields(body))`），按 run id 铺到结果 run 上。
   *
   * 为什么不在这里现扫：域**跨段落**，而这一趟是按段落递归下去的，扫不出跨段的配对。
   * 由调用方（`loadDocument`）先扫一遍整份 body 再传进来，顺序上也对 ——
   * 扫描要的是**直接格式**那棵树，级联改不了界桩的位置。
   */
  hyperlinks?: ReadonlyMap<NodeId, FieldHyperlink>;
  /** 跨调用复用没变的段落（见 `ResolveCache`）。门面每份文档持有一份 */
  cache?: ResolveCache;
}

export function resolveBody(
  context: CascadeContext,
  body: Body,
  opts: ResolveBodyOptions = {},
): ResolvedBody {
  // 树上带着定义时以它为准：编辑期新建的列表只在这份里，级联上下文还是加载时那份；
  // 新增的样式同理（样式表本身不可变，补一张扩展表）
  const ctx = {
    ...context,
    ...(body.numbering === undefined ? {} : { numbering: body.numbering }),
    ...(body.styles?.length ? { styles: extendStyleSheet(context.styles, body.styles) } : {}),
  };
  const pass: Pass = {
    ctx,
    counters: createNumberingCounters(ctx.numbering, ctx.styles),
    hyperlinks: opts.hyperlinks,
    memo: memoFor(opts.cache, context, body),
    layersKey: '',
  };
  return {
    sections: body.sections.map((s): ResolvedSection => {
      // 节属性本来就是解析完的纯数据，直接带过来；克隆是为了断开与可编辑树的共享，
      // 否则编辑那边改一个页边距，已经发给 Worker 的这棵树会跟着变（或者反过来）
      return { id: s.id, props: structuredClone(s.props), blocks: s.blocks.map((b) => block(pass, b)) };
    }),
  };
}

/**
 * 一列块的级联 —— 页眉页脚用的入口（它们没有节，只有块）。
 *
 * **计数器是新的一份**：编号在页眉里几乎不出现，但真出现时它绝不该推动正文的
 * 「第几条」往前走 —— 页眉会在每一页重排一遍，跟着推的话正文编号会随页数漂移。
 */
export function resolveBlocks(
  ctx: CascadeContext,
  blocks: readonly Block[],
  opts: ResolveBodyOptions = {},
): ResolvedBlock[] {
  // 页眉页脚不走备忘：它们只在加载时级联一次
  const pass: Pass = {
    ctx,
    counters: createNumberingCounters(ctx.numbering, ctx.styles),
    hyperlinks: opts.hyperlinks,
    memo: undefined,
    layersKey: '',
  };
  return blocks.map((b) => block(pass, b));
}

function block(pass: Pass, b: Block): ResolvedBlock {
  return b.kind === 'paragraph' ? paragraph(pass, b) : table(pass, b);
}

function paragraph(pass: Pass, p: Paragraph): ResolvedParagraph {
  const memo = pass.memo !== undefined && Object.isFrozen(p) ? pass.memo : undefined;
  const hit = memo?.get(p);
  if (hit !== undefined && reusable(pass, p, hit)) return hit.resolved;
  const out = resolveParagraph(pass, p);
  memo?.set(p, { resolved: deepFreeze(out), layersKey: pass.layersKey });
  return out;
}

function reusable(pass: Pass, p: Paragraph, hit: MemoEntry): boolean {
  if (hit.layersKey !== pass.layersKey || hit.resolved.props.numbering.numId !== 0) return false;
  return p.runs.every((r, i) => {
    // 容器链接跟着源 run 走，源没变它就没变；只有域给的链接要对
    if (r.hyperlink !== undefined) return true;
    const now = pass.hyperlinks?.get(r.id);
    const before = hit.resolved.runs[i]?.hyperlink;
    return now?.url === before?.url && now?.anchor === before?.anchor;
  });
}

function deepFreeze<T>(value: T): T {
  if (typeof value !== 'object' || value === null || Object.isFrozen(value)) return value;
  for (const child of Object.values(value)) deepFreeze(child);
  return Object.freeze(value);
}

function resolveParagraph(pass: Pass, p: Paragraph): ResolvedParagraph {
  const ctx = pass.ctx;
  // 段落的直接 pPr 要同时喂给字符级联 —— 段落样式链上的 rPr 是字符属性的一层，
  // 而 ResolvedParaProps 里已经没有「段落样式 id 之外的原始信息」了。
  // 计数器一段只能推进一次，所以整棵树里只有这一处传它
  const props: ResolvedParaProps = resolveParaProps(ctx, p.props, pass.counters);
  return {
    kind: 'paragraph',
    id: p.id,
    props,
    // 正文 run 不吃编号的 rPr（那份只作用于编号文字，见 cascade.ts 文件头第 3 条）
    runs: p.runs.map((r): ResolvedRun => {
      const rp: ResolvedRunProps = resolveRunProps(ctx, p.props, r.props);
      const out: ResolvedRun = { kind: 'run', id: r.id, props: rp, content: structuredClone(r.content) };
      // 容器（`w:hyperlink`）优先于 HYPERLINK 域：两者同时罩着一个 run 不合法，
      // 真遇上时听那个**写在正文结构里**的，域是派生量
      const link = r.hyperlink ?? pass.hyperlinks?.get(r.id);
      if (link !== undefined) out.hyperlink = { ...link };
      if (r.fieldSimple !== undefined) out.fieldSimple = { ...r.fieldSimple };
      return out;
    }),
  };
}

/**
 * 表格。
 *
 * 两件事必须在这一层做，因为只有这里同时知道「表格样式」和「单元格在表里的位置」：
 *
 * 1. **条件格式的命中**（首行 / 末列 / 隔行带）要行列号，所以这里逐格数列号 ——
 *    数的是**网格列**（累加 `gridSpan`），不是第几个 `w:tc`
 * 2. 单元格命中的那些层要**派生一个带层的 ctx** 交给格内段落用，出了这个格就没了
 *
 * 表格里的段落**参与同一条编号计数**：Word 里单元格中的列表和正文里的列表
 * 共用一个编号实例时是连着数的。所以递归下去的是同一个 `pass.counters`。
 */
function table(pass: Pass, t: Table): ResolvedTable {
  const props = resolveTableProps(pass.ctx, t.props);
  const rowCount = t.rows.length;
  const colCount = gridColumnCount(t.grid, t.rows);

  return {
    kind: 'table',
    id: t.id,
    props,
    grid: [...t.grid],
    rows: t.rows.map((r, rowIndex): ResolvedTableRow => {
      // 本行的 `w:tblPrEx` 盖在表级结果上，只影响这一行的行 / 格（见 applyRowExceptions）
      const rowTable = r.propsEx === undefined ? props : applyRowExceptions(props, r.propsEx);
      const resolvedRow = resolveRowProps(pass.ctx, rowTable, t.props, r.props, {
        row: rowIndex,
        rowCount,
      });
      // 例外改过的表级边框要带到布局层去：冲突解析的「退到表级」对这一行说的是它
      const rowProps =
        r.propsEx === undefined ? resolvedRow : { ...resolvedRow, tableBorders: rowTable.borders };
      // 本行被 w:gridBefore 跳掉的那几列也占位置，列号要从它之后开始数
      let col = rowProps.gridBefore;
      return {
        kind: 'row',
        id: r.id,
        props: rowProps,
        cells: r.cells.map((c): ResolvedTableCell => {
          const pos: CellPosition = { row: rowIndex, rowCount, col, span: c.gridSpan, colCount };
          col += c.gridSpan;
          const { props: cellProps, layers } = resolveCellProps(pass.ctx, rowTable, t.props, c.props, pos);
          // 格内的段落 / run 走同一条级联，只是多了这几层前置样式
          const inner: Pass = {
            ...pass,
            ctx: { ...pass.ctx, tableStyleLayers: layers },
            // 层按格的位置命中（首行 / 隔行带），插一行就能换掉下面每一格的层
            layersKey: pass.memo === undefined ? '' : JSON.stringify(layers),
          };
          return {
            kind: 'cell',
            id: c.id,
            props: cellProps,
            gridSpan: c.gridSpan,
            vMerge: c.vMerge,
            blocks: c.blocks.map((b) => block(inner, b)),
          };
        }),
      };
    }),
  };
}
