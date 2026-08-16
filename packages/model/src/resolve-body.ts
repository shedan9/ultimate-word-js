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
import type {
  Block,
  Body,
  ParagraphNode,
  ResolvedBlock,
  ResolvedBody,
  ResolvedParagraph,
  ResolvedRun,
  ResolvedSection,
  TableCellNode,
  TableNode,
  TableRowNode,
} from './nodes.ts';
import type { NumberingCounters } from './numbering-counter.ts';
import { createNumberingCounters } from './numbering-counter.ts';
import type { ParaProps, ResolvedParaProps, ResolvedRunProps, RunProps } from './props.ts';

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
}

export function resolveBody(ctx: CascadeContext, body: Body): ResolvedBody {
  const pass: Pass = { ctx, counters: createNumberingCounters(ctx.numbering, ctx.styles) };
  return {
    sections: body.sections.map((s): ResolvedSection => {
      // 节属性本来就是解析完的纯数据，直接带过来；克隆是为了断开与可编辑树的共享，
      // 否则编辑那边改一个页边距，已经发给 Worker 的这棵树会跟着变（或者反过来）
      return { id: s.id, props: structuredClone(s.props), blocks: s.blocks.map((b) => block(pass, b)) };
    }),
  };
}

function block(pass: Pass, b: Block): ResolvedBlock {
  return b.kind === 'paragraph' ? paragraph(pass, b) : table(pass, b);
}

function paragraph(pass: Pass, p: ParagraphNode<ParaProps, RunProps>): ResolvedParagraph {
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
      if (r.hyperlink !== undefined) out.hyperlink = { ...r.hyperlink };
      return out;
    }),
  };
}

/**
 * 表格里的段落**参与同一条编号计数**：Word 里表格单元格中的列表和正文里的列表
 * 共用一个编号实例时是连着数的。所以这里递归下去的是同一个 `pass`，不是新建一个。
 */
function table(
  pass: Pass,
  t: TableNode<ParaProps, RunProps>,
): TableNode<ResolvedParaProps, ResolvedRunProps> {
  return {
    kind: 'table',
    id: t.id,
    rows: t.rows.map((r): TableRowNode<ResolvedParaProps, ResolvedRunProps> => {
      return {
        kind: 'row',
        id: r.id,
        cells: r.cells.map((c): TableCellNode<ResolvedParaProps, ResolvedRunProps> => {
          return {
            kind: 'cell',
            id: c.id,
            gridSpan: c.gridSpan,
            vMerge: c.vMerge,
            blocks: c.blocks.map((b) => block(pass, b)),
          };
        }),
      };
    }),
  };
}

// 表格样式的条件格式（首行 / 末列 / 隔行底纹）那一层还没有位置，Phase 4 补：
// 它插在段落样式链**之前**，且要知道单元格在表里的行列号 —— 也就是说届时
// `table()` 这条路径要往下传一个「单元格在表中的位置」，不只是无脑递归。
