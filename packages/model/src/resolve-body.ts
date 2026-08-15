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
import type { ParaProps, ResolvedParaProps, ResolvedRunProps, RunProps } from './props.ts';

export function resolveBody(ctx: CascadeContext, body: Body): ResolvedBody {
  return {
    sections: body.sections.map((s): ResolvedSection => {
      // 节属性本来就是解析完的纯数据，直接带过来；克隆是为了断开与可编辑树的共享，
      // 否则编辑那边改一个页边距，已经发给 Worker 的这棵树会跟着变（或者反过来）
      return { id: s.id, props: structuredClone(s.props), blocks: s.blocks.map((b) => block(ctx, b)) };
    }),
  };
}

function block(ctx: CascadeContext, b: Block): ResolvedBlock {
  return b.kind === 'paragraph' ? paragraph(ctx, b) : table(ctx, b);
}

function paragraph(ctx: CascadeContext, p: ParagraphNode<ParaProps, RunProps>): ResolvedParagraph {
  // 段落的直接 pPr 要同时喂给字符级联 —— 段落样式链上的 rPr 是字符属性的一层，
  // 而 ResolvedParaProps 里已经没有「段落样式 id 之外的原始信息」了
  const props: ResolvedParaProps = resolveParaProps(ctx, p.props);
  return {
    kind: 'paragraph',
    id: p.id,
    props,
    runs: p.runs.map((r): ResolvedRun => {
      const rp: ResolvedRunProps = resolveRunProps(ctx, p.props, r.props);
      const out: ResolvedRun = { kind: 'run', id: r.id, props: rp, content: structuredClone(r.content) };
      if (r.hyperlink !== undefined) out.hyperlink = { ...r.hyperlink };
      return out;
    }),
  };
}

function table(
  ctx: CascadeContext,
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
            blocks: c.blocks.map((b) => block(ctx, b)),
          };
        }),
      };
    }),
  };
}

// 表格样式的条件格式（首行 / 末列 / 隔行底纹）那一层还没有位置，Phase 4 补：
// 它插在段落样式链**之前**，且要知道单元格在表里的行列号 —— 也就是说届时
// `table()` 这条路径要往下传一个「单元格在表中的位置」，不只是无脑递归。
