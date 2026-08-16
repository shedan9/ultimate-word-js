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
import { gridColumnCount, resolveCellProps, resolveRowProps, resolveTableProps } from './cascade-table.ts';
import type {
  Block,
  Body,
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

function paragraph(pass: Pass, p: Paragraph): ResolvedParagraph {
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
      const rowProps = resolveRowProps(pass.ctx, props, t.props, r.props, {
        row: rowIndex,
        rowCount,
      });
      // 本行被 w:gridBefore 跳掉的那几列也占位置，列号要从它之后开始数
      let col = rowProps.gridBefore;
      return {
        kind: 'row',
        id: r.id,
        props: rowProps,
        cells: r.cells.map((c): ResolvedTableCell => {
          const pos: CellPosition = { row: rowIndex, rowCount, col, span: c.gridSpan, colCount };
          col += c.gridSpan;
          const { props: cellProps, layers } = resolveCellProps(pass.ctx, props, t.props, c.props, pos);
          // 格内的段落 / run 走同一条级联，只是多了这几层前置样式
          const inner: Pass = { ...pass, ctx: { ...pass.ctx, tableStyleLayers: layers } };
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
