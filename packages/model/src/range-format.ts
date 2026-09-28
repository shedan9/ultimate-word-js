import type { CascadeContext } from './cascade.ts';
import type { Body, ResolvedBody, ResolvedParagraph } from './nodes.ts';
import { walkParagraphs } from './nodes.ts';
import { buildRunOrder, compareDocPositions, runEnd, runStart } from './order.ts';
import type { DocPosition, DocRange } from './position.ts';
import type { ResolvedRunProps } from './props.ts';
import { resolveBody } from './resolve-body.ts';
import type { RunPropsPatch } from './text-transaction.ts';

/**
 * 暂存格式的查询必须重新级联：null 是回到样式，不能被当作 false，也不能再读旧的直接格式。
 * 在副本上应用，不创建事务 / 历史。复用整棵树的级联，以保留单元格条件样式与本次新建样式。
 * 只在带暂存格式的切换命令里调用，不进入输入与布局的热路径。
 */
export function runPropsAtInsertion(
  ctx: CascadeContext,
  body: Body,
  position: DocPosition,
  patch: RunPropsPatch,
): ResolvedRunProps[] {
  const preview = structuredClone(body);
  for (const paragraph of walkParagraphs(preview)) {
    if (paragraph.id !== position.nodeId && !paragraph.runs.some((r) => r.id === position.nodeId)) continue;
    // 字缝沿用左 run、空段看标记；给所在段的候选都打补丁，具体选择仍由 runPropsOfRange 决定。
    const candidates = paragraph.runs.map((r) => r.props);
    if (!candidates.length) {
      paragraph.props.markRunProps ??= {};
      candidates.push(paragraph.props.markRunProps);
    }
    for (const props of candidates) {
      for (const [key, value] of Object.entries(patch)) {
        if (value === null) delete props[key as keyof typeof props];
        else if (value !== undefined) Object.assign(props, { [key]: value });
      }
    }
    break;
  }
  return runPropsOfRange(resolveBody(ctx, preview), { start: position, end: position });
}

/**
 * 选区里的字符格式，给工具栏状态与 Ctrl+B 这类「全是才取消」的切换用。
 *
 * 非折叠范围取与之有交集的 run 与被覆盖的空段落标记；折叠光标取**将要输入**时的格式：
 * 空段落看段落标记，run 开头看前一个 run（Word 在字缝上沿用左边的字），否则看所在 run。
 * 吃级联完的树：加粗可能来自样式，只看直接格式会把标题里的 Ctrl+B 判成「加粗」。
 * 隐藏文字不显示，不参与「是否全部加粗」的判断。
 */
export function runPropsOfRange(body: ResolvedBody, range: DocRange): ResolvedRunProps[] {
  const order = buildRunOrder(body);
  const start = range.start;
  const end = range.end;
  const cmp = (a: typeof start, b: typeof start) => {
    const r = compareDocPositions(order, a, b);
    if (r === undefined) throw new RangeError('选区不在正文中');
    return r;
  };
  const collapsed = cmp(start, end) === 0;
  if (cmp(start, end) > 0) return runPropsOfRange(body, { start: end, end: start });
  const out: ResolvedRunProps[] = [];
  for (const paragraph of walkParagraphs(body)) {
    if (!paragraph.runs.length) {
      const at = { nodeId: paragraph.id, contentIndex: 0, offset: 0 };
      if (cmp(start, at) <= 0 && cmp(at, end) <= 0) out.push(paragraph.props.markRunProps);
      continue;
    }
    for (const [i, run] of paragraph.runs.entries()) {
      if (collapsed) {
        if (run.id !== start.nodeId) continue;
        const previous = paragraph.runs[i - 1];
        out.push(cmp(start, runStart(run)) === 0 && previous ? previous.props : run.props);
        return out;
      }
      if (
        !run.props.hidden &&
        run.content.length &&
        cmp(runEnd(run), start) > 0 &&
        cmp(runStart(run), end) < 0
      )
        out.push(run.props);
    }
  }
  return out;
}

/**
 * 选区触及的每个段落（含其间的单元格段落），级联完的节点；折叠光标即所在段。
 * 返回节点而不只是属性：列表升降级要按段各算各的层级，得知道每段在哪。
 */
export function paragraphsOfRange(body: ResolvedBody, range: DocRange): ResolvedParagraph[] {
  const paragraphs = [...walkParagraphs(body)];
  const indexOf = (nodeId: string) => {
    const i = paragraphs.findIndex((p) => p.id === nodeId || p.runs.some((r) => r.id === nodeId));
    if (i < 0) throw new RangeError('选区不在正文中');
    return i;
  };
  const a = indexOf(range.start.nodeId);
  const b = indexOf(range.end.nodeId);
  return paragraphs.slice(Math.min(a, b), Math.max(a, b) + 1);
}
