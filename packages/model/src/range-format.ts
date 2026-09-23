import type { ResolvedBody } from './nodes.ts';
import { walkParagraphs } from './nodes.ts';
import { buildRunOrder, compareDocPositions, runEnd, runStart } from './order.ts';
import type { DocRange } from './position.ts';
import type { ResolvedParaProps, ResolvedRunProps } from './props.ts';

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

/** 选区触及的每个段落（含其间的单元格段落）的级联段落格式；折叠光标即所在段。 */
export function paraPropsOfRange(body: ResolvedBody, range: DocRange): ResolvedParaProps[] {
  const paragraphs = [...walkParagraphs(body)];
  const indexOf = (nodeId: string) => {
    const i = paragraphs.findIndex((p) => p.id === nodeId || p.runs.some((r) => r.id === nodeId));
    if (i < 0) throw new RangeError('选区不在正文中');
    return i;
  };
  const a = indexOf(range.start.nodeId);
  const b = indexOf(range.end.nodeId);
  return paragraphs.slice(Math.min(a, b), Math.max(a, b) + 1).map((p) => p.props);
}
