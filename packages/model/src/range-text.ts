import type { NodeId, ResolvedBody, RunContent } from './nodes.ts';
import { walkParagraphs } from './nodes.ts';
import { buildRunOrder, compareDocPositions, contentLength, rangeOfNode } from './order.ts';
import type { DocPosition, DocRange } from './position.ts';
import type { Justification, ResolvedRunProps } from './props.ts';

/** 选区里的一段文字与它**级联后**的格式：剪贴板的另一头不认识我们的样式表，只能带最终值。 */
export interface FragmentRun {
  text: string;
  props: ResolvedRunProps;
}
/** 选区触及的一个段落；首末段只含选中的那部分文字，段落属性只带对齐（其余排版属性依赖目标文档的网格与样式）。 */
export interface FragmentParagraph {
  justification: Justification;
  runs: FragmentRun[];
}
/** 富文本复制的中间结果：纯数据，HTML 的生成在 `@uw/view`（model 不碰 DOM 与 CSS）。 */
export interface RichFragment {
  paragraphs: FragmentParagraph[];
}

/** 复制模型选区：自动折行不换行，段落（包括空段与单元格内段落）之间保留一个换行。 */
export function textOfRange(
  body: ResolvedBody,
  range: DocRange,
  fieldValues?: ReadonlyMap<NodeId, string>,
): string {
  return fragmentOfRange(body, range, fieldValues)
    .paragraphs.map((p) => p.runs.map((r) => r.text).join(''))
    .join('\n');
}

/**
 * 选区的带格式片段，与 `textOfRange` 同一套取舍（跳过隐藏文字与域代码、域取当前显示值），
 * 纯文本就是把它拼起来 —— 两份剪贴板内容不会各说各话。折叠选区答空片段（零个段落）。
 */
export function fragmentOfRange(
  body: ResolvedBody,
  range: DocRange,
  fieldValues?: ReadonlyMap<NodeId, string>,
): RichFragment {
  const paragraphs = [...walkParagraphs(body)];
  const order = buildRunOrder(body);
  const compare = (a: DocPosition, b: DocPosition) => {
    const result = compareDocPositions(order, a, b);
    if (result === undefined) throw new RangeError('选区不在正文中');
    return result;
  };
  function paragraphAt(at: DocPosition): number {
    const index = paragraphs.findIndex((p) => p.id === at.nodeId || p.runs.some((r) => r.id === at.nodeId));
    const paragraph = paragraphs[index];
    const run = paragraph?.runs.find((r) => r.id === at.nodeId);
    const content = run?.content[at.contentIndex];
    const empty = paragraph && (run ? !run.content.length : !paragraph.runs.length);
    if (
      !Number.isInteger(at.contentIndex) ||
      !Number.isInteger(at.offset) ||
      (empty
        ? at.contentIndex !== 0 || at.offset !== 0
        : !content || at.offset < 0 || at.offset > contentLength(content))
    )
      throw new RangeError('无效的文字位置');
    return index;
  }
  const first = paragraphAt(range.start);
  const last = paragraphAt(range.end);
  const backwards = compare(range.start, range.end) > 0;
  const start = backwards ? range.end : range.start;
  const end = backwards ? range.start : range.end;
  if (compare(start, end) === 0) return { paragraphs: [] };
  const selected = paragraphs.slice(Math.min(first, last), Math.max(first, last) + 1).map((paragraph) => {
    const runs: FragmentRun[] = [];
    const push = (text: string, props: ResolvedRunProps) => {
      if (!text) return;
      const previous = runs.at(-1);
      // 同一 run 的几个内容片（文字 + 制表位）拼回一段，下游不必逐片比格式。
      if (previous?.props === props) previous.text += text;
      else runs.push({ text, props });
    };
    for (const run of paragraph.runs) {
      if (run.props.hidden) continue;
      if (fieldValues?.has(run.id)) {
        // 域显示值没有逐字源位置；选区覆盖它时复制当前求值结果，不能泄漏文件里的旧页码。
        const bounds = rangeOfNode(run);
        if (
          bounds &&
          (compare(start, bounds.end) < 0 ||
            (compare(bounds.start, bounds.end) === 0 && compare(start, bounds.start) <= 0)) &&
          compare(bounds.start, end) < 0
        )
          push(fieldValues.get(run.id) ?? '', run.props);
        continue;
      }
      for (const [contentIndex, content] of run.content.entries()) {
        const at = { nodeId: run.id, contentIndex, offset: 0 };
        const length = contentLength(content);
        if (compare(start, { ...at, offset: length }) >= 0 || compare(end, at) <= 0) continue;
        const from = start.nodeId === run.id && start.contentIndex === contentIndex ? start.offset : 0;
        const to = end.nodeId === run.id && end.contentIndex === contentIndex ? end.offset : length;
        push(contentText(content).slice(from, to), run.props);
      }
    }
    return { justification: paragraph.props.justification, runs };
  });
  return { paragraphs: selected };
}

function contentText(content: RunContent): string {
  switch (content.kind) {
    case 'text':
      return content.text;
    case 'tab':
      return '\t';
    case 'break':
      return '\n';
    case 'symbol':
      return String.fromCodePoint(content.char.codePointAt(0) ?? 0xfffd);
    case 'noBreakHyphen':
      return '-';
    case 'object':
      return '\uFFFC';
    default:
      return '';
  }
}
