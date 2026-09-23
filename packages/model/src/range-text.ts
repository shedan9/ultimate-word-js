import type { NodeId, ResolvedBody, RunContent } from './nodes.ts';
import { walkParagraphs } from './nodes.ts';
import { buildRunOrder, compareDocPositions, contentLength, rangeOfNode } from './order.ts';
import type { DocPosition, DocRange } from './position.ts';

/** 复制模型选区：自动折行不换行，段落（包括空段与单元格内段落）之间保留一个换行。 */
export function textOfRange(
  body: ResolvedBody,
  range: DocRange,
  fieldValues?: ReadonlyMap<NodeId, string>,
): string {
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
  if (compare(start, end) === 0) return '';
  return paragraphs
    .slice(Math.min(first, last), Math.max(first, last) + 1)
    .map((paragraph) => {
      let text = '';
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
            text += fieldValues.get(run.id);
          continue;
        }
        for (const [contentIndex, content] of run.content.entries()) {
          const at = { nodeId: run.id, contentIndex, offset: 0 };
          const length = contentLength(content);
          if (compare(start, { ...at, offset: length }) >= 0 || compare(end, at) <= 0) continue;
          const from = start.nodeId === run.id && start.contentIndex === contentIndex ? start.offset : 0;
          const to = end.nodeId === run.id && end.contentIndex === contentIndex ? end.offset : length;
          text += contentText(content).slice(from, to);
        }
      }
      return text;
    })
    .join('\n');
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
