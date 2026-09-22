import type { DocPosition, DocRange, Paragraph } from '@uw/model';

export type TextGranularity = 'grapheme' | 'word';

const graphemes = new Intl.Segmenter('zh', { granularity: 'grapheme' });
const words = new Intl.Segmenter('zh', { granularity: 'word' });

/** 先拼段落再分词；样式与链接边界不能把一个词或 emoji 拆成两次移动。 */
export function paragraphNavigation(paragraph: Paragraph) {
  let text = '';
  const spans: { start: number; end: number; position: DocPosition }[] = [];
  for (const run of paragraph.runs) {
    if (!run.content.length)
      spans.push({
        start: text.length,
        end: text.length,
        position: { nodeId: run.id, contentIndex: 0, offset: 0 },
      });
    for (const [contentIndex, content] of run.content.entries()) {
      if (content.kind !== 'text') {
        // 不可编辑的对象、制表位与域界桩隔断词语，也不产生虚构的文字位置。
        text += '\ufffc';
        continue;
      }
      const start = text.length;
      text += content.text;
      spans.push({ start, end: text.length, position: { nodeId: run.id, contentIndex, offset: 0 } });
    }
  }
  if (!paragraph.runs.length)
    spans.push({ start: 0, end: 0, position: { nodeId: paragraph.id, contentIndex: 0, offset: 0 } });

  function offsetOf(at: DocPosition): number | undefined {
    const span = spans.find(
      (s) => s.position.nodeId === at.nodeId && s.position.contentIndex === at.contentIndex,
    );
    return span && Number.isInteger(at.offset) && at.offset >= 0 && at.offset <= span.end - span.start
      ? span.start + at.offset
      : undefined;
  }
  function positionAt(offset: number, end: boolean): DocPosition | undefined {
    // 同一条字缝有两个模型位置：范围起点取后片段，终点取前片段，避免带入空 run。
    const candidates = spans.filter((s) => offset >= s.start && offset <= s.end);
    const span = end
      ? (candidates.find((s) => s.start < offset) ?? candidates[0])
      : (candidates.findLast((s) => s.end > offset) ?? candidates.at(-1));
    return span && { ...span.position, offset: offset - span.start };
  }
  return {
    move(at: DocPosition, direction: 'backward' | 'forward', granularity: TextGranularity) {
      const offset = offsetOf(at);
      if (offset === undefined) return undefined;
      const segments = [...(granularity === 'word' ? words : graphemes).segment(text)].filter(
        (s) => !s.segment.includes('\ufffc') && (granularity === 'grapheme' || /\S/u.test(s.segment)),
      );
      const forward = direction === 'forward';
      const segment = forward
        ? segments.find((s) => s.index + s.segment.length > offset)
        : segments.findLast((s) => s.index < offset);
      const target = segment
        ? forward
          ? segment.index + segment.segment.length
          : segment.index
        : forward
          ? text.length
          : 0;
      return target === offset ? undefined : positionAt(target, forward);
    },
    wordAt(at: DocPosition, affinity: 'before' | 'after' = 'after'): DocRange | undefined {
      const offset = offsetOf(at);
      if (offset === undefined || !text.length) return undefined;
      // 命中测试返回最近字缝；点在该字缝左侧时，要选前一个字所属的词。
      const index = Math.max(0, Math.min(offset - (affinity === 'before' ? 1 : 0), text.length - 1));
      const segment = words.segment(text).containing(index);
      if (!segment || segment.segment.includes('\ufffc')) return undefined;
      const start = positionAt(segment.index, false);
      const end = positionAt(segment.index + segment.segment.length, true);
      return start && end ? { start, end } : undefined;
    },
  };
}
