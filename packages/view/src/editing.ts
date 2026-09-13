/** 输入状态只存模型位置；组合文字暂存，compositionend 才提交一个撤销单元。 */
import type { DocPosition, DocRange, TextChangeSet, TextEditor } from '@uw/model';
import { buildRunOrder, compareDocPositions, mapTextRange, rangeOfNode, walkParagraphs } from '@uw/model';

export interface EditingController {
  readonly selection: DocRange | undefined;
  readonly composing: boolean;
  select(range: DocRange): void;
  insert(text: string, input?: boolean): void;
  enter(): void;
  delete(direction: 'backward' | 'forward'): void;
  move(direction: 'backward' | 'forward', extend?: boolean): void;
  compositionStart(): void;
  compositionEnd(text: string): void;
  compositionCancel(): void;
  apply(change: TextChangeSet): void;
}

export function createEditingController(editor: TextEditor): EditingController {
  let selection: DocRange | undefined;
  let composing = false;
  let compositionRange: DocRange | undefined;
  let anchor: DocPosition | undefined;
  let focus: DocPosition | undefined;
  const segmenter = new Intl.Segmenter('zh', { granularity: 'grapheme' });
  const equal = (a: DocPosition, b: DocPosition) =>
    a.nodeId === b.nodeId && a.contentIndex === b.contentIndex && a.offset === b.offset;
  function select(range: DocRange): void {
    const order = compareDocPositions(buildRunOrder(editor.body), range.start, range.end);
    if (order === undefined) throw new RangeError('选区不在正文中');
    anchor = { ...range.start };
    focus = { ...range.end };
    selection = order > 0 ? { start: { ...focus }, end: { ...anchor } } : structuredClone(range);
  }
  function collapse(at: DocPosition): void {
    select({ start: at, end: at });
  }
  function neighbor(at: DocPosition, direction: 'backward' | 'forward'): DocPosition | undefined {
    const paragraphs = [...walkParagraphs(editor.body)];
    const pi = paragraphs.findIndex((p) => p.id === at.nodeId || p.runs.some((r) => r.id === at.nodeId));
    const p = paragraphs[pi];
    if (!p) return undefined;
    const points: DocPosition[] = [];
    if (!p.runs.length) points.push({ nodeId: p.id, contentIndex: 0, offset: 0 });
    for (const run of p.runs) {
      if (!run.content.length) points.push({ nodeId: run.id, contentIndex: 0, offset: 0 });
      for (const [ci, c] of run.content.entries()) {
        if (c.kind !== 'text') continue;
        for (const part of segmenter.segment(c.text))
          points.push({ nodeId: run.id, contentIndex: ci, offset: part.index });
        points.push({ nodeId: run.id, contentIndex: ci, offset: c.text.length });
      }
    }
    const index = points.findIndex((p) => equal(p, at));
    const step = direction === 'backward' ? -1 : 1;
    let next = index + step;
    // run / 片段边界的两个位置是同一条字缝，不让方向键在那里多停一次。
    while (index >= 0 && next >= 0 && next < points.length) {
      const candidate = points[next] as DocPosition;
      const prev = points[next - step] as DocPosition;
      if (candidate.nodeId === prev.nodeId && candidate.contentIndex === prev.contentIndex) return candidate;
      next += step;
    }
    const adjacent = paragraphs[pi + step];
    const range = adjacent && rangeOfNode(adjacent);
    return direction === 'backward' ? range?.end : range?.start;
  }
  const controller: EditingController = {
    get selection() {
      return selection && structuredClone(selection);
    },
    get composing() {
      return composing;
    },
    select(range) {
      if (composing) return;
      editor.breakHistory();
      select(range);
    },
    insert(text, input = false) {
      if (!selection || composing || text === '') return;
      let at = selection.start;
      const range = selection;
      editor.tx(
        (tx) => {
          if (!equal(range.start, range.end)) at = tx.deleteRange(range);
          const lines = text.replace(/\r\n?/g, '\n').split('\n');
          for (const [i, line] of lines.entries()) {
            if (i > 0) at = tx.splitParagraph(at);
            if (line) at = tx.insertText(at, line);
          }
        },
        { origin: input ? 'input' : 'command' },
      );
      collapse(at);
    },
    enter() {
      if (!selection || composing) return;
      let at = selection.start;
      const range = selection;
      editor.tx((tx) => {
        if (!equal(range.start, range.end)) at = tx.deleteRange(range);
        at = tx.splitParagraph(at);
      });
      collapse(at);
    },
    delete(direction) {
      if (!selection || composing) return;
      let range = selection;
      if (equal(range.start, range.end)) {
        const next = neighbor(range.start, direction);
        if (!next) return;
        range =
          direction === 'backward' ? { start: next, end: range.end } : { start: range.start, end: next };
      }
      let at = range.start;
      editor.tx((tx) => {
        at = tx.deleteRange(range);
      });
      collapse(at);
    },
    move(direction, extend = false) {
      if (!selection || composing) return;
      editor.breakHistory();
      if (!extend && !equal(selection.start, selection.end)) {
        collapse(direction === 'backward' ? selection.start : selection.end);
        return;
      }
      const next = neighbor(focus ?? selection.end, direction);
      if (next) select({ start: extend ? (anchor ?? selection.start) : next, end: next });
    },
    compositionStart() {
      if (composing) return;
      editor.breakHistory();
      composing = true;
      compositionRange = selection && structuredClone(selection);
    },
    compositionEnd(text) {
      if (!composing) return;
      composing = false;
      selection = compositionRange;
      compositionRange = undefined;
      controller.insert(text);
      editor.breakHistory();
    },
    compositionCancel() {
      composing = false;
      compositionRange = undefined;
      editor.breakHistory();
    },
    apply(change) {
      if (selection) select(mapTextRange(selection, change));
      if (compositionRange) compositionRange = mapTextRange(compositionRange, change);
    },
  };
  const first = [...walkParagraphs(editor.body)][0];
  if (first) {
    const range = rangeOfNode(first);
    if (range) collapse(range.start);
  }
  return controller;
}
