/** 输入状态只存模型位置；组合文字暂存，compositionend 才提交一个撤销单元。 */
import type { DocPosition, DocRange, TextChangeSet, TextEditor } from '@uw/model';
import { buildRunOrder, compareDocPositions, mapTextRange, rangeOfNode, walkParagraphs } from '@uw/model';
import type { TextGranularity } from './text-navigation.ts';
import { paragraphNavigation } from './text-navigation.ts';

export interface EditingController {
  readonly selection: DocRange | undefined;
  readonly anchor: DocPosition | undefined;
  readonly focus: DocPosition | undefined;
  readonly composing: boolean;
  select(range: DocRange): void;
  insert(text: string, input?: boolean): void;
  enter(): void;
  delete(direction: 'backward' | 'forward', granularity?: TextGranularity): void;
  move(direction: 'backward' | 'forward', extend?: boolean, granularity?: TextGranularity): void;
  moveToDocumentBoundary(boundary: 'start' | 'end', extend?: boolean): void;
  selectAll(): void;
  selectWord(at: DocPosition, affinity?: 'before' | 'after'): void;
  cut(write: () => void): void;
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
  function documentRange(): DocRange | undefined {
    let start: DocPosition | undefined;
    let end: DocPosition | undefined;
    // 从模型取正文边界，分页、重复表头和虚拟化不能改变全文选区；保留首末空段。
    for (const paragraph of walkParagraphs(editor.body)) {
      const range = rangeOfNode(paragraph);
      if (!range) continue;
      start ??= range.start;
      end = range.end;
    }
    return start && end ? { start, end } : undefined;
  }
  function neighbor(
    at: DocPosition,
    direction: 'backward' | 'forward',
    granularity: TextGranularity = 'grapheme',
  ): DocPosition | undefined {
    const paragraphs = [...walkParagraphs(editor.body)];
    const pi = paragraphs.findIndex((p) => p.id === at.nodeId || p.runs.some((r) => r.id === at.nodeId));
    const p = paragraphs[pi];
    if (!p) return undefined;
    const next = paragraphNavigation(p).move(at, direction, granularity);
    if (next) return next;
    const step = direction === 'backward' ? -1 : 1;
    const adjacent = paragraphs[pi + step];
    const range = adjacent && rangeOfNode(adjacent);
    return direction === 'backward' ? range?.end : range?.start;
  }
  const controller: EditingController = {
    get selection() {
      return selection && structuredClone(selection);
    },
    get anchor() {
      return anchor && { ...anchor };
    },
    get focus() {
      return focus && { ...focus };
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
    cut(write) {
      if (!selection || composing || equal(selection.start, selection.end)) return;
      const range = selection;
      let at = range.start;
      editor.tx((tx) => {
        at = tx.deleteRange(range);
        // 先验证删除范围；剪贴板写入抛错时整笔事务回滚，不能丢失原文。
        write();
      });
      collapse(at);
    },
    delete(direction, granularity = 'grapheme') {
      if (!selection || composing) return;
      let range = selection;
      if (equal(range.start, range.end)) {
        const next = neighbor(range.start, direction, granularity);
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
    move(direction, extend = false, granularity = 'grapheme') {
      if (!selection || composing) return;
      editor.breakHistory();
      if (!extend && !equal(selection.start, selection.end)) {
        collapse(direction === 'backward' ? selection.start : selection.end);
        return;
      }
      const next = neighbor(focus ?? selection.end, direction, granularity);
      if (next) select({ start: extend ? (anchor ?? selection.start) : next, end: next });
    },
    moveToDocumentBoundary(boundary, extend = false) {
      if (composing) return;
      const at = documentRange()?.[boundary];
      if (at) controller.select({ start: extend ? (anchor ?? at) : at, end: at });
    },
    selectAll() {
      if (composing) return;
      const range = documentRange();
      if (range) controller.select(range);
    },
    selectWord(at, affinity) {
      if (composing) return;
      const paragraph = [...walkParagraphs(editor.body)].find(
        (p) => p.id === at.nodeId || p.runs.some((r) => r.id === at.nodeId),
      );
      const range = paragraph && paragraphNavigation(paragraph).wordAt(at, affinity);
      if (range) controller.select(range);
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
      if (selection) {
        const backward = focus && equal(focus, selection.start);
        const range = mapTextRange(selection, change);
        select(backward ? { start: range.end, end: range.start } : range);
      }
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
