/** 输入状态只存模型位置；组合文字暂存，compositionend 才提交一个撤销单元。 */
import type {
  DocPosition,
  DocRange,
  Justification,
  ResolvedParaProps,
  ResolvedRunProps,
  RunPropsPatch,
  TextChangeSet,
  TextEditor,
} from '@uw/model';
import { buildRunOrder, compareDocPositions, mapTextRange, rangeOfNode, walkParagraphs } from '@uw/model';
import type { TextGranularity } from './text-navigation.ts';
import { paragraphNavigation } from './text-navigation.ts';

/** 快捷键切换的三种格式；其余格式走 `doc.tx` 的 `setRunProps`。 */
export type ToggleFormat = 'bold' | 'italic' | 'underline';
/** 选区当前的最终格式（级联后），切换按「全部都是才取消」判断。 */
export type FormatQuery = (range: DocRange) => readonly Pick<ResolvedRunProps, ToggleFormat>[];
/** 选区触及段落的级联对齐方式。 */
export type ParagraphQuery = (range: DocRange) => readonly Pick<ResolvedParaProps, 'justification'>[];

export interface EditingController {
  readonly selection: DocRange | undefined;
  readonly anchor: DocPosition | undefined;
  readonly focus: DocPosition | undefined;
  readonly composing: boolean;
  /** 折叠光标上切换、尚未落到文字上的格式；下一次输入使用，移动光标即丢弃。 */
  readonly pendingFormat: RunPropsPatch | undefined;
  select(range: DocRange): void;
  toggleFormat(format: ToggleFormat): void;
  /** Word 的 Ctrl+L/E/R/J：全部已是该对齐时回到左对齐，否则设置；选区不变。 */
  align(justification: Justification): void;
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

const underlined = (value: string) => value !== 'none' && value !== '';

export function createEditingController(
  editor: TextEditor,
  formatOf?: FormatQuery,
  paragraphsOf?: ParagraphQuery,
): EditingController {
  let selection: DocRange | undefined;
  let pending: RunPropsPatch | undefined;
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
    get pendingFormat() {
      return pending && { ...pending };
    },
    select(range) {
      if (composing) return;
      editor.breakHistory();
      pending = undefined;
      select(range);
    },
    toggleFormat(format) {
      if (!selection || composing) return;
      const range = selection;
      const collapsed = equal(range.start, range.end);
      const on = (p: Pick<ResolvedRunProps, ToggleFormat>) =>
        format === 'underline' ? underlined(p.underline) : p[format];
      const current = formatOf?.(range) ?? [];
      const queued = collapsed ? pending?.[format] : undefined;
      const active =
        queued !== undefined && queued !== null
          ? typeof queued === 'string'
            ? underlined(queued)
            : queued
          : current.length > 0 && current.every(on);
      const patch: RunPropsPatch =
        format === 'underline' ? { underline: active ? 'none' : 'single' } : { [format]: !active };
      const backward = focus !== undefined && !collapsed && equal(focus, range.start);
      let next = range;
      editor.breakHistory();
      // 空段落的折叠光标由模型改段落标记；非空段落内折叠时模型无修改，格式暂存到下一次输入。
      const change = editor.tx((tx) => {
        next = tx.setRunProps(range, patch);
      });
      if (collapsed) {
        if (!change) pending = { ...pending, ...patch };
        return;
      }
      select(backward ? { start: next.end, end: next.start } : next);
    },
    align(justification) {
      if (!selection || composing) return;
      const range = selection;
      const current = paragraphsOf?.(range) ?? [];
      const active = current.length > 0 && current.every((p) => p.justification === justification);
      editor.breakHistory();
      // 段落格式不拆 run、不动位置，选区与方向原样保留，无需重新 select。
      editor.tx((tx) => {
        tx.setParagraphProps(range, { justification: active ? 'left' : justification });
      });
    },
    insert(text, input = false) {
      if (!selection || composing || text === '') return;
      let at = selection.start;
      const range = selection;
      const format = pending;
      editor.tx(
        (tx) => {
          if (!equal(range.start, range.end)) at = tx.deleteRange(range);
          const lines = text.replace(/\r\n?/g, '\n').split('\n');
          for (const [i, line] of lines.entries()) {
            if (i > 0) at = tx.splitParagraph(at);
            if (!line) continue;
            at = tx.insertText(at, line);
            // 空段落首次输入会先建 run，插入起点只能从返回的终点倒推。
            if (format)
              at = tx.setRunProps({ start: { ...at, offset: at.offset - line.length }, end: at }, format).end;
          }
        },
        { origin: input ? 'input' : 'command' },
      );
      pending = undefined;
      collapse(at);
    },
    enter() {
      if (!selection || composing) return;
      pending = undefined;
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
      pending = undefined;
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
      pending = undefined;
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
