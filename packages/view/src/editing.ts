/** 输入状态只存模型位置；组合文字暂存，compositionend 才提交一个撤销单元。 */
import type {
  DocPosition,
  DocRange,
  Justification,
  ListKind,
  NodeId,
  ParaPropsPatch,
  ResolvedParaProps,
  ResolvedRunProps,
  RunProps,
  RunPropsPatch,
  TextChangeSet,
  TextEditor,
} from '@uw/model';
import {
  buildRunOrder,
  compareDocPositions,
  mapTextRange,
  paragraphText,
  rangeOfNode,
  walkParagraphs,
} from '@uw/model';
import type { PasteParagraph } from './clipboard.ts';
import type { TextGranularity } from './text-navigation.ts';
import { paragraphNavigation } from './text-navigation.ts';

/** 快捷键切换的三种格式；其余格式走 `doc.tx` 的 `setRunProps`。 */
export type ToggleFormat = 'bold' | 'italic' | 'underline';
/** 选区当前的最终格式（级联后），切换按「全部都是才取消」判断。 */
export type FormatQuery = (range: DocRange) => readonly Pick<ResolvedRunProps, ToggleFormat>[];
/** 选区触及段落的级联对齐、编号与缩进。 */
export type ParagraphQuery = (
  range: DocRange,
) => readonly { id: NodeId; props: Pick<ResolvedParaProps, 'justification' | 'numbering' | 'indent'> }[];

/** Word 的列表只有 0–8 九级（`w:ilvl`）。 */
const MAX_LIST_LEVEL = 8;
/**
 * 左缩进写的是字符单位（`w:leftChars`，中文版 Word 的默认写法）时每次挪两个字：
 * 字符单位的缩进换算成 twips 要知道这一段的字号，控制器手上没有。两个字是照中文版
 * 「默认制表位 2 字符」取的，**没有与 Word 的增加缩进量逐格对过**。
 */
const INDENT_CHARS_STEP = 200;

export interface EditingOptions {
  /** Ctrl+M 的步长（`w:defaultTabStop`，twips）；缺省 420，中文版 Word 的默认值。 */
  tabStop?: number;
}

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
  /** 选区全是列表段落，且（折叠时）光标在段首：Tab 交给列表升降级而不是移走焦点。 */
  listIndentable(): boolean;
  /** 每段各自升 / 降一级，夹在 0–8；整次一个撤销单元。 */
  indentList(direction: 'in' | 'out'): void;
  /**
   * Word 的「项目符号 / 编号」按钮：选区触及的段落全是这一种列表时取消编号，否则套用 ——
   * 紧邻的上一段是同类列表就接着它数，选区里已有同类列表就并进去，都没有才新建定义。
   * 已是列表的段落保留层级；整次一个撤销单元（含新建的定义）。
   */
  toggleList(kind: ListKind): void;
  /**
   * Word 的 Ctrl+M / Ctrl+Shift+M：左缩进推到下一个 / 退到上一个默认制表位的整数倍（不低于 0）。
   * 选区全是列表段落时改为逐段升降级 —— Word 的「增加缩进量」在列表上就是降一级。
   */
  indent(direction: 'in' | 'out'): void;
  /** Word 的 Ctrl+1 / 2 / 5：单倍 / 双倍 / 1.5 倍行距，改为多倍行距规则（固定值行距一并改掉）。 */
  lineSpacing(multiple: 1 | 1.5 | 2): void;
  insert(text: string, input?: boolean): void;
  /**
   * 粘贴带格式的段落（见 clipboard.ts）：首段并进光标所在段、其后每段拆出新段，一个撤销单元。
   * 源段落的对齐落在**以源段落标记结尾**的那几段上，最后一块并进原段落的后半截、保留原段落的对齐
   * —— Word 的段落格式跟着段落标记走。
   */
  insertParagraphs(paragraphs: readonly PasteParagraph[]): void;
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
  options: EditingOptions = {},
): EditingController {
  const tabStop = options.tabStop !== undefined && options.tabStop > 0 ? options.tabStop : 420;
  let selection: DocRange | undefined;
  let pending: RunPropsPatch | undefined;
  let composing = false;
  let compositionRange: DocRange | undefined;
  let anchor: DocPosition | undefined;
  let focus: DocPosition | undefined;
  const equal = (a: DocPosition, b: DocPosition) =>
    a.nodeId === b.nodeId && a.contentIndex === b.contentIndex && a.offset === b.offset;
  type ListParagraph = ReturnType<ParagraphQuery>[number];
  /** 只有真的画出了编号（计数器给了 label）才算列表：numId 指向不存在的定义时不显示编号。 */
  function listParagraphs(range: DocRange): readonly ListParagraph[] | undefined {
    const paragraphs = paragraphsOf?.(range) ?? [];
    return paragraphs.length && paragraphs.every((p) => p.props.numbering.label !== undefined)
      ? paragraphs
      : undefined;
  }
  /** 项目符号看 numFmt，其余格式（1. / 一、/ a)）都算编号。 */
  function listKind(p: ListParagraph): ListKind | undefined {
    const label = p.props.numbering.label;
    return label && (label.format === 'bullet' ? 'bullet' : 'decimal');
  }
  function paragraphStart(id: NodeId): DocPosition | undefined {
    const paragraph = [...walkParagraphs(editor.body)].find((p) => p.id === id);
    return paragraph && rangeOfNode(paragraph)?.start;
  }
  /** 折叠光标所在的列表段落，且光标在段首时才返回 —— 列表的 Tab / Backspace 只在这里接管。 */
  function listItemAtStart(): ListParagraph | undefined {
    if (!selection || !equal(selection.start, selection.end)) return undefined;
    const item = listParagraphs(selection)?.[0];
    const start = item && paragraphStart(item.id);
    return start && equal(start, selection.start) ? item : undefined;
  }
  /** 每段各自升 / 降一级，夹在 0–8；整次一个撤销单元。 */
  function shiftListLevels(items: readonly ListParagraph[], direction: 'in' | 'out'): void {
    const starts = items.map((p) => paragraphStart(p.id));
    editor.breakHistory();
    pending = undefined;
    editor.tx((tx) => {
      for (const [i, item] of items.entries()) {
        const current = item.props.numbering.level;
        const level = Math.min(MAX_LIST_LEVEL, Math.max(0, current + (direction === 'in' ? 1 : -1)));
        const at = starts[i];
        // 只写 level：numId 可能来自样式（标题列表），写死会把样式的编号钉在直接格式上。
        if (at && level !== current) tx.setParagraphProps({ start: at, end: at }, { numbering: { level } });
      }
    });
  }
  function patchParagraph(id: NodeId, patch: ParaPropsPatch): void {
    const at = paragraphStart(id);
    if (!at) return;
    editor.breakHistory();
    pending = undefined;
    // 段落格式不拆 run、不动位置，选区原样有效。
    editor.tx((tx) => {
      tx.setParagraphProps({ start: at, end: at }, patch);
    });
  }
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
  function insertParagraphs(paragraphs: readonly PasteParagraph[], origin: 'input' | 'command'): void {
    if (!selection || composing) return;
    let at = selection.start;
    const range = selection;
    const format = pending;
    const first =
      (compareDocPositions(buildRunOrder(editor.body), range.start, range.end) ?? 0) > 0
        ? range.end
        : range.start;
    // 删除选区后剩下的是前一段（合段保留前段格式），光标所在的就是它。
    const host = [...walkParagraphs(editor.body)].find(
      (p) => p.id === first.nodeId || p.runs.some((r) => r.id === first.nodeId),
    );
    // 插入的文字继承左边的 run —— 前一块刚粘进来的那个。前一块改过、这一块没写的格式，
    // 要显式还原成光标处原来的直接格式，否则「粗体 + 普通」粘出来是两段粗体。
    const destination: RunProps =
      host?.runs.find((r) => r.id === first.nodeId)?.props ?? host?.props.markRunProps ?? {};
    // 拆段会把对齐带进新段；源段对齐改过之后，下一段先还原成原段落自己的直接格式。
    const original = host?.props.justification ?? null;
    editor.tx(
      (tx) => {
        if (!equal(range.start, range.end)) at = tx.deleteRange(range);
        let changed = false;
        // 跨段累计：拆出来的空段从拆点那个 run（刚粘进来的）继承格式，同样要还原。
        const touched = new Set<string>();
        for (const [i, paragraph] of paragraphs.entries()) {
          if (i > 0) {
            at = tx.splitParagraph(at);
            if (changed) tx.setParagraphProps({ start: at, end: at }, { justification: original });
            changed = false;
          }
          for (const run of paragraph.runs) {
            // 模型还没有插入制表位的命令，粘贴进来的制表符先当空格，不让整次粘贴失败。
            const text = run.text.replace(/[\t\r\n]/g, ' ');
            if (!text) continue;
            at = tx.insertText(at, text);
            const reset: Record<string, unknown> = {};
            for (const key of touched)
              if (!(key in run.patch)) reset[key] = destination[key as keyof RunProps] ?? null;
            const patch: RunPropsPatch = { ...(reset as RunPropsPatch), ...format, ...run.patch };
            for (const key of Object.keys(patch)) touched.add(key);
            // 空段落首次输入会先建 run，插入起点只能从返回的终点倒推。
            if (Object.keys(patch).length)
              at = tx.setRunProps({ start: { ...at, offset: at.offset - text.length }, end: at }, patch).end;
          }
          if (paragraph.justification && i < paragraphs.length - 1) {
            tx.setParagraphProps({ start: at, end: at }, { justification: paragraph.justification });
            changed = true;
          }
        }
      },
      { origin },
    );
    pending = undefined;
    collapse(at);
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
      const active = current.length > 0 && current.every((p) => p.props.justification === justification);
      editor.breakHistory();
      // 段落格式不拆 run、不动位置，选区与方向原样保留，无需重新 select。
      editor.tx((tx) => {
        tx.setParagraphProps(range, { justification: active ? 'left' : justification });
      });
    },
    listIndentable() {
      if (!selection || composing) return false;
      return equal(selection.start, selection.end) ? !!listItemAtStart() : !!listParagraphs(selection);
    },
    indentList(direction) {
      if (!selection || !controller.listIndentable()) return;
      shiftListLevels(listParagraphs(selection) ?? [], direction);
    },
    indent(direction) {
      if (!selection || composing || !paragraphsOf) return;
      const listItems = listParagraphs(selection);
      if (listItems) {
        shiftListLevels(listItems, direction);
        return;
      }
      const items = paragraphsOf(selection);
      const starts = items.map((p) => paragraphStart(p.id));
      // 推到「下一个」整数倍而不是加一个步长：已经在 1.3 格的段落按一次到 2 格，与 Word 一致。
      const step = (value: number, unit: number) =>
        direction === 'in'
          ? (Math.floor(value / unit) + 1) * unit
          : Math.max(0, (Math.ceil(value / unit) - 1) * unit);
      editor.breakHistory();
      pending = undefined;
      editor.tx((tx) => {
        for (const [i, item] of items.entries()) {
          const at = starts[i];
          const { left, leftChars } = item.props.indent;
          // 字符单位优先（盖过 twips），所以有 leftChars 时只能改它
          const patch =
            leftChars !== 0
              ? { leftChars: step(leftChars, INDENT_CHARS_STEP) }
              : { left: step(left, tabStop) };
          if (at) tx.setParagraphProps({ start: at, end: at }, { indent: patch });
        }
      });
    },
    lineSpacing(multiple) {
      if (!selection || composing) return;
      const range = selection;
      editor.breakHistory();
      pending = undefined;
      // 多倍行距的 line 以 1/240 行计；lineRule 必须一起写，否则固定值行距会把 240 读成 12pt。
      editor.tx((tx) => {
        tx.setParagraphProps(range, { spacing: { line: 240 * multiple, lineRule: 'auto' } });
      });
    },
    toggleList(kind) {
      if (!selection || composing || !paragraphsOf) return;
      const items = paragraphsOf(selection);
      const first = items[0];
      if (!first) return;
      const active = items.every((p) => listKind(p) === kind);
      const all = [...walkParagraphs(editor.body)];
      const previous = all[all.findIndex((p) => p.id === first.id) - 1];
      const at = previous && rangeOfNode(previous)?.start;
      const before = at && paragraphsOf({ start: at, end: at })[0];
      // 接着上一段数是 Word 的习惯：列表中间插了一段正文、再点编号，编号不从 1 重来。
      const reuse = [...(before ? [before] : []), ...items].find((p) => listKind(p) === kind)?.props.numbering
        .numId;
      const starts = items.map((p) => paragraphStart(p.id));
      editor.breakHistory();
      pending = undefined;
      editor.tx((tx) => {
        const numId = active ? 0 : (reuse ?? tx.addList(kind));
        for (const [i, item] of items.entries()) {
          const start = starts[i];
          const level = listKind(item) ? item.props.numbering.level : 0;
          if (start)
            tx.setParagraphProps(
              { start, end: start },
              { numbering: active ? { numId: 0 } : { numId, level } },
            );
        }
      });
    },
    insert(text, input = false) {
      if (text === '') return;
      const lines = text.replace(/\r\n?/g, '\n').split('\n');
      insertParagraphs(
        lines.map((line) => ({ runs: line ? [{ text: line, patch: {} }] : [] })),
        input ? 'input' : 'command',
      );
    },
    insertParagraphs(paragraphs) {
      if (paragraphs.length) insertParagraphs(paragraphs, 'command');
    },
    enter() {
      if (!selection || composing) return;
      const item = listItemAtStart();
      const paragraph = item && [...walkParagraphs(editor.body)].find((p) => p.id === item.id);
      // Word：空列表项上的 Enter 不再造一个空项，而是先降级、到顶层就结束列表。
      if (item && paragraph && paragraphText(paragraph) === '') {
        const level = item.props.numbering.level;
        patchParagraph(item.id, { numbering: level > 0 ? { level: level - 1 } : { numId: 0 } });
        return;
      }
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
      const item = direction === 'backward' ? listItemAtStart() : undefined;
      // Word：列表项段首的退格先去掉编号，再按一次才合段。
      if (item) {
        // 缩进原本由编号层级给，去编号后随之失效、文字跳回页边。Word 把文字所在的位置
        // （级联后的左缩进）写成直接格式留住它；首行 / 悬挂写显式 0 而不是删掉 ——
        // 删掉会退回样式的值，公文正文样式的首行缩进两字会把首行顶出去。
        const { left, leftChars } = item.props.indent;
        patchParagraph(item.id, {
          numbering: { numId: 0 },
          indent: { left, leftChars, firstLine: 0, firstLineChars: 0, hanging: 0, hangingChars: 0 },
        });
        return;
      }
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
