/** 输入状态只存模型位置；组合文字暂存，compositionend 才提交一个撤销单元。 */
import type {
  BuiltinStyleName,
  DocPosition,
  DocRange,
  InlineKind,
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
  adjacentCellRange,
  buildRunOrder,
  builtinStyleDefinition,
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
/**
 * 选区触及段落的级联对齐、编号、缩进与段距。`charUnit` 是字符单位缩进里一个字的宽度（twips，
 * 见 `@uw/layout` 的 `indentCharUnit`）；没有它就换算不了 `w:leftChars`，Ctrl+T 遇到字符单位时不动。
 */
export type ParagraphQuery = (range: DocRange) => readonly {
  id: NodeId;
  props: Pick<ResolvedParaProps, 'justification' | 'numbering' | 'indent' | 'spacing'>;
  charUnit?: number;
}[];

/** Word 的列表只有 0–8 九级（`w:ilvl`）。 */
const MAX_LIST_LEVEL = 8;
/**
 * 左缩进写的是字符单位（`w:leftChars`，中文版 Word 的默认写法）时每次挪两个字：
 * 字符单位的缩进换算成 twips 要知道这一段的字号，控制器手上没有。两个字是照中文版
 * 「默认制表位 2 字符」取的，**没有与 Word 的增加缩进量逐格对过**。
 */
const INDENT_CHARS_STEP = 200;
/** Ctrl+0 加的段前间距：Word 的「一行」按 12pt 算，与段落字号无关。 */
const SPACE_BEFORE = 240;

/** 文档里已有的一份段落样式（`StyleSheet` 的纯数据摘要）。 */
export interface ParagraphStyleInfo {
  id: string;
  /** `w:name`：内建样式是英文小写名（`heading 1`），套用内建样式按它找，不按 id —— 中文版 Word 的 id 是 `1`。 */
  name: string;
  /** `w:next`：段尾回车后新段落用哪个样式。 */
  next?: string | undefined;
  isDefault: boolean;
}

export interface EditingOptions {
  /** Ctrl+M 的步长（`w:defaultTabStop`，twips）；缺省 420，中文版 Word 的默认值。 */
  tabStop?: number;
  /** 文档的段落样式；缺省时套用样式只能套内建样式，回车也不按 `w:next` 换样式。 */
  styles?: readonly ParagraphStyleInfo[];
}

/**
 * 套用段落样式时，字符直接格式的某一项覆盖了段落**过半**的字才清掉（Word 的行为：
 * 整段手工设成四号的正文套「标题 1」变成标题的字号，只有几个字加粗的仍保留加粗）。
 * 逐项判断、按字数算。**没有与 Word 逐格对过**：「正好一半」归哪边、段落标记算不算一个字都是猜的。
 */
const STYLE_CLEAR_RATIO = 0.5;
/** 这几项不是「格式」：语言标记与字符样式跟着内容走，套段落样式不该动它们。 */
const STYLE_KEEP_RUN_KEYS: ReadonlySet<string> = new Set(['styleId', 'langEastAsia']);

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
   * Word 表格里的 Tab / Shift+Tab：选中下一个 / 上一个单元格的全部内容（按 focus 所在的最内层表格走）。
   * 返回 false 表示不在单元格里，调用方照常处理 Tab。末格 Tab 学 Word 在下方加一行（照末行抄结构），
   * 光标进新行首格，一个撤销单元；首格 Shift+Tab 返回 true 但不动 —— 不该退回去插制表位。
   */
  moveCell(direction: 'backward' | 'forward'): boolean;
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
  /**
   * Word 的 Ctrl+T / Ctrl+Shift+T：首行原地不动，其余行（左缩进）推到下一个 / 退到上一个默认制表位的整数倍，
   * 差出来的就是悬挂缩进。退只在已有悬挂时退，且不退过首行 —— 不会反过来变成首行缩进。
   * 结果一律写 twips 并把字符单位清零（字符单位优先，留着会盖掉新值）。
   */
  hangingIndent(direction: 'in' | 'out'): void;
  /** Word 的 Ctrl+0：段前间距在 0 与 12pt 之间切换，选区全都已有段前间距才取消。 */
  toggleSpaceBefore(): void;
  /** Word 的 Ctrl+1 / 2 / 5：单倍 / 双倍 / 1.5 倍行距，改为多倍行距规则（固定值行距一并改掉）。 */
  lineSpacing(multiple: 1 | 1.5 | 2): void;
  insert(text: string, input?: boolean): void;
  /**
   * 粘贴带格式的段落（见 clipboard.ts）：首段并进光标所在段、其后每段拆出新段，一个撤销单元。
   * 源段落的对齐落在**以源段落标记结尾**的那几段上，最后一块并进原段落的后半截、保留原段落的对齐
   * —— Word 的段落格式跟着段落标记走。
   */
  insertParagraphs(paragraphs: readonly PasteParagraph[]): void;
  /** Word 的 Tab（非列表位置）与 Shift+Enter：替换选区，插入制表位 / 软换行，用上暂存格式。 */
  insertInline(kind: InlineKind): void;
  /**
   * Word 的 Ctrl+Enter：替换选区，插入分页符并紧接着拆段 —— Word 2013 起分页符后面跟一个段落标记，
   * 后文从新的一段、新的一页开始，而不是同一段接着排。单元格里不插（事务拒绝，模型不变）。
   */
  pageBreak(): void;
  /**
   * Word 的样式库：给选区触及的每一段套用段落样式（整次一个撤销单元）。段落直接格式全部清掉、
   * 回到样式（编号也一样 —— 列表项套「标题 1」就不再是列表项）；字符直接格式按「过半」规则清
   * （见 `STYLE_CLEAR_RATIO`）。id 不在文档样式里时不动。
   */
  applyStyle(styleId: string): void;
  /**
   * 按名字套内建样式（Word 的 Ctrl+Alt+1 / 2 / 3 与 Ctrl+Shift+N）：文档里有这份样式就用它，
   * 没有就照中文版 Word 的默认模板补一份定义（`builtinStyleDefinition`），随本次事务提交与撤销。
   * `'normal'` 是默认段落样式。
   */
  applyBuiltinStyle(name: BuiltinStyleName | 'normal'): void;
  /**
   * 回车拆段。光标（或选区末端）在段尾、且本段样式的 `w:next` 是别的样式时，新段落换成那个样式、
   * 不带本段的直接格式 —— 标题后面回车接着写的是正文，而不是又一个居中加粗的标题。
   */
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

/** 粘贴文字里代表非文字片段的控制字符。 */
const INLINE_PIECES: ReadonlyMap<string, InlineKind> = new Map([
  ['\t', 'tab'],
  ['\n', 'lineBreak'],
  ['\f', 'pageBreak'],
]);

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
  const styles = options.styles ?? [];
  const defaultStyle = styles.find((st) => st.isDefault)?.id ?? '';
  /** 已有样式 + 本树新增的定义（后者随撤销进出，每次现查） */
  function styleInfo(id: string): ParagraphStyleInfo | undefined {
    const added = editor.body.styles?.find((st) => st.id === id);
    return added
      ? { id, name: added.name, next: added.next, isDefault: false }
      : styles.find((st) => st.id === id);
  }
  function paragraphAt(at: DocPosition) {
    return [...walkParagraphs(editor.body)].find(
      (p) => p.id === at.nodeId || p.runs.some((r) => r.id === at.nodeId),
    );
  }
  /** 把一段的段落直接格式换成「只有样式」：每一项置 null，样式是默认样式时连 pStyle 也不写。 */
  function styleOnlyPatch(props: object, styleId: string): ParaPropsPatch {
    const patch: Record<string, unknown> = {};
    for (const key of Object.keys(props)) if (key !== 'markRunProps') patch[key] = null;
    patch.styleId = styleId === defaultStyle ? null : styleId;
    return patch as ParaPropsPatch;
  }
  function nullPatch(keys: Iterable<string>): RunPropsPatch {
    const patch: Record<string, null> = {};
    for (const key of keys) if (!STYLE_KEEP_RUN_KEYS.has(key)) patch[key] = null;
    return patch as RunPropsPatch;
  }
  /** 本树里的段落（含单元格）按选区触及的顺序；折叠选区即光标所在段。 */
  function paragraphsInSelection(range: DocRange) {
    const order = buildRunOrder(editor.body);
    const all = [...walkParagraphs(editor.body)];
    const first = paragraphAt(range.start);
    const last = paragraphAt(range.end);
    const from = all.findIndex((p) => p.id === first?.id);
    const to = all.findIndex((p) => p.id === last?.id);
    if (from < 0 || to < 0) return [];
    // 选区可能跨表格：嵌套结构的深度优先序就是文档序，首尾之间的都算
    return compareDocPositions(order, range.start, range.end) === undefined ? [] : all.slice(from, to + 1);
  }
  function applyStyleId(styleId: string, define?: BuiltinStyleName): void {
    if (!selection || composing) return;
    const paragraphs = paragraphsInSelection(selection);
    if (!paragraphs.length) return;
    editor.breakHistory();
    pending = undefined;
    editor.tx((tx) => {
      const id = define
        ? tx.addStyle(builtinStyleDefinition(define, (x) => styleInfo(x) !== undefined, defaultStyle))
        : styleId;
      for (const p of paragraphs) {
        const range = rangeOfNode(p);
        if (!range) continue;
        tx.setParagraphProps({ start: range.start, end: range.start }, styleOnlyPatch(p.props, id));
        // 字符直接格式按「过半」清：逐项数有多少字带着它
        let total = 0;
        const counts = new Map<string, number>();
        for (const run of p.runs) {
          const length = run.content.reduce((n, c) => n + (c.kind === 'text' ? c.text.length : 0), 0);
          total += length;
          for (const key of Object.keys(run.props)) counts.set(key, (counts.get(key) ?? 0) + length);
        }
        const clear =
          total === 0
            ? Object.keys(p.props.markRunProps ?? {})
            : [...counts].filter(([, n]) => n > total * STYLE_CLEAR_RATIO).map(([key]) => key);
        const patch = nullPatch(clear);
        // 整段范围的端点落在段首段尾，不拆 run，选区原样有效；空段落改的是段落标记
        if (Object.keys(patch).length) tx.setRunProps(range, patch);
      }
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
    // 拆段不出容器：光标在正文顶层，粘进来的每一段就都在顶层。
    const topLevel = editor.body.sections.some((s) => s.blocks.some((b) => b.id === host?.id));
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
            if (!run.text) continue;
            const reset: Record<string, unknown> = {};
            for (const key of touched)
              if (!(key in run.patch)) reset[key] = destination[key as keyof RunProps] ?? null;
            const patch: RunPropsPatch = { ...(reset as RunPropsPatch), ...format, ...run.patch };
            for (const key of Object.keys(patch)) touched.add(key);
            // 段内的 \t / \n / \f 是制表位、软换行与分页符（Tab、Shift+Enter、HTML 的 <br>），不是分段。
            for (const piece of run.text.replace(/\r/g, '').split(/([\t\n\f])/)) {
              if (!piece) continue;
              const inline = INLINE_PIECES.get(piece);
              // 单元格里事务不收分页符（Word 会拆表），退成软换行，别让整次粘贴回滚。
              at = inline
                ? tx.insertInline(at, inline === 'pageBreak' && !topLevel ? 'lineBreak' : inline)
                : tx.insertText(at, piece);
              // 空段落首次输入会先建 run，插入起点只能从返回的终点倒推；这些片段的长度是 1。
              const length = inline ? 1 : piece.length;
              if (Object.keys(patch).length)
                at = tx.setRunProps({ start: { ...at, offset: at.offset - length }, end: at }, patch).end;
            }
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
    moveCell(direction) {
      if (!selection || composing) return false;
      const target = adjacentCellRange(editor.body, (focus ?? selection.end).nodeId, direction);
      if (target === undefined) return false;
      if (target) {
        controller.select(target);
      } else if (direction === 'forward') {
        const from = focus ?? selection.end;
        let at = from;
        editor.breakHistory();
        pending = undefined;
        editor.tx((tx) => {
          at = tx.insertRow(from, 'below');
        });
        collapse(at);
      }
      return true;
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
    hangingIndent(direction) {
      if (!selection || composing || !paragraphsOf) return;
      const items = paragraphsOf(selection);
      const starts = items.map((p) => paragraphStart(p.id));
      const patches = items.map((item) => {
        const ind = item.props.indent;
        const chars = ind.leftChars !== 0 || ind.firstLineChars !== 0 || ind.hangingChars !== 0;
        if (chars && !(item.charUnit && item.charUnit > 0)) return undefined;
        const unit = item.charUnit ?? 0;
        // 与布局的 indentGeometry 同一套取舍：字符单位优先、hanging 压过 firstLine
        const left = ind.leftChars !== 0 ? (ind.leftChars / 100) * unit : ind.left;
        const hanging = ind.hangingChars !== 0 ? (ind.hangingChars / 100) * unit : ind.hanging;
        const firstLine = ind.firstLineChars !== 0 ? (ind.firstLineChars / 100) * unit : ind.firstLine;
        const first = hanging !== 0 ? left - hanging : left + firstLine;
        if (direction === 'out' && first >= left) return undefined;
        const next =
          direction === 'in'
            ? (Math.floor(left / tabStop) + 1) * tabStop
            : Math.max(first, 0, (Math.ceil(left / tabStop) - 1) * tabStop);
        const diff = Math.round(next - first);
        return {
          left: Math.round(first) + diff,
          leftChars: 0,
          hanging: Math.max(0, diff),
          hangingChars: 0,
          firstLine: Math.max(0, -diff),
          firstLineChars: 0,
        };
      });
      if (!patches.some(Boolean)) return;
      editor.breakHistory();
      pending = undefined;
      editor.tx((tx) => {
        for (const [i, patch] of patches.entries()) {
          const at = starts[i];
          if (at && patch) tx.setParagraphProps({ start: at, end: at }, { indent: patch });
        }
      });
    },
    toggleSpaceBefore() {
      if (!selection || composing || !paragraphsOf) return;
      const range = selection;
      const items = paragraphsOf(range);
      if (!items.length) return;
      const spaced = items.every(
        (p) =>
          p.props.spacing.before > 0 || p.props.spacing.beforeLines > 0 || p.props.spacing.beforeAutospacing,
      );
      editor.breakHistory();
      pending = undefined;
      // 行单位与自动间距都压过 before，必须一起清掉，否则写进去的 12pt 不生效。
      editor.tx((tx) => {
        tx.setParagraphProps(range, {
          spacing: { before: spaced ? 0 : SPACE_BEFORE, beforeLines: 0, beforeAutospacing: false },
        });
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
    insertInline(kind) {
      if (kind === 'pageBreak') {
        controller.pageBreak();
        return;
      }
      insertParagraphs([{ runs: [{ text: kind === 'tab' ? '\t' : '\n', patch: {} }] }], 'command');
    },
    pageBreak() {
      if (!selection || composing) return;
      const range = selection;
      let at = range.start;
      editor.breakHistory();
      pending = undefined;
      editor.tx((tx) => {
        if (!equal(range.start, range.end)) at = tx.deleteRange(range);
        at = tx.splitParagraph(tx.insertInline(at, 'pageBreak'));
      });
      collapse(at);
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
      // 段尾回车按 w:next 换样式：选区末端在段尾，删掉选区后光标同样落在（前一段的）段尾
      const host = paragraphAt(range.start);
      const tail = paragraphAt(range.end);
      const tailEnd = tail && rangeOfNode(tail)?.end;
      const current = host && (host.props.styleId ?? defaultStyle);
      const next =
        current !== undefined && tailEnd && equal(range.end, tailEnd) ? styleInfo(current)?.next : undefined;
      const restyle = next !== undefined && next !== current && styleInfo(next) ? next : undefined;
      editor.tx((tx) => {
        if (!equal(range.start, range.end)) at = tx.deleteRange(range);
        // 换样式时新段落一点直接格式都不带：标题的居中、加粗、字号不该带进正文
        at = tx.splitParagraph(at, restyle === undefined ? {} : { inherit: false });
        if (restyle !== undefined && restyle !== defaultStyle)
          tx.setParagraphProps({ start: at, end: at }, { styleId: restyle });
      });
      collapse(at);
    },
    applyStyle(styleId) {
      if (styleInfo(styleId)) applyStyleId(styleId);
    },
    applyBuiltinStyle(name) {
      if (name === 'normal') {
        if (defaultStyle) applyStyleId(defaultStyle);
        return;
      }
      const existing = [...styles, ...(editor.body.styles ?? [])].find(
        (st) => st.name.toLowerCase() === name,
      );
      if (existing) applyStyleId(existing.id);
      else applyStyleId('', name);
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
        try {
          select(backward ? { start: range.end, end: range.start } : range);
        } catch {
          // 映射不到的选区（宿主的事务删掉了光标所在的整块、却没给位置迁移）退到正文开头，
          // 而不是让刷新视图的整条链抛错 —— 模型已经提交了，视图必须跟上
          const first = [...walkParagraphs(editor.body)][0];
          const start = first && rangeOfNode(first)?.start;
          if (start) collapse(start);
          else selection = anchor = focus = undefined;
        }
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
