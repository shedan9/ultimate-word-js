/**
 * Phase 7 的文字事务：只改直接格式树，级联与重排由消费侧在提交后执行。
 * 草稿只复制被改的 run，提交时复制祖先路径；失败不会把半次编辑泄漏给读者。
 * 结构命令在同一块容器内拆段 / 合段，保留直接格式；域和非文字删除仍拒绝。
 * 格式命令只在范围端点拆 run，不合并相邻同格式 run：合并会改掉后方 run 的 id 与片段下标，
 * 已存的批注 / 选区就得跟着映射，而多几个 run 对排版和回写都没有影响。
 */
import type { Block, Body, NodeId, Paragraph, Run, RunContent, TableCell, TableRow } from './nodes.ts';
import { walkBlocks, walkParagraphs } from './nodes.ts';
import { EMPTY_NUMBERING } from './numbering.ts';
import type { ListKind } from './numbering-edit.ts';
import { addListDefinition } from './numbering-edit.ts';
import { buildRunOrder, compareDocPositions, contentLength, rangeOfNode, runEnd, runStart } from './order.ts';
import type { DocPosition, DocRange } from './position.ts';
import type { Indent, NumberingRef, ParagraphSpacing, ParaProps, RunProps } from './props.ts';
import {
  cellColumns,
  cellRect,
  findRow,
  withInsertedColumn,
  withInsertedRow,
  withMergedCells,
  withoutColumns,
  withoutRows,
  withSplitCell,
} from './table-edit.ts';
import type { PositionMove, TextChange, TextChangeSet } from './text-change.ts';
import { invertTextChanges, mapTextRange } from './text-change.ts';

/** 直接字符格式的修改：给值即写入，`null` 删除这一项直接格式、回到样式的值。 */
export type RunPropsPatch = { [K in keyof RunProps]?: RunProps[K] | null };

/**
 * 直接段落格式的修改。`indent` / `spacing` / `numbering` 按字段合并（只改首行缩进不该抹掉左缩进；
 * 只写 `numbering.level` 时沿用样式给的 numId，标题列表降级靠它），`numbering.numId = 0` 取消编号，
 * 其中的 `null` 删除单个字段，整项 `null` 删除整组。段落标记格式走 `setRunProps`。
 * 注意字符单位优先（`firstLineChars` 盖过 `firstLine`）：改 twips 版本时要把对应的 `*Chars` 置 null。
 */
export type ParaPropsPatch = {
  [K in Exclude<keyof ParaProps, 'markRunProps' | 'indent' | 'spacing' | 'numbering'>]?: ParaProps[K] | null;
} & {
  numbering?: { [K in keyof NumberingRef]?: NumberingRef[K] | null } | null;
  indent?: { [K in keyof Indent]?: Indent[K] | null } | null;
  spacing?: { [K in keyof ParagraphSpacing]?: ParagraphSpacing[K] | null } | null;
};

export interface TextTransaction {
  /** 在位置处拆段，返回新段开头；继承段落与字符直接格式。 */
  splitParagraph(position: DocPosition): DocPosition;
  /** 与同一容器内紧随其后的段落合并，保留前段格式。 */
  joinParagraph(paragraphId: NodeId): DocPosition;
  /** 返回插入后的光标位置，可继续传给本事务的下一条命令。 */
  insertText(position: DocPosition, text: string): DocPosition;
  /**
   * 在位置处插入制表位（`w:tab`）或软换行（`w:br`，不结束段落），返回它后面的位置。
   * 文字片段从插入点切开，格式与所在 run 相同；Word 的 Tab / Shift+Enter 就是这两个。
   */
  insertInline(position: DocPosition, kind: InlineKind): DocPosition;
  /** 同一块容器内的文字范围，可跨 run / 段落。返回删除起点。 */
  deleteRange(range: DocRange): DocPosition;
  /**
   * 修改范围内文字的直接字符格式，可跨段落 / 表格 / 分节；端点落在 run 中间时拆出新 run。
   * 覆盖到的段落标记（空段落、或范围越过段尾）同步修改，空段落首次输入与编号跟着它走。
   * 返回拆分后的同一段文字范围，供后续命令继续使用。
   */
  setRunProps(range: DocRange, patch: RunPropsPatch): DocRange;
  /**
   * 修改范围触及的每个段落（含其间的单元格段落）的直接段落格式；折叠范围即光标所在段。
   * 不拆 run、不移动位置，返回原范围。
   */
  setParagraphProps(range: DocRange, patch: ParaPropsPatch): DocRange;
  /**
   * 新增一份列表定义（九级，见 numbering-edit.ts），返回它的 numId，交给 `setParagraphProps` 引用。
   * 定义随本次事务提交与撤销；本次没有段落改动时整次无修改，定义也不留下。
   */
  addList(kind: ListKind): number;
  /**
   * 在位置所在的行（最内层表格）上方 / 下方插一行，照这一行抄结构与格式，每格一个空段落
   * （见 table-edit.ts）。返回新行首格（跳过纵向合并的续格）的空段落位置。
   */
  insertRow(position: DocPosition, side: 'above' | 'below'): DocPosition;
  /**
   * 删掉范围两端所在的行及其间的行，两端须在同一张（最内层）表里；删光就删整张表。
   * 被删行里的位置收拢到返回值：下一行首格、没有下一行就上一行、整表删掉就表后（或表前）的块。
   * 行里有域时拒绝（域可能跨出这一行，删一半配不上对）。
   */
  deleteRows(range: DocRange): DocPosition;
  /**
   * 在位置所在的格（最内层表格）左边 / 右边插一列，宽度照这一列、整表变宽（见 table-edit.ts）。
   * 返回位置所在那一行的新格的空段落位置；这一行的新列被跨列格吃掉时返回原位置。
   */
  insertColumn(position: DocPosition, side: 'left' | 'right'): DocPosition;
  /**
   * 删掉范围两端所在的格覆盖的网格列（两端须在同一张最内层表里）；跨列格缩窄，删空的行删掉，
   * 删光就删整张表。被删格里的位置收拢到返回值：原来那一行里接替它的格、否则左边的格，
   * 整表删掉时同 `deleteRows`。格里有域时拒绝。
   */
  deleteColumns(range: DocRange): DocPosition;
  /**
   * 把范围两端所在的格撑成的矩形（被合并格撑大，见 table-edit.ts）并成一格，两端须在同一张最内层表里。
   * 内容按行、行内从左到右接进首格，空格不贡献段落；多行时下面各行留纵向合并的续格。
   * 返回合并后那一格的开头。矩形里有 `w:gridBefore` / `w:gridAfter` 的空缺时拒绝，只有一格时不修改。
   */
  mergeCells(range: DocRange): DocPosition;
  /**
   * 把位置所在的合并格拆回合并前的样子（跨列拆成一列一格、纵向合并区逐格解开），内容留在原格，
   * 新格各一个空段落。返回原位置；这一格没有合并过时不修改。
   */
  splitCell(position: DocPosition): DocPosition;
}

/**
 * `insertInline` 能插的非文字片段。`pageBreak` 只插 `w:br w:type="page"` 本身 ——
 * Word 的 Ctrl+Enter 还要在它后面拆段，那是命令层（view 的 `pageBreak()`）的事。
 */
export type InlineKind = 'tab' | 'lineBreak' | 'pageBreak';

/**
 * 删除选区时可以删掉的非文字片段。对象（图片）、域界桩不在其中 —— 删了回不来的东西
 * 仍按原先的约定整次拒绝。删掉的片段换成空 text 占住槽位，与文字删空一样不挪后面的 contentIndex。
 */
const DELETABLE = new Set<RunContent['kind']>(['tab', 'break', 'symbol', 'noBreakHyphen', 'softHyphen']);

export interface TextHistoryOptions {
  /** 默认保留 100 个撤销单元。 */
  historyLimit?: number;
  /** 连续输入的间隔上限，默认 1000ms。0 禁止合并。 */
  mergeDelay?: number;
  now?: () => number;
  /** 提交与历史跳转前检查候选快照，抛错时保留模型和历史。 */
  validate?: (body: Body) => void;
}

export interface TextTransactionOptions {
  /** 只有显式标为 input 的单次纯插入可合并；粘贴和程序命令默认独立撤销。 */
  origin?: 'input' | 'command';
}

export interface TextEditor {
  /** 独立于传入树的冻结快照，可交给 resolveBody；不要直接修改。 */
  readonly body: Body;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  /** 同步回调。命令失败（即使在回调中被捕获）也撤销整次事务。无修改返回 undefined。 */
  tx(
    callback: (transaction: TextTransaction) => undefined,
    options?: TextTransactionOptions,
  ): TextChangeSet | undefined;
  undo(): TextChangeSet | undefined;
  redo(): TextChangeSet | undefined;
  /** 光标移动、焦点切换、IME 边界等由输入层显式中断输入合并。 */
  breakHistory(): void;
}

interface RunEntry {
  run: Run;
  paragraph: Paragraph;
  runIndex: number;
  protected: boolean;
}

interface HistoryEntry {
  before: Body;
  after: Body;
  changes: TextChange[];
}

interface InputGroup {
  entry: HistoryEntry;
  last: TextChange;
  time: number;
}

/** 冻结只走新路径，历史快照共享的子树不再遍历。 */
function freeze<T>(value: T): T {
  if (value !== null && typeof value === 'object' && !Object.isFrozen(value)) {
    for (const child of Object.values(value)) freeze(child);
    Object.freeze(value);
  }
  return value;
}

function indexRuns(body: Body): Map<NodeId, RunEntry> {
  const entries = new Map<NodeId, RunEntry>();
  let depth = 0;
  for (const paragraph of walkParagraphs(body)) {
    for (const [runIndex, run] of paragraph.runs.entries()) {
      if (entries.has(run.id)) throw new Error(`重复的 run id：${run.id}`);
      let protectedRun = depth > 0 || run.fieldSimple !== undefined;
      for (const c of run.content) {
        if (c.kind === 'fieldChar') {
          protectedRun = true;
          if (c.charType === 'begin') depth++;
          else if (c.charType === 'end') depth = Math.max(0, depth - 1);
        } else if (c.kind === 'fieldInstruction') protectedRun = true;
      }
      // 域可跨段，不能按段重置 depth；混合界桩与文字的 run 整体保护。
      entries.set(run.id, { run, paragraph, runIndex, protected: protectedRun });
    }
  }
  return entries;
}

function replaceRuns(body: Body, replacements: Map<NodeId, Run>): Body {
  return mapParagraphs(body, (block) => {
    const runs = block.runs.map((r) => replacements.get(r.id) ?? r);
    return runs.some((r, i) => r !== block.runs[i]) ? { ...block, runs } : block;
  });
}

/** 只复制改动段落的祖先路径，未改的节 / 表格与历史快照共享。 */
function mapParagraphs(body: Body, update: (paragraph: Paragraph) => Paragraph): Body {
  function blocks(items: Block[]): Block[] {
    const next = items.map((block): Block => {
      if (block.kind === 'paragraph') return update(block);
      const rows = block.rows.map((row) => {
        const cells = row.cells.map((cell) => {
          const content = blocks(cell.blocks);
          return content === cell.blocks ? cell : { ...cell, blocks: content };
        });
        return cells.some((c, i) => c !== row.cells[i]) ? { ...row, cells } : row;
      });
      return rows.some((r, i) => r !== block.rows[i]) ? { ...block, rows } : block;
    });
    return next.some((b, i) => b !== items[i]) ? next : items;
  }
  return {
    ...body,
    sections: body.sections.map((section) => {
      const content = blocks(section.blocks);
      return content === section.blocks ? section : { ...section, blocks: content };
    }),
  };
}

function textAt(entry: RunEntry, position: DocPosition): string {
  const { contentIndex: ci, offset } = position;
  if (!Number.isInteger(ci) || ci < 0 || !Number.isInteger(offset) || offset < 0) {
    throw new RangeError('文字位置必须是非负整数');
  }
  if (entry.protected) throw new Error('文字事务暂不支持修改域');
  const c = entry.run.content[ci];
  // 已有空 run 有合法位置 {0, 0}；没有 run 的空段落先由 materialize 建立插入 run。
  const text = c?.kind === 'text' ? c.text : entry.run.content.length === 0 && ci === 0 ? '' : undefined;
  if (text === undefined) throw new Error('文字事务只能修改 text 片段');
  if (offset > text.length) throw new RangeError('文字位置超出片段长度');
  const prev = text.charCodeAt(offset - 1);
  const next = text.charCodeAt(offset);
  if (prev >= 0xd800 && prev <= 0xdbff && next >= 0xdc00 && next <= 0xdfff) {
    throw new RangeError('文字位置不能切开 UTF-16 代理对');
  }
  return text;
}

/** 仅应用有变化的项；全部相同返回 undefined，调用方据此不拆 run、不记变更。 */
function patchRunProps(props: RunProps, patch: RunPropsPatch): RunProps | undefined {
  return patchFields(props, patch);
}

/** 缩进与间距是一组独立字段，按字段合并；合并后为空的组整个删掉，不留 `<w:ind/>` 空壳。 */
function patchParaProps(props: ParaProps, patch: ParaPropsPatch): ParaProps | undefined {
  if ('markRunProps' in patch) throw new TypeError('段落标记格式请用 setRunProps');
  const { indent, spacing, numbering, ...rest } = patch;
  let next = patchFields(props, rest) ?? props;
  for (const [key, group] of [
    ['indent', indent],
    ['spacing', spacing],
    ['numbering', numbering],
  ] as const) {
    if (group === undefined || group === null) {
      if (group === null) next = patchFields(next, { [key]: null }) ?? next;
      continue;
    }
    if (typeof group !== 'object' || Array.isArray(group)) throw new TypeError(`${key} 需要属性对象`);
    const merged = patchFields(next[key] ?? {}, group);
    if (merged) next = patchFields(next, { [key]: Object.keys(merged).length ? merged : null }) ?? next;
  }
  return next === props ? undefined : next;
}

function patchFields<T extends object>(props: T, patch: object): T | undefined {
  const next: Record<string, unknown> = { ...(props as Record<string, unknown>) };
  let changed = false;
  for (const [key, value] of Object.entries(patch)) {
    if (value === undefined) continue;
    if (value === null) {
      if (key in next) {
        delete next[key];
        changed = true;
      }
    } else if (JSON.stringify(next[key]) !== JSON.stringify(value)) {
      // 克隆调用方的值：冻结快照不能与外部对象共享可变引用。
      next[key] = structuredClone(value);
      changed = true;
    }
  }
  return changed ? (next as T) : undefined;
}

function samePosition(a: DocPosition, b: DocPosition): boolean {
  return a.nodeId === b.nodeId && a.contentIndex === b.contentIndex && a.offset === b.offset;
}

function changeSet(changes: readonly TextChange[]): TextChangeSet {
  // 不把历史里的数组交给调用方，避免外部修改破坏后续撤销和位置映射。
  return {
    changes: structuredClone([...changes]),
    paragraphIds: [...new Set(changes.flatMap((c) => c.affectedParagraphIds ?? [c.paragraphId]))],
  };
}

export function createTextEditor(source: Body, options: TextHistoryOptions = {}): TextEditor {
  const limit = options.historyLimit ?? 100;
  const delay = options.mergeDelay ?? 1000;
  const now = options.now ?? Date.now;
  if (!Number.isInteger(limit) || limit < 0) throw new RangeError('historyLimit 必须是非负整数');
  if (!Number.isFinite(delay) || delay < 0) throw new RangeError('mergeDelay 必须是非负有限数');
  let body = freeze(structuredClone(source));
  const ids = new Set<string>();
  for (const section of body.sections) {
    ids.add(section.id);
    for (const b of walkBlocks(section.blocks)) {
      ids.add(b.id);
      if (b.kind === 'paragraph') for (const r of b.runs) ids.add(r.id);
      else
        for (const row of b.rows) {
          ids.add(row.id);
          for (const cell of row.cells) ids.add(cell.id);
        }
    }
  }
  let serial = 0;
  function newId(): string {
    let id: string;
    do {
      id = `edit${serial++}`;
    } while (ids.has(id));
    ids.add(id);
    return id;
  }
  indexRuns(body);
  const undo: HistoryEntry[] = [];
  const redo: HistoryEntry[] = [];
  let group: InputGroup | undefined;
  let active = false;

  function assertIdle(): void {
    if (active) throw new Error('事务回调中不能开启事务或操作历史');
  }

  function travel(from: HistoryEntry[], to: HistoryEntry[], backwards: boolean): TextChangeSet | undefined {
    assertIdle();
    group = undefined;
    const entry = from.at(-1);
    if (entry === undefined) return undefined;
    options.validate?.(backwards ? entry.before : entry.after);
    from.pop();
    to.push(entry);
    body = backwards ? entry.before : entry.after;
    return changeSet(backwards ? invertTextChanges(entry.changes) : entry.changes);
  }

  return {
    get body() {
      return body;
    },
    get canUndo() {
      return undo.length > 0;
    },
    get canRedo() {
      return redo.length > 0;
    },
    undo: () => travel(undo, redo, true),
    redo: () => travel(redo, undo, false),
    breakHistory() {
      assertIdle();
      group = undefined;
    },
    tx(callback, txOptions = {}) {
      assertIdle();
      active = true;
      let open = true;
      let failed = false;
      let draft = body;
      let entries = indexRuns(draft);
      const replacements = new Map<NodeId, Run>();
      const changes: TextChange[] = [];

      function flush(): void {
        draft = replaceRuns(draft, replacements);
        replacements.clear();
        entries = indexRuns(draft);
      }
      function paragraphAt(id: NodeId): { paragraph: Paragraph; blocks: Block[]; index: number } {
        function find(blocks: Block[]): ReturnType<typeof paragraphAt> | undefined {
          for (const [index, block] of blocks.entries()) {
            if (block.kind === 'paragraph' && block.id === id) return { paragraph: block, blocks, index };
            if (block.kind === 'table')
              for (const row of block.rows)
                for (const cell of row.cells) {
                  const found = find(cell.blocks);
                  if (found) return found;
                }
          }
          return undefined;
        }
        for (const section of draft.sections) {
          const found = find(section.blocks);
          if (found) return found;
        }
        throw new RangeError(`文档中没有段落：${id}`);
      }
      function replaceParagraphs(id: NodeId, count: number, next: Paragraph[]): void {
        function blocks(items: Block[]): Block[] {
          return items.flatMap((b): Block[] => {
            if (b.id === id) return next;
            const target = items.findIndex((item) => item.id === id);
            if (target >= 0 && items.indexOf(b) > target && items.indexOf(b) < target + count) return [];
            if (b.kind === 'paragraph') return [b];
            return [
              {
                ...b,
                rows: b.rows.map((r) => ({
                  ...r,
                  cells: r.cells.map((c) => ({ ...c, blocks: blocks(c.blocks) })),
                })),
              },
            ];
          });
        }
        draft = { ...draft, sections: draft.sections.map((s) => ({ ...s, blocks: blocks(s.blocks) })) };
        entries = indexRuns(draft);
      }
      function structural(
        paragraphIds: NodeId[],
        moves: PositionMove[],
        inverseMoves = moves.map((m) => ({ from: m.to, to: m.from, length: m.length })),
      ): void {
        changes.push({
          paragraphId: paragraphIds[0] as string,
          nodeId: '',
          contentIndex: 0,
          offset: 0,
          deletedText: '',
          insertedText: '',
          moves,
          inverseMoves,
          affectedParagraphIds: paragraphIds,
        });
      }
      function validatePosition(position: DocPosition): void {
        if (entries.has(position.nodeId)) {
          editableAt(entryAt(position), position);
          return;
        }
        const { paragraph } = paragraphAt(position.nodeId);
        if (paragraph.runs.length || position.contentIndex !== 0 || position.offset !== 0)
          throw new RangeError('空段落位置只能是 {0, 0}');
      }
      function materialize(position: DocPosition): DocPosition {
        if (entries.has(position.nodeId)) return position;
        flush();
        const { paragraph } = paragraphAt(position.nodeId);
        if (paragraph.runs.length || position.contentIndex !== 0 || position.offset !== 0)
          throw new RangeError('空段落位置只能是 {0, 0}');
        const run: Run = {
          kind: 'run',
          id: newId(),
          props: { ...paragraph.props.markRunProps },
          content: [],
        };
        const to = { nodeId: run.id, contentIndex: 0, offset: 0 };
        replaceParagraphs(paragraph.id, 1, [{ ...paragraph, runs: [run] }]);
        structural([paragraph.id], [{ from: position, to, length: 0 }]);
        return to;
      }
      function entryAt(position: DocPosition): RunEntry {
        const entry = entries.get(position.nodeId);
        if (entry === undefined) throw new RangeError(`文档中没有 run：${position.nodeId}`);
        return { ...entry, run: replacements.get(position.nodeId) ?? entry.run };
      }

      /** 可删除的非文字片段上的位置（前 0 / 后 1）；文字位置照旧交给 textAt 检查。 */
      function editableAt(entry: RunEntry, position: DocPosition): void {
        const c = entry.run.content[position.contentIndex];
        if (c === undefined || c.kind === 'text') {
          textAt(entry, position);
          return;
        }
        if (entry.protected) throw new Error('文字事务暂不支持修改域');
        if (!DELETABLE.has(c.kind)) throw new Error('文字事务只能修改 text 片段');
        const { offset } = position;
        if (!Number.isInteger(offset) || offset < 0 || offset > contentLength(c))
          throw new RangeError('文字位置超出片段长度');
      }
      /**
       * 在 run 里插入一个片段，返回它的下标。文字中间就切成「前半 / 片段 / 后半」，
       * 片段边界上就直接插在那里，不造空文字。后面的片段整体后移，位置跟着挪。
       */
      function insertContent(position: DocPosition, item: RunContent): number {
        const entry = entryAt(position);
        editableAt(entry, position);
        const { run } = entry;
        const ci = position.contentIndex;
        const c = run.content[ci];
        const content = [...run.content];
        const moves: PositionMove[] = [];
        let index: number;
        let shift = 1;
        if (c === undefined) {
          index = 0;
          content.push(item);
        } else if (c.kind === 'text' && position.offset > 0 && position.offset < c.text.length) {
          index = ci + 1;
          shift = 2;
          content.splice(ci, 1, { kind: 'text', text: c.text.slice(0, position.offset) }, item, {
            kind: 'text',
            text: c.text.slice(position.offset),
          });
          moves.push({
            from: position,
            to: { nodeId: run.id, contentIndex: ci + 2, offset: 0 },
            length: c.text.length - position.offset,
            afterOnly: true,
          });
        } else {
          index = position.offset === 0 ? ci : ci + 1;
          content.splice(index, 0, item);
        }
        for (let i = shift === 2 ? ci + 1 : index; i < run.content.length; i++)
          moves.push({
            from: { nodeId: run.id, contentIndex: i, offset: 0 },
            to: { nodeId: run.id, contentIndex: i + shift, offset: 0 },
            length: contentLength(run.content[i] as RunContent),
            afterOnly: i === index,
          });
        replacements.set(run.id, { ...run, content });
        flush();
        // 撤销时新片段上的位置回到插入点，不能留在已经不存在的下标上。
        const back = [0, contentLength(item)].map((offset) => ({
          from: { nodeId: run.id, contentIndex: index, offset },
          to: { ...position },
          length: 0,
        }));
        structural([entry.paragraph.id], moves, [
          ...back,
          ...moves.map((m) => ({ from: m.to, to: m.from, length: m.length })),
        ]);
        return index;
      }
      /**
       * 落在制表位 / 换行这类片段上的位置换成能写字的文字位置：紧挨着的文字片段优先，
       * 没有就插一个空文字片段 —— 否则 Tab 之后接着打字会抛「只能修改 text 片段」。
       */
      function textPosition(position: DocPosition): DocPosition {
        position = materialize(position);
        flush();
        const entry = entryAt(position);
        const content = entry.run.content;
        const c = content[position.contentIndex];
        if (c === undefined || c.kind === 'text') return position;
        editableAt(entry, position);
        const ci = position.contentIndex;
        const neighbor = position.offset === 0 ? content[ci - 1] : content[ci + 1];
        if (neighbor?.kind === 'text')
          return position.offset === 0
            ? { nodeId: entry.run.id, contentIndex: ci - 1, offset: neighbor.text.length }
            : { nodeId: entry.run.id, contentIndex: ci + 1, offset: 0 };
        return {
          nodeId: entry.run.id,
          contentIndex: insertContent(position, { kind: 'text', text: '' }),
          offset: 0,
        };
      }

      /** 格式命令也接受域结果与对象旁的位置，只检查坐标本身是否存在。 */
      function checkPosition(at: DocPosition): void {
        const { contentIndex: ci, offset } = at;
        if (!Number.isInteger(ci) || ci < 0 || !Number.isInteger(offset) || offset < 0)
          throw new RangeError('文字位置必须是非负整数');
        const entry = entries.get(at.nodeId);
        if (entry === undefined) {
          validatePosition(at);
          return;
        }
        const c = entry.run.content[ci];
        if (c === undefined) {
          if (entry.run.content.length || ci || offset) throw new RangeError('文字位置超出 run');
          return;
        }
        if (offset > contentLength(c)) throw new RangeError('文字位置超出片段长度');
        if (offset > 0 && offset < contentLength(c) && c.kind !== 'text')
          throw new RangeError('文字位置不能切开非文字片段');
        if (c.kind === 'text') textAt({ ...entry, protected: false }, at);
      }
      /** 在 run 内部位置拆成两段，前段保留原 id；拆点在片段边界时不留空文字片段。 */
      function splitRun(run: Run, at: DocPosition, log: TextChange[]): [Run, Run] {
        let ci = at.contentIndex;
        let offset = at.offset;
        const current = run.content[ci];
        if (current && offset === contentLength(current)) {
          ci++;
          offset = 0;
        }
        const right: Run = { ...run, id: newId(), content: [] };
        const moves: PositionMove[] = [];
        let left: Run;
        if (offset === 0) {
          left = { ...run, content: run.content.slice(0, ci) };
          right.content = run.content.slice(ci);
        } else {
          const text = (run.content[ci] as { text: string }).text;
          left = {
            ...run,
            content: [...run.content.slice(0, ci), { kind: 'text', text: text.slice(0, offset) }],
          };
          right.content = [{ kind: 'text', text: text.slice(offset) }, ...run.content.slice(ci + 1)];
          moves.push({
            from: { nodeId: run.id, contentIndex: ci, offset },
            to: { nodeId: right.id, contentIndex: 0, offset: 0 },
            length: text.length - offset,
            afterOnly: true,
          });
          ci++;
        }
        const shift = run.content.length - right.content.length;
        for (let i = ci; i < run.content.length; i++)
          moves.push({
            from: { nodeId: run.id, contentIndex: i, offset: 0 },
            to: { nodeId: right.id, contentIndex: i - shift, offset: 0 },
            length: contentLength(run.content[i] as Run['content'][number]),
          });
        const paragraphId = (entries.get(run.id) as RunEntry).paragraph.id;
        log.push({
          paragraphId,
          nodeId: '',
          contentIndex: 0,
          offset: 0,
          deletedText: '',
          insertedText: '',
          moves,
          inverseMoves: moves.map((m) => ({ from: m.to, to: m.from, length: m.length })),
          affectedParagraphIds: [paragraphId],
        });
        return [left, right];
      }

      function splice(position: DocPosition, count: number, insertedText: string): void {
        const entry = entryAt(position);
        const text = textAt(entry, position);
        const deletedText = text.slice(position.offset, position.offset + count);
        if (deletedText === insertedText) return;
        const content = [...entry.run.content];
        content[position.contentIndex] = {
          kind: 'text',
          text: text.slice(0, position.offset) + insertedText + text.slice(position.offset + count),
        };
        replacements.set(entry.run.id, { ...entry.run, content });
        changes.push({ ...position, paragraphId: entry.paragraph.id, deletedText, insertedText });
      }

      function command<T>(action: () => T): T {
        if (!open) throw new Error('事务已结束');
        try {
          return action();
        } catch (error) {
          failed = true;
          throw error;
        }
      }

      const transaction: TextTransaction = {
        splitParagraph(position) {
          return command(() => {
            position = textPosition(position);
            const entry = entryAt(position);
            const text = textAt(entry, position);
            // 域可能跨 run / 段；拆段不能切断其结构。
            for (const r of entry.paragraph.runs)
              if (entries.get(r.id)?.protected) throw new Error('暂不支持拆分包含域的段落');
            const { run, paragraph, runIndex } = entry;
            const ci = position.contentIndex;
            const left: Run = {
              ...run,
              content: [...run.content.slice(0, ci), { kind: 'text', text: text.slice(0, position.offset) }],
            };
            const right: Run = {
              ...run,
              id: newId(),
              content: [{ kind: 'text', text: text.slice(position.offset) }, ...run.content.slice(ci + 1)],
            };
            const next: Paragraph = {
              ...paragraph,
              id: newId(),
              runs: [right, ...paragraph.runs.slice(runIndex + 1)],
            };
            const to = { nodeId: right.id, contentIndex: 0, offset: 0 };
            const moves: PositionMove[] = [
              { from: position, to, length: text.length - position.offset, afterOnly: true },
            ];
            for (let i = ci + 1; i < run.content.length; i++) {
              const c = run.content[i];
              moves.push({
                from: { nodeId: run.id, contentIndex: i, offset: 0 },
                to: { nodeId: right.id, contentIndex: i - ci, offset: 0 },
                length: c?.kind === 'text' ? c.text.length : 1,
              });
            }
            replaceParagraphs(paragraph.id, 1, [
              { ...paragraph, runs: [...paragraph.runs.slice(0, runIndex), left] },
              next,
            ]);
            structural([paragraph.id, next.id], moves);
            return to;
          });
        },
        joinParagraph(paragraphId) {
          return command(() => {
            flush();
            const { paragraph, blocks, index } = paragraphAt(paragraphId);
            const next = blocks[index + 1];
            if (next?.kind !== 'paragraph') throw new Error('只能合并同一节或单元格内相邻的段落');
            for (const p of [paragraph, next])
              for (const r of p.runs)
                if (entries.get(r.id)?.protected) throw new Error('暂不支持合并包含域的段落');
            const boundary = (paragraph.runs.length
              ? rangeOfNode(paragraph)?.end
              : next.runs.length
                ? rangeOfNode(next)?.start
                : undefined) ?? { nodeId: paragraph.id, contentIndex: 0, offset: 0 };
            const moves: PositionMove[] = [];
            if (!paragraph.runs.length && next.runs.length)
              moves.push({
                from: { nodeId: paragraph.id, contentIndex: 0, offset: 0 },
                to: boundary,
                length: 0,
              });
            if (!next.runs.length)
              moves.push({ from: { nodeId: next.id, contentIndex: 0, offset: 0 }, to: boundary, length: 0 });
            replaceParagraphs(paragraph.id, 2, [{ ...paragraph, runs: [...paragraph.runs, ...next.runs] }]);
            // 合并边界的原位置仍有效，撤销不应把它挪到被恢复的空段落。
            structural([paragraph.id, next.id], moves, []);
            return boundary;
          });
        },
        insertText(position, text) {
          return command(() => {
            // Unicode 模式按码点匹配：合法代理对不会落进代理项区间。
            if (typeof text !== 'string' || /[\r\n\t\uD800-\uDFFF]/u.test(text)) {
              throw new Error('insertText 只接受无换行、无制表位的有效 Unicode 文字');
            }
            if (text === '') {
              validatePosition(position);
              return { ...position };
            }
            position = textPosition(position);
            splice(position, 0, text);
            return { ...position, offset: position.offset + text.length };
          });
        },
        insertInline(position, kind) {
          return command(() => {
            const item: RunContent =
              kind === 'tab'
                ? { kind: 'tab' }
                : kind === 'lineBreak'
                  ? { kind: 'break', breakType: 'line' }
                  : kind === 'pageBreak'
                    ? { kind: 'break', breakType: 'page' }
                    : (() => {
                        throw new TypeError(`未知的行内片段：${String(kind)}`);
                      })();
            position = materialize(position);
            flush();
            // 单元格里的分页符 Word 会把整张表从这一行拆开，拆表还没有（插行 / 删行有了，见 table-edit.ts）；
            // 照插的话布局只在格内截断、表格不拆，屏幕上看不出分页。
            if (kind === 'pageBreak') {
              const host = paragraphAt(entryAt(position).paragraph.id).blocks;
              if (!draft.sections.some((s) => s.blocks === host)) throw new Error('表格里不能插分页符');
            }
            const index = insertContent(position, item);
            return { nodeId: position.nodeId, contentIndex: index, offset: 1 };
          });
        },
        deleteRange(range) {
          return command(() => {
            if (samePosition(range.start, range.end)) {
              validatePosition(range.start);
              return { ...range.start };
            }
            const start = materialize(range.start);
            const end = samePosition(range.start, range.end) ? start : materialize(range.end);
            flush();
            const first = entryAt(start);
            const last = entryAt(end);
            editableAt(first, start);
            editableAt(last, end);
            if (first.paragraph !== last.paragraph) {
              const a = paragraphAt(first.paragraph.id);
              const b = paragraphAt(last.paragraph.id);
              if (
                a.blocks !== b.blocks ||
                a.index >= b.index ||
                a.blocks.slice(a.index, b.index + 1).some((p) => p.kind !== 'paragraph')
              )
                throw new Error('跨段删除不能跨表格、单元格或分节');
              const paragraphs = a.blocks.slice(a.index, b.index + 1) as Paragraph[];
              for (const [i, p] of paragraphs.entries()) {
                const r = rangeOfNode(p);
                if (r !== undefined)
                  transaction.deleteRange({
                    start: i === 0 ? start : r.start,
                    end: i === paragraphs.length - 1 ? end : r.end,
                  });
              }
              for (let i = 1; i < paragraphs.length; i++) transaction.joinParagraph(first.paragraph.id);
              return start;
            }
            const order =
              first.runIndex - last.runIndex ||
              start.contentIndex - end.contentIndex ||
              start.offset - end.offset;
            if (order > 0) throw new RangeError('删除范围的起点必须在终点之前');
            if (samePosition(start, end)) return { ...start };
            // 保留所有片段与 run 的槽位，后方样式与原有位置不会因为删除而换号。
            for (let ri = first.runIndex; ri <= last.runIndex; ri++) {
              const run = first.paragraph.runs[ri] as Run;
              const current = entryAt({ nodeId: run.id, contentIndex: 0, offset: 0 });
              const from = ri === first.runIndex ? start.contentIndex : 0;
              const to =
                ri === last.runIndex ? end.contentIndex : Math.max(0, current.run.content.length - 1);
              for (let ci = from; ci <= to; ci++) {
                const offset = ri === first.runIndex && ci === from ? start.offset : 0;
                const pos = { nodeId: run.id, contentIndex: ci, offset };
                const entry = entryAt(pos);
                const c = entry.run.content[ci];
                if (c !== undefined && c.kind !== 'text') {
                  editableAt(entry, pos);
                  const stop = ri === last.runIndex && ci === to ? end.offset : contentLength(c);
                  if (stop > offset) {
                    const content = [...entry.run.content];
                    content[ci] = { kind: 'text', text: '' };
                    replacements.set(run.id, { ...entry.run, content });
                    // 片段后面的位置（偏移 1）落回空文字的 0；撤销走快照，位置原样有效。
                    structural(
                      [entry.paragraph.id],
                      [{ from: { ...pos, offset: 1 }, to: { ...pos, offset: 0 }, length: 0 }],
                      [],
                    );
                  }
                  continue;
                }
                const text = textAt(entry, pos);
                const stop = ri === last.runIndex && ci === to ? end.offset : text.length;
                splice(pos, stop - offset, '');
              }
            }
            return { ...start };
          });
        },
        setRunProps(range, patch) {
          return command(() => {
            if (patch === null || typeof patch !== 'object' || Array.isArray(patch))
              throw new TypeError('setRunProps 需要属性对象');
            flush();
            for (const at of [range.start, range.end]) checkPosition(at);
            const order = buildRunOrder(draft);
            const cmp = (a: DocPosition, b: DocPosition) => compareDocPositions(order, a, b) as number;
            if (cmp(range.start, range.end) > 0) throw new RangeError('格式范围的起点必须在终点之前');
            const { start, end } = range;
            const collapsed = cmp(start, end) === 0;
            const paragraphs = [...walkParagraphs(draft)];
            const indexOf = (at: DocPosition) =>
              paragraphs.findIndex((p) => p.id === at.nodeId || p.runs.some((r) => r.id === at.nodeId));
            const last = indexOf(end);
            const updates = new Map<NodeId, Paragraph>();
            const local: TextChange[] = [];
            for (let i = indexOf(start); i <= last; i++) {
              const paragraph = paragraphs[i] as Paragraph;
              let props = paragraph.props;
              // 空段落只有段落标记；非空段落的标记只在范围越过段尾时被选中，折叠光标不改标记。
              if (!paragraph.runs.length || (!collapsed && i < last)) {
                const mark = patchRunProps(props.markRunProps ?? {}, patch);
                if (mark) props = { ...props, markRunProps: mark };
              }
              const runs: Run[] = [];
              for (const run of paragraph.runs) {
                const rs = runStart(run);
                const re = runEnd(run);
                const covered =
                  !collapsed &&
                  (run.content.length
                    ? cmp(re, start) > 0 && cmp(rs, end) < 0
                    : cmp(start, rs) <= 0 && cmp(rs, end) < 0);
                const next = covered ? patchRunProps(run.props, patch) : undefined;
                if (!next) {
                  runs.push(run);
                  continue;
                }
                // 先切末端：起点坐标仍对前半段有效；位置映射按记录顺序逐次应用。
                let head = run;
                let tail: Run | undefined;
                if (end.nodeId === run.id && cmp(end, re) < 0) [head, tail] = splitRun(head, end, local);
                let before: Run | undefined;
                let middle = head;
                if (start.nodeId === run.id && cmp(start, rs) > 0)
                  [before, middle] = splitRun(head, start, local);
                if (before) runs.push(before);
                runs.push({ ...middle, props: next });
                if (tail) runs.push(tail);
              }
              if (props !== paragraph.props || runs.some((r, ri) => r !== paragraph.runs[ri]))
                updates.set(paragraph.id, { ...paragraph, props, runs });
            }
            if (!updates.size) return structuredClone(range);
            draft = mapParagraphs(draft, (p) => updates.get(p.id) ?? p);
            entries = indexRuns(draft);
            changes.push(...local);
            structural([...updates.keys()], []);
            return mapTextRange(range, { changes: local, paragraphIds: [] });
          });
        },
        setParagraphProps(range, patch) {
          return command(() => {
            if (patch === null || typeof patch !== 'object' || Array.isArray(patch))
              throw new TypeError('setParagraphProps 需要属性对象');
            flush();
            for (const at of [range.start, range.end]) checkPosition(at);
            const order = buildRunOrder(draft);
            if ((compareDocPositions(order, range.start, range.end) as number) > 0)
              throw new RangeError('格式范围的起点必须在终点之前');
            const paragraphs = [...walkParagraphs(draft)];
            const indexOf = (at: DocPosition) =>
              paragraphs.findIndex((p) => p.id === at.nodeId || p.runs.some((r) => r.id === at.nodeId));
            const updates = new Map<NodeId, Paragraph>();
            for (let i = indexOf(range.start); i <= indexOf(range.end); i++) {
              const paragraph = paragraphs[i] as Paragraph;
              const props = patchParaProps(paragraph.props, patch);
              if (props) updates.set(paragraph.id, { ...paragraph, props });
            }
            if (updates.size) {
              draft = mapParagraphs(draft, (p) => updates.get(p.id) ?? p);
              entries = indexRuns(draft);
              structural([...updates.keys()], []);
            }
            return structuredClone(range);
          });
        },
        insertRow(position, side) {
          return command(() => {
            if (side !== 'above' && side !== 'below') throw new TypeError(`未知的插入方向：${String(side)}`);
            flush();
            const hit = findRow(draft, position.nodeId);
            if (hit === undefined) throw new Error('位置不在表格里');
            const { table, row } = withInsertedRow(hit.table, hit.rowIndex, side, newId);
            draft = hit.replace(table, () => {
              throw new Error('插行不会删表');
            });
            entries = indexRuns(draft);
            const added = row.cells.flatMap((c) => c.blocks.map((b) => b.id));
            // 正向不挪任何位置；撤销时新行整个消失，落在新行里的光标收拢回插行的地方（Word 也是这样）
            const back = { ...position };
            structural(
              added,
              [],
              added.map((id) => ({
                from: { nodeId: id, contentIndex: 0, offset: 0 },
                to: back,
                length: 0,
                collapse: true,
              })),
            );
            return cellStart(row);
          });
        },
        deleteRows(range) {
          return command(() => {
            flush();
            const a = findRow(draft, range.start.nodeId);
            const b = findRow(draft, range.end.nodeId);
            if (a === undefined || b === undefined || a.table.id !== b.table.id)
              throw new Error('只能删除同一张表里的行');
            const from = Math.min(a.rowIndex, b.rowIndex);
            const to = Math.max(a.rowIndex, b.rowIndex);
            const removed = a.table.rows.slice(from, to + 1);
            const paragraphs = removed.flatMap((r) =>
              r.cells.flatMap((c) =>
                [...walkBlocks(c.blocks)].filter((x): x is Paragraph => x.kind === 'paragraph'),
              ),
            );
            for (const p of paragraphs)
              for (const r of p.runs)
                if (entries.get(r.id)?.protected) throw new Error('暂不支持删除包含域的行');
            const next = withoutRows(a.table, from, to);
            const { after, before } = a.neighbour();
            let filler: Paragraph | undefined;
            draft = a.replace(next, () => {
              filler = { kind: 'paragraph', id: newId(), props: {}, runs: [] };
              return filler;
            });
            entries = indexRuns(draft);
            let target: DocPosition | undefined;
            if (next !== undefined) {
              const row = next.rows[from] ?? next.rows[from - 1];
              if (row !== undefined) target = cellStart(row);
            } else if (filler !== undefined) target = { nodeId: filler.id, contentIndex: 0, offset: 0 };
            else if (after !== undefined) target = rangeOfNode(after)?.start;
            else if (before !== undefined) target = rangeOfNode(before)?.end;
            if (target === undefined) throw new Error('删行后找不到光标位置');
            // 撤销整棵树回到删之前，原位置本来就有效，不必反向映射
            structural(
              paragraphs.map((p) => p.id),
              collapseInto(paragraphs, target),
              [],
            );
            return { ...target };
          });
        },
        insertColumn(position, side) {
          return command(() => {
            if (side !== 'left' && side !== 'right') throw new TypeError(`未知的插入方向：${String(side)}`);
            flush();
            const hit = findRow(draft, position.nodeId);
            if (hit === undefined) throw new Error('位置不在表格里');
            const own = hit.table.rows[hit.rowIndex]?.cells[hit.cellIndex];
            const { table, cells } = withInsertedColumn(hit.table, hit.rowIndex, hit.cellIndex, side, newId);
            draft = hit.replace(table, () => {
              throw new Error('插列不会删表');
            });
            entries = indexRuns(draft);
            const added = cells.flatMap((c) => c.blocks.map((b) => b.id));
            const back = { ...position };
            // 没有新格（整表的新列都被跨列格吃掉）时不算修改 —— 可网格变宽了，照样记一笔让撤销能退回去
            structural(
              added.length ? added : (own?.blocks.slice(0, 1).map((b) => b.id) ?? []),
              [],
              added.map((id) => ({
                from: { nodeId: id, contentIndex: 0, offset: 0 },
                to: back,
                length: 0,
                collapse: true,
              })),
            );
            const row = table.rows[hit.rowIndex] as TableRow;
            const mine = row.cells.find((c) => cells.includes(c));
            const range = mine === undefined ? undefined : rangeOfNode(mine);
            return range === undefined ? { ...position } : { ...range.start };
          });
        },
        deleteColumns(range) {
          return command(() => {
            flush();
            const a = findRow(draft, range.start.nodeId);
            const b = findRow(draft, range.end.nodeId);
            if (a === undefined || b === undefined || a.table.id !== b.table.id)
              throw new Error('只能删除同一张表里的列');
            const ca = cellColumns(a.table.rows[a.rowIndex] as TableRow, a.cellIndex);
            const cb = cellColumns(b.table.rows[b.rowIndex] as TableRow, b.cellIndex);
            const from = Math.min(ca.start, cb.start);
            const to = Math.max(ca.end, cb.end);
            const { table: next, removed } = withoutColumns(a.table, from, to);
            const paragraphs = removed.flatMap((c) =>
              [...walkBlocks(c.blocks)].filter((x): x is Paragraph => x.kind === 'paragraph'),
            );
            for (const p of paragraphs)
              for (const r of p.runs)
                if (entries.get(r.id)?.protected) throw new Error('暂不支持删除包含域的列');
            const { after, before } = a.neighbour();
            let filler: Paragraph | undefined;
            draft = a.replace(next, () => {
              filler = { kind: 'paragraph', id: newId(), props: {}, runs: [] };
              return filler;
            });
            entries = indexRuns(draft);
            let target: DocPosition | undefined;
            if (next !== undefined) {
              const own = a.table.rows[a.rowIndex]?.id;
              const row =
                next.rows.find((r) => r.id === own) ?? next.rows[Math.min(a.rowIndex, next.rows.length - 1)];
              if (row !== undefined) {
                // 接替被删列的是从 `from` 开始的那一格，删的是最右几列就退到左边最后一格
                const i = row.cells.findIndex((_, k) => cellColumns(row, k).end > from);
                const cell = row.cells[i >= 0 ? i : row.cells.length - 1];
                target = cell === undefined ? undefined : rangeOfNode(cell)?.start;
              }
            } else if (filler !== undefined) target = { nodeId: filler.id, contentIndex: 0, offset: 0 };
            else if (after !== undefined) target = rangeOfNode(after)?.start;
            else if (before !== undefined) target = rangeOfNode(before)?.end;
            if (target === undefined) throw new Error('删列后找不到光标位置');
            structural(
              paragraphs.map((p) => p.id),
              collapseInto(paragraphs, target),
              [],
            );
            return { ...target };
          });
        },
        mergeCells(range) {
          return command(() => {
            flush();
            const a = findRow(draft, range.start.nodeId);
            const b = findRow(draft, range.end.nodeId);
            if (a === undefined || b === undefined || a.table.id !== b.table.id)
              throw new Error('只能合并同一张表里的格');
            const rect = cellRect(a.table, a, b);
            if (rect === undefined) throw new Error('选中的格撑不成完整的矩形（有跳过的网格列）');
            const merged = withMergedCells(a.table, rect, newId);
            if (merged === undefined) return structuredClone(range.start);
            const { table, head, dropped, added: fresh } = merged;
            draft = a.replace(table, () => {
              throw new Error('合并不会删表');
            });
            entries = indexRuns(draft);
            const target = rangeOfNode(head)?.start;
            if (target === undefined) throw new Error('合并后找不到光标位置');
            // 搬进首格的段落 id 不变，位置照旧有效；只有丢掉的空段落收拢、撤销时续格里新造的空段落收拢回首格
            structural(
              [...head.blocks.map((x) => x.id), ...fresh],
              collapseInto(dropped, target),
              fresh.map((id) => ({
                from: { nodeId: id, contentIndex: 0, offset: 0 },
                to: target,
                length: 0,
                collapse: true,
              })),
            );
            return { ...target };
          });
        },
        splitCell(position) {
          return command(() => {
            flush();
            const hit = findRow(draft, position.nodeId);
            if (hit === undefined) throw new Error('位置不在表格里');
            const split = withSplitCell(hit.table, hit.rowIndex, hit.cellIndex, newId);
            if (split === undefined) return { ...position };
            draft = hit.replace(split.table, () => {
              throw new Error('拆分不会删表');
            });
            entries = indexRuns(draft);
            const added = split.cells.flatMap((c) => c.blocks.map((x) => x.id));
            const back = { ...position };
            structural(
              added,
              [],
              added.map((id) => ({
                from: { nodeId: id, contentIndex: 0, offset: 0 },
                to: back,
                length: 0,
                collapse: true,
              })),
            );
            return { ...position };
          });
        },
        addList(kind) {
          return command(() => {
            if (kind !== 'bullet' && kind !== 'decimal')
              throw new TypeError(`未知的列表类型：${String(kind)}`);
            flush();
            const { numbering, numId } = addListDefinition(draft.numbering ?? EMPTY_NUMBERING, kind);
            // 不记变更：定义本身不动任何位置，没有段落引用它时不该产生撤销单元
            draft = { ...draft, numbering };
            return numId;
          });
        },
      };

      try {
        const result: unknown = callback(transaction);
        open = false;
        if (result !== null && typeof result === 'object' && 'then' in result) {
          // JS 可绕过同步返回类型；接住后续异步失败，避免拒绝后又产生未处理的 Promise。
          void Promise.resolve(result).catch(() => {});
          throw new TypeError('事务回调必须同步执行');
        }
        if (failed) throw new Error('事务包含失败命令，已全部回滚');
        if (changes.length === 0) return undefined;
        const after = freeze(replaceRuns(draft, replacements));
        options.validate?.(after);
        const time = now();
        if (!Number.isFinite(time)) throw new RangeError('now 必须返回有限时间');
        const only = changes.length === 1 ? changes[0] : undefined;
        const input =
          txOptions.origin === 'input' && only?.moves === undefined && only?.deletedText === ''
            ? only
            : undefined;
        const merge =
          input !== undefined &&
          group !== undefined &&
          delay > 0 &&
          time >= group.time &&
          time - group.time <= delay &&
          samePosition(input, { ...group.last, offset: group.last.offset + group.last.insertedText.length });
        let entry: HistoryEntry;
        if (merge && group !== undefined) {
          entry = group.entry;
          entry.after = after;
          entry.changes.push(...changes);
        } else {
          entry = { before: body, after, changes };
          if (limit > 0) undo.push(entry);
          if (undo.length > limit) undo.shift();
        }
        body = after;
        redo.length = 0;
        group = input !== undefined && limit > 0 ? { entry, last: input, time } : undefined;
        return changeSet(changes);
      } finally {
        open = false;
        active = false;
      }
    },
  };
}

/** 一行的首格（纵向合并的续格不显示内容，跳过）第一个段落的开头 */
/** 这些段落里每一个位置都收拢到 `dest`（被删掉的行 / 列：按偏移平移会指到不存在的字） */
function collapseInto(paragraphs: readonly Paragraph[], dest: DocPosition): PositionMove[] {
  return paragraphs.flatMap((p): PositionMove[] =>
    p.runs.length === 0
      ? [{ from: { nodeId: p.id, contentIndex: 0, offset: 0 }, to: dest, length: 0, collapse: true }]
      : p.runs.flatMap((r) =>
          (r.content.length ? r.content : [undefined]).map((c, i) => ({
            from: { nodeId: r.id, contentIndex: i, offset: 0 },
            to: dest,
            length: c === undefined ? 0 : contentLength(c),
            collapse: true,
          })),
        ),
  );
}

function cellStart(row: TableRow): DocPosition {
  const cell: TableCell | undefined = row.cells.find((c) => c.vMerge !== 'continue') ?? row.cells[0];
  const range = cell === undefined ? undefined : rangeOfNode(cell);
  if (range === undefined) throw new Error('行里没有可放光标的段落');
  return { ...range.start };
}
