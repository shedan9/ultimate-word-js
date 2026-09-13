/**
 * Phase 7 的文字事务：只改直接格式树，级联与重排由消费侧在提交后执行。
 * 草稿只复制被改的 run，提交时复制祖先路径；失败不会把半次编辑泄漏给读者。
 * 结构命令在同一块容器内拆段 / 合段，保留直接格式；域和非文字删除仍拒绝。
 */
import type { Block, Body, NodeId, Paragraph, Run } from './nodes.ts';
import { walkBlocks, walkParagraphs } from './nodes.ts';
import { rangeOfNode } from './order.ts';
import type { DocPosition, DocRange } from './position.ts';
import type { PositionMove, TextChange, TextChangeSet } from './text-change.ts';
import { invertTextChanges } from './text-change.ts';

export interface TextTransaction {
  /** 在位置处拆段，返回新段开头；继承段落与字符直接格式。 */
  splitParagraph(position: DocPosition): DocPosition;
  /** 与同一容器内紧随其后的段落合并，保留前段格式。 */
  joinParagraph(paragraphId: NodeId): DocPosition;
  /** 返回插入后的光标位置，可继续传给本事务的下一条命令。 */
  insertText(position: DocPosition, text: string): DocPosition;
  /** 同一块容器内的文字范围，可跨 run / 段落。返回删除起点。 */
  deleteRange(range: DocRange): DocPosition;
}

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
  function blocks(items: Block[]): Block[] {
    const next = items.map((block): Block => {
      if (block.kind === 'paragraph') {
        const runs = block.runs.map((r) => replacements.get(r.id) ?? r);
        return runs.some((r, i) => r !== block.runs[i]) ? { ...block, runs } : block;
      }
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
          textAt(entryAt(position), position);
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

      function command(action: () => DocPosition): DocPosition {
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
            position = materialize(position);
            flush();
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
            position = materialize(position);
            splice(position, 0, text);
            return { ...position, offset: position.offset + text.length };
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
            textAt(first, start);
            textAt(last, end);
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
                const text = textAt(entryAt(pos), pos);
                const stop = ri === last.runIndex && ci === to ? end.offset : text.length;
                splice(pos, stop - offset, '');
              }
            }
            return { ...start };
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
