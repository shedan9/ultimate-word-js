/**
 * 修订痕迹（track changes）的**显示**。编辑修订（接受 / 拒绝、修订模式下打字）是开发计划 §5 的非目标。
 *
 * 原来解析时把 `w:del` 整块跳过、`w:ins` 当透明容器，等于永远显示「最终状态、无标记」——
 * 公文审阅稿里看不出谁改了什么。现在被删的字也进树（`RunNode.revision`），
 * 显示哪一版是**级联**那一步定的，三种视图对应 Word「审阅」里的三档：
 *
 * | 视图 | 插入的字 | 删除的字 |
 * |---|---|---|
 * | `final`（默认） | 照常显示 | 折成 `hidden`，不占位 |
 * | `markup` | 作者色 + 下划线 | 作者色 + 删除线，**占位** |
 * | `original` | 折成 `hidden` | 照常显示 |
 *
 * 默认 `final` 不是随手挑的：它就是原来的行为，也就是 Word 2013 起「简单标记」看到的版式
 * （简单标记只在页边多一条红竖线，字的位置与最终状态一字不差）。`markup` 是「所有标记」
 * 关掉批注框之后的内嵌显示 —— 批注框（balloon）要在页边另排一栏，那是另一件事。
 *
 * 为什么折成 `hidden` 而不是在布局里再判一次：`w:vanish` 那条路已经被查找、复制、
 * 排版、命中测试、可选文本层各自走通了，「看不见的字不算」在每一处都对；
 * 修订再开一条路，就得在每一处再学一遍。
 *
 * **段落标记与整行**的修订（`ParagraphNode.markRevision` / `TableRowNode.revision`，2026-10-12）改的是结构不是格式，
 * 也在级联那一步折（resolve-body.ts 的 `blockList` / `table`）：这一版里不存在的段落标记让本段并进下一段，
 * 不存在的行整行不进树。「所有标记」下两者照常显示 —— Word 给被删的行涂底色、在段落标记处画 ¶，都没做。
 *
 * 没做（写下来免得以为做了）：格式修订（`w:rPrChange` / `w:pPrChange`，显示的是改后的格式，没有标记）、
 * 批注框、「所有标记」下行与段落标记的标记画法、单元格级的修订（`w:tcPr/w:cellIns|cellDel`）。
 */
import type {
  BlockNode,
  Body,
  DocumentBody,
  NodeId,
  ParagraphNode,
  PropSet,
  RunNode,
  RunRevision,
  TableRowNode,
} from './nodes.ts';
import { walkBlocks } from './nodes.ts';
import { contentLength, rangeOfNode, runEnd, runStart } from './order.ts';
import type { DocRange } from './position.ts';
import type { ResolvedRunProps, RevisionMark } from './props.ts';

export type RevisionView = 'final' | 'markup' | 'original';

/** 级联要知道的全部：显示哪一版、每位作者什么颜色 */
export interface RevisionDisplay {
  view: RevisionView;
  /** 作者 → 六位十六进制色。缺了的作者落到调色板第一色 */
  colors: Readonly<Record<string, string>>;
}

export const DEFAULT_REVISION_DISPLAY: RevisionDisplay = Object.freeze({
  view: 'final',
  colors: Object.freeze({}),
});

/**
 * 作者配色，按作者在文档里**第一次出现**的顺序轮着取。
 *
 * ⚠️ **没有真值**：Word「按作者」配色的确切色值与轮换顺序没量过（本机没有 Word），
 * 这里只保证「第一位是红、不同作者颜色不同、深到能在白纸上看清」。
 * 钉死办法：一份三位作者各插一段、各删一段的 docx，切到「所有标记」导 PDF，
 * 读 `truth.json` 里每个片段的颜色（需要先给 `extract-truth.ts` 加一路 `fillColor`）。
 * 另外 Word 的配色是**按本机会话**分配的（同一份文档换台电脑颜色可能换），
 * 所以真值也只能钉死「第一次打开时」的那一种。
 */
export const REVISION_AUTHOR_COLORS: readonly string[] = [
  'C00000',
  '1F4E9F',
  '2E7D32',
  '7B1FA2',
  'C55A11',
  '00838F',
  'AD1457',
  '5D4037',
];

/**
 * 文档里出现过的修订作者，按第一次出现的顺序（正文在前，再按调用方给的顺序看页眉页脚 / 注 / 文本框）。
 * 作者缺席的修订记成空串 —— 规范要求写 `w:author`，第三方生成器不一定写
 */
export function revisionAuthors<S extends PropSet>(bodies: readonly DocumentBody<S>[]): string[] {
  const seen = new Set<string>();
  for (const body of bodies) {
    for (const section of body.sections) {
      for (const node of revisionNodes(section.blocks)) {
        if (node.kind === 'row') {
          if (node.revision !== undefined) seen.add(node.revision.author ?? '');
        } else {
          for (const r of node.runs) if (r.revision !== undefined) seen.add(r.revision.author ?? '');
          if (node.markRevision !== undefined) seen.add(node.markRevision.author ?? '');
        }
      }
    }
  }
  return [...seen];
}

/** 行的修订在行首，段落标记在段尾；两处都与文字走同一份文档序，不能另扫一遍排到全文末尾 */
function* revisionNodes<S extends PropSet>(
  blocks: readonly BlockNode<S>[],
): Generator<ParagraphNode<S> | TableRowNode<S>> {
  for (const block of blocks) {
    if (block.kind === 'paragraph') {
      yield block;
    } else {
      for (const row of block.rows) {
        yield row;
        for (const cell of row.cells) yield* revisionNodes(cell.blocks);
      }
    }
  }
}

export function revisionDisplay(view: RevisionView, authors: readonly string[]): RevisionDisplay {
  const colors: Record<string, string> = {};
  authors.forEach((a, i) => {
    colors[a] = REVISION_AUTHOR_COLORS[i % REVISION_AUTHOR_COLORS.length] as string;
  });
  return { view, colors };
}

/** 这一版里它**不存在**：最终状态下被删的、原始状态下插入的。「所有标记」下两者都在 */
export function revisionHidden(revision: RunRevision | undefined, display: RevisionDisplay): boolean {
  if (revision === undefined) return false;
  return display.view === 'final'
    ? revision.kind === 'delete'
    : display.view === 'original' && revision.kind === 'insert';
}

/**
 * 把一个 run 的修订折进它的级联结果。没有修订、或这一版里它就是普通文字时**原样返回同一个对象**
 * （下游按身份缓存字形样式，见 `@uw/layout` 的 `styleOf`）。
 */
export function applyRevision(
  props: ResolvedRunProps,
  revision: RunRevision | undefined,
  display: RevisionDisplay,
): ResolvedRunProps {
  if (revision === undefined) return props;
  switch (display.view) {
    case 'final':
      return revision.kind === 'delete' ? { ...props, hidden: true } : props;
    case 'original':
      return revision.kind === 'insert' ? { ...props, hidden: true } : props;
    case 'markup': {
      const mark: RevisionMark = {
        kind: revision.kind,
        color: display.colors[revision.author ?? ''] ?? (REVISION_AUTHOR_COLORS[0] as string),
      };
      if (revision.move) mark.move = true;
      return { ...props, revision: mark };
    }
  }
}

/** 一处修订：文字、段落标记与整行各自列出，不能因作者相同就把结构修订吞进文字修订 */
export interface RevisionSpan extends RunRevision {
  target: 'text' | 'paragraph-mark' | 'row';
  /** 所在段落；整行修订取本行的第一个段落（包括嵌套表格与空段落） */
  paragraphId: NodeId;
  /** 只在整行修订上带，范围覆盖本行全部单元格 */
  rowId?: NodeId;
  /** 文字照 runText；段落标记写 \n；整行以 \t 分格、\n 分段 / 分嵌套表格的行 */
  text: string;
  /** 段落标记没有字符位置，以本段末尾的折叠范围表示；隐藏修订的范围不保证能在当前视图定位 */
  range: DocRange;
}

/**
 * 全文的修订，文档序。给「修订面板」用：谁、什么时候、插了 / 删了什么、在哪儿。
 *
 * 吃**可编辑的那棵树**，不吃级联树：级联树上最终状态里被删的 run 也还在（只是 hidden），
 * 但这里要的是「文件里记着什么」，与这一刻显示哪一版无关。整行修订在格里的文字修订前列出，
 * 段落标记在本段的文字修订后列出。没有任何段落的异常空行没有可定位范围，不列（作者仍收）。
 */
export function listRevisions(body: Body): RevisionSpan[] {
  const out: RevisionSpan[] = [];
  for (const section of body.sections) {
    for (const node of revisionNodes(section.blocks)) {
      if (node.kind === 'row') {
        if (node.revision === undefined) continue;
        const range = rangeOfNode(node);
        const first = node.cells
          .flatMap((cell) => [...walkBlocks(cell.blocks)])
          .find((block) => block.kind === 'paragraph');
        if (range !== undefined && first !== undefined) {
          out.push({
            ...node.revision,
            target: 'row',
            rowId: node.id,
            paragraphId: first.id,
            text: node.cells.map((cell) => blocksText(cell.blocks)).join('\t'),
            range,
          });
        }
        continue;
      }
      paragraphRevisions(node, out);
    }
  }
  return out;
}

function paragraphRevisions<S extends PropSet>(p: ParagraphNode<S>, out: RevisionSpan[]): void {
  let open: RevisionSpan | undefined;
  for (const r of p.runs) {
    const rev = r.revision;
    if (rev === undefined) {
      open = undefined;
      continue;
    }
    if (open !== undefined && sameRevision(open, rev)) {
      open.text += runText(r);
      open.range = { start: open.range.start, end: runEnd(r) };
      continue;
    }
    open = {
      ...rev,
      target: 'text',
      paragraphId: p.id,
      text: runText(r),
      range: { start: runStart(r), end: runEnd(r) },
    };
    out.push(open);
  }
  if (p.markRevision !== undefined) {
    const range = rangeOfNode(p);
    if (range === undefined) return;
    out.push({
      ...p.markRevision,
      target: 'paragraph-mark',
      paragraphId: p.id,
      text: '\n',
      range: { start: { ...range.end }, end: { ...range.end } },
    });
  }
}

/** 行摘要读原始树：最终状态会隐藏的文字也要列出来，不能借 paragraphText 跳过被删的 run */
function blocksText<S extends PropSet>(blocks: readonly BlockNode<S>[]): string {
  return blocks
    .map((block) => {
      if (block.kind === 'paragraph') return block.runs.map(runText).join('');
      return block.rows.map((row) => row.cells.map((cell) => blocksText(cell.blocks)).join('\t')).join('\n');
    })
    .join('\n');
}

function sameRevision(a: RunRevision, b: RunRevision): boolean {
  return a.kind === b.kind && a.move === b.move && a.author === b.author && a.date === b.date;
}

/** 修订里看得见的那部分文字：文字原样，制表位 / 换行写成 `\t` / `\n`，对象写 U+FFFC，界桩与指令不算 */
export function runText<S extends PropSet>(run: RunNode<S>): string {
  let s = '';
  for (const c of run.content) {
    if (c.kind === 'text') s += c.text;
    else if (c.kind === 'tab') s += '\t';
    else if (c.kind === 'break') s += '\n';
    else if (contentLength(c) > 0) s += c.kind === 'symbol' ? c.char : '￼';
  }
  return s;
}
