/**
 * 脚注 / 尾注的号 —— 按文档序现数。
 *
 * 文件里**不存**号（`w:footnoteReference` 只有 `w:id`），Word 打开时按引用在正文里的先后
 * 从头数一遍：删掉第 2 条，后面的全部往前挪一个。所以这件事与 SEQ 一样在布局这一侧做，
 * 每趟排版前数一遍（一遍走完整棵树的 run，比排版本身便宜两个数量级）。
 *
 * 编号规则分两层（`w:footnotePr`）：`settings.xml` 里是文档级的，`w:sectPr` 里的逐字段盖上去。
 * 默认值是规范给的：脚注 `decimal`、尾注 `lowerRoman`，都从 1 起、连续编号。
 *
 * **每页重新编号**（`w:numRestart="eachPage"`，只对脚注有效 —— 规范不许尾注这么设）：
 * 号要等分页才知道，号的宽度又会改分页（「⑩」换成「①」可能让一行多收一个字），
 * 所以与域求值同一套办法迭代：先按连续编号排一趟，读出每个引用落在哪一页（`referencePages`），
 * 按页重数、再排，直到**每个引用所在的页不再变**（`page.ts` 的 `layoutDocument`）。
 * 判据选「引用在哪一页」而不是「号的文字」：号是页的函数，页不变号就不变，反过来不成立 ——
 * 两页都只有一条脚注时号都是 ①，内容却可能在两页之间挪来挪去。
 * 迭代放在 `layoutDocument` 里面而不是接进 `layoutDocumentWithFields` 的循环：
 * 两件事的收敛判据互不相干，拼成一个判据反而要处理「页码收敛了号没收敛」的交叉情形；
 * 代价是开着这个开关、又有页码域的文档每趟域求值里多排一两趟。
 */
import type { DiagnosticSink } from '@uw/core';
import type {
  NodeId,
  NoteNumbering,
  NoteType,
  ResolvedBlock,
  ResolvedBody,
  ResolvedParagraph,
  SectionProps,
} from '@uw/model';
import { formatNumber, walkBlocks } from '@uw/model';
import { noteKey } from './items.ts';
import type { DocumentLayout } from './page.ts';
import type { BlockLayout } from './table.ts';

/**
 * 每页重新编号最多排几趟（含第一趟）。与 `MAX_FIELD_PASSES` 同一个数、同一个理由 ——
 * 正常两趟就收敛（第一趟连续编号、第二趟按页重数），五趟还不收敛只会是来回跳
 */
export const MAX_NOTE_PASSES = 5;

/** 一条注（或分隔线）的内容，只要级联完的块。直接传 `LoadedDocument.notes` 即可 */
export interface NoteContentSource {
  resolved: readonly ResolvedBlock[];
}

export interface NotePartSource {
  notes: Readonly<Record<string, NoteContentSource>>;
  separator?: NoteContentSource;
  continuationSeparator?: NoteContentSource;
  numbering: NoteNumbering;
}

/** `LayoutDocumentOptions.notes` 的形状，与 `@uw/model` 的 `LoadedNotes` 结构兼容 */
export interface NoteSource {
  footnotes: NotePartSource;
  endnotes: NotePartSource;
}

const DEFAULT_FORMAT: Record<NoteType, string> = { footnote: 'decimal', endnote: 'lowerRoman' };

/**
 * 数号：run id → 这个 run 里那个号显示什么（正文的引用与注内容开头的 `noteMark` 都在里面）。
 *
 * 同一条注被引用两次（极少见，手工改 XML 才会有）时，**每次引用各占一个号**，
 * 注内容开头的号取第一次的 —— Word 在界面上造不出这种局面，没有可以对照的行为。
 *
 * `pages`（引用 run → 物理页序，上一趟排出来的）只给每页重新编号的脚注用：换页就归到 `numStart`。
 * 不给（第一趟）就按连续编号数 —— 猜「全在同一页」也行，但连续编号的号只宽不窄，
 * 第一趟排出来的页只会比实际的多占地方、不会少，下一趟往回收比往外挤稳当。
 * 不在表里的引用（号排不出片段）当作与前一个引用同页。
 */
export function noteLabels(
  body: ResolvedBody,
  source: NoteSource,
  _diagnostics?: DiagnosticSink,
  pages?: ReadonlyMap<NodeId, number>,
): Map<string, string> {
  const out = new Map<string, string>();
  const noteNumber = new Map<string, string>();
  const counters: Record<NoteType, number> = { footnote: 0, endnote: 0 };
  const state: NumberState = { counters, pages, lastPage: undefined };

  body.sections.forEach((section, index) => {
    for (const type of ['footnote', 'endnote'] as const) {
      const rule = numberingOf(type, source, section.props);
      // 第一节从 numStart 起；之后的节只有 eachSect 才归位
      if (index === 0 || rule.numRestart === 'eachSect') counters[type] = (rule.numStart ?? 1) - 1;
    }
    for (const block of walkBlocks(section.blocks)) {
      if (block.kind !== 'paragraph') continue;
      numberParagraph(block, section.props, source, state, out, noteNumber);
    }
  });

  // 注内容开头的号（`noteMark`）显示所在那一条的号
  for (const type of ['footnote', 'endnote'] as const) {
    const part = type === 'footnote' ? source.footnotes : source.endnotes;
    for (const [id, note] of Object.entries(part.notes)) {
      const label = noteNumber.get(noteKey(type, id));
      if (label === undefined) continue;
      for (const block of walkBlocks(note.resolved)) {
        if (block.kind !== 'paragraph') continue;
        for (const run of block.runs) {
          if (run.content.some((c) => c.kind === 'noteMark')) out.set(run.id, label);
        }
      }
    }
  }
  return out;
}

interface NumberState {
  counters: Record<NoteType, number>;
  pages: ReadonlyMap<NodeId, number> | undefined;
  /** 上一个脚注引用落在哪一页（每页重新编号看它换没换页） */
  lastPage: number | undefined;
}

function numberParagraph(
  p: ResolvedParagraph,
  section: SectionProps,
  source: NoteSource,
  state: NumberState,
  out: Map<string, string>,
  noteNumber: Map<string, string>,
): void {
  const { counters } = state;
  for (const run of p.runs) {
    // 隐藏的引用不占号：Word 里隐藏文字不打印，脚注跟着不出现，号也不数它
    if (run.props.hidden) continue;
    for (const c of run.content) {
      if (c.kind !== 'noteReference' || c.customMark === true) continue;
      const rule = numberingOf(c.noteType, source, section);
      if (c.noteType === 'footnote' && state.pages !== undefined) {
        const page = state.pages.get(run.id) ?? state.lastPage;
        if (rule.numRestart === 'eachPage' && page !== state.lastPage)
          counters.footnote = (rule.numStart ?? 1) - 1;
        state.lastPage = page;
      }
      counters[c.noteType] += 1;
      const label = formatNumber(counters[c.noteType], rule.numFmt ?? DEFAULT_FORMAT[c.noteType]);
      out.set(run.id, label);
      const key = noteKey(c.noteType, c.noteId);
      if (!noteNumber.has(key)) noteNumber.set(key, label);
    }
  }
}

/** 文档级规则盖上节的规则（逐字段） */
function numberingOf(type: NoteType, source: NoteSource, section: SectionProps): NoteNumbering {
  const doc = type === 'footnote' ? source.footnotes.numbering : source.endnotes.numbering;
  const own = type === 'footnote' ? section.footnotePr : section.endnotePr;
  return { ...doc, ...own };
}

/**
 * 每一节引到了哪几条**尾注**（`w:id`，按引用的先后、一条只收一次）—— 文末那一摞就按这个顺序排。
 *
 * 与数号同一套取舍：隐藏的引用不算（它不打印，尾注跟着不出现）；自定义标记的**算** ——
 * 它不占号，内容照样要排出来（号是用户自己打在 `customMarkFollows` 后面那个字）。
 * 表格里的引用也算：`walkBlocks` 会下钻到单元格。
 */
export function endnoteIdsBySection(body: ResolvedBody): string[][] {
  return body.sections.map((section) => {
    const out: string[] = [];
    for (const block of walkBlocks(section.blocks)) {
      if (block.kind !== 'paragraph') continue;
      for (const run of block.runs) {
        if (run.props.hidden) continue;
        for (const c of run.content) {
          if (c.kind === 'noteReference' && c.noteType === 'endnote' && !out.includes(c.noteId)) {
            out.push(c.noteId);
          }
        }
      }
    }
    return out;
  });
}

/**
 * 尾注排在哪儿：`docEnd`（默认，全文最后）还是 `sectEnd`（每节末尾）。
 * 与编号规则一样文档级打底、节上的盖上去 —— Word 的界面只给文档级的一个开关，
 * 节上写 `w:pos` 的只有手改的 XML
 */
export function endnotePosition(source: NoteSource, section: SectionProps): 'docEnd' | 'sectEnd' {
  return numberingOf('endnote', source, section).pos === 'sectEnd' ? 'sectEnd' : 'docEnd';
}

/** 有没有哪一节的脚注设成了每页重新编号 —— 没有就不必迭代 */
export function restartsEachPage(body: ResolvedBody, source: NoteSource): boolean {
  return body.sections.some((s) => numberingOf('footnote', source, s.props).numRestart === 'eachPage');
}

/** 占号的脚注引用所在的 run（与 `noteLabels` 同一套取舍：隐藏的与自定义标记的不算） */
export function footnoteReferenceRuns(body: ResolvedBody): Set<NodeId> {
  const out = new Set<NodeId>();
  for (const section of body.sections) {
    for (const block of walkBlocks(section.blocks)) {
      if (block.kind !== 'paragraph') continue;
      for (const run of block.runs) {
        if (run.props.hidden) continue;
        const counted = run.content.some(
          (c) => c.kind === 'noteReference' && c.noteType === 'footnote' && c.customMark !== true,
        );
        if (counted) out.add(run.id);
      }
    }
  }
  return out;
}

/**
 * 每个引用 run 的号排在了哪一页（物理页序）。认的是号的片段（`field` 标记 + `runId`）——
 * `LineLayout.notes` 只记引到哪条注，同一条注被引两次时分不出是哪个 run。
 * 表格里的一路下钻到格内，重复的表头跳过（它的引用在表头第一次出现的那一页已经记过）。
 */
export function referencePages(layout: DocumentLayout, runs: ReadonlySet<NodeId>): Map<NodeId, number> {
  const out = new Map<NodeId, number>();
  if (runs.size === 0) return out;
  const visitLines = (
    lines: readonly { fragments: readonly { runId: NodeId; field?: true }[] }[],
    page: number,
  ) => {
    for (const line of lines) {
      for (const f of line.fragments) {
        if (f.field === true && runs.has(f.runId) && !out.has(f.runId)) out.set(f.runId, page);
      }
    }
  };
  const visitCells = (blocks: readonly BlockLayout[], page: number): void => {
    for (const b of blocks) {
      if (b.kind === 'paragraph') visitLines(b.layout.lines, page);
      else for (const r of b.layout.rows) for (const c of r.cells) visitCells(c.blocks, page);
    }
  };
  for (const page of layout.pages) {
    for (const b of page.blocks) {
      if (b.kind === 'paragraph')
        visitLines(
          b.lines.map((l) => l.line),
          page.index,
        );
      else
        for (const r of b.rows)
          if (r.repeated !== true) for (const c of r.row.cells) visitCells(c.blocks, page.index);
    }
  }
  return out;
}

export function samePages(a: ReadonlyMap<NodeId, number>, b: ReadonlyMap<NodeId, number>): boolean {
  if (a.size !== b.size) return false;
  for (const [id, page] of a) if (b.get(id) !== page) return false;
  return true;
}
