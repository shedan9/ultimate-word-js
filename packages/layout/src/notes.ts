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
 * **没做**：`w:numRestart="eachPage"`（每页重新编号）—— 号要等分页才知道，号的宽度又会改分页，
 * 得进域求值那种迭代。现在按连续编号数并记一条诊断；中文论文里这个开关不少见，
 * 做的时候把「每页第一条归 1」接进 `layoutDocumentWithFields` 的收敛循环即可。
 */
import type { DiagnosticSink } from '@uw/core';
import type {
  NoteNumbering,
  NoteType,
  ResolvedBlock,
  ResolvedBody,
  ResolvedParagraph,
  SectionProps,
} from '@uw/model';
import { formatNumber, walkBlocks } from '@uw/model';
import { noteKey } from './items.ts';

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
 */
export function noteLabels(
  body: ResolvedBody,
  source: NoteSource,
  diagnostics?: DiagnosticSink,
): Map<string, string> {
  const out = new Map<string, string>();
  const noteNumber = new Map<string, string>();
  const counters: Record<NoteType, number> = { footnote: 0, endnote: 0 };
  let warnedEachPage = false;

  body.sections.forEach((section, index) => {
    for (const type of ['footnote', 'endnote'] as const) {
      const rule = numberingOf(type, source, section.props);
      // 第一节从 numStart 起；之后的节只有 eachSect 才归位
      if (index === 0 || rule.numRestart === 'eachSect') counters[type] = (rule.numStart ?? 1) - 1;
      if (rule.numRestart === 'eachPage' && !warnedEachPage) {
        warnedEachPage = true;
        diagnostics?.warn('note-restart-each-page', '脚注设成了「每页重新编号」，暂按连续编号显示');
      }
    }
    for (const block of walkBlocks(section.blocks)) {
      if (block.kind !== 'paragraph') continue;
      numberParagraph(block, section.props, source, counters, out, noteNumber);
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

function numberParagraph(
  p: ResolvedParagraph,
  section: SectionProps,
  source: NoteSource,
  counters: Record<NoteType, number>,
  out: Map<string, string>,
  noteNumber: Map<string, string>,
): void {
  for (const run of p.runs) {
    // 隐藏的引用不占号：Word 里隐藏文字不打印，脚注跟着不出现，号也不数它
    if (run.props.hidden) continue;
    for (const c of run.content) {
      if (c.kind !== 'noteReference' || c.customMark === true) continue;
      counters[c.noteType] += 1;
      const rule = numberingOf(c.noteType, source, section);
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
