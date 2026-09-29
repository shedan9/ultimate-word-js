/**
 * 脚注与尾注：`footnotes.xml` / `endnotes.xml` 的内容，加上它们怎么编号。
 *
 * 三件事分在三处，理由与页眉页脚一样：
 *
 * - **内容**按 `w:id` 摊成一张表（`LoadedNotes.footnotes`），不挂在引用它的 run 上 ——
 *   引用可以被移动、复制，内容只有一份；挂上去就成了跨节点引用（原则 1.1）
 * - **引用**是 run 里的一个片段（`RunContent` 的 `noteReference`），跟着正文编辑走
 * - **显示的号**不在文件里：Word 按文档序现数（删掉第 2 条，后面的全部往前挪），
 *   所以数号在布局那一侧（layout 的 notes.ts），与 SEQ 同理 —— 模型只给编号规则
 *
 * 分隔线那几条（`w:type="separator"` 等）不是注，是**脚注区的装饰**：它们的段落定了
 * 分隔线那一行占多高，线本身画多长多粗是布局与渲染的事（没有真值，见各自的 uncalibrated.ts）。
 */
import type { DiagnosticSink } from '@uw/core';
import type { OpcPackage, XmlElement } from '@uw/ooxml';
import { child, RelType } from '@uw/ooxml';
import type { Block, NoteNumbering, ResolvedBlock } from './nodes.ts';
import { parseNotes } from './parse-body.ts';
import { attrInt, attrOf, enumVal, put } from './xml-values.ts';

/** 一条注（或一条分隔线）解析完的样子。两棵树的分工与页眉页脚一致 */
export interface NoteContent {
  id: string;
  /** 部件名，诊断与回写用 */
  part: string;
  blocks: Block[];
  resolved: ResolvedBlock[];
}

/** 一个部件（脚注或尾注）的全部内容 */
export interface NotePart {
  /** 部件名（`word/footnotes.xml`）。文档没有这个部件时缺席 */
  part?: string;
  /** 普通的注，按 `w:id` 索引 */
  notes: Record<string, NoteContent>;
  /** `w:type="separator"`：正文与注之间那条短线所在的段落 */
  separator?: NoteContent;
  /** `w:type="continuationSeparator"`：注跨页续排时，续页上那条通栏线 */
  continuationSeparator?: NoteContent;
  /** 文档级（`settings.xml` 的 `w:footnotePr` / `w:endnotePr`）编号规则，节上的盖在它上面 */
  numbering: NoteNumbering;
}

export interface LoadedNotes {
  footnotes: NotePart;
  endnotes: NotePart;
}

const RESTARTS = ['continuous', 'eachSect', 'eachPage'] as const;

/**
 * `w:footnotePr` / `w:endnotePr` → 编号规则。一个字段都没写时返回 undefined，
 * 让 `SectionProps` 上缺席就是「这一节没覆盖」。
 */
export function parseNoteNumbering(el: XmlElement | undefined): NoteNumbering | undefined {
  if (el === undefined) return undefined;
  const out: NoteNumbering = {};
  put(out, 'numFmt', attrOf(child(el, 'w:numFmt'), 'w:val'));
  put(out, 'numStart', attrInt(child(el, 'w:numStart'), 'w:val'));
  put(out, 'numRestart', enumVal(attrOf(child(el, 'w:numRestart'), 'w:val'), RESTARTS));
  put(out, 'pos', attrOf(child(el, 'w:pos'), 'w:val'));
  return Object.keys(out).length === 0 ? undefined : out;
}

/** 没有注的文档（绝大多数）给这个，布局一个判断就跳过 */
export function emptyNotes(): LoadedNotes {
  return { footnotes: { notes: {}, numbering: {} }, endnotes: { notes: {}, numbering: {} } };
}

/**
 * 解析两个部件（缺哪个都不是错误）。**只收解析树**，级联由调用方做 ——
 * 与页眉页脚一样要先扫一遍 HYPERLINK 域，那张表是全文档一张的。
 *
 * id 前缀 `fn:` / `en:`：注里的节点 id 不能与正文、页眉页脚撞车（域求值、命中测试都按 id 认 run）。
 */
export function parseNoteParts(pkg: OpcPackage, diagnostics: DiagnosticSink): LoadedNotes {
  const settingsName = pkg.partNameByRelType(RelType.SETTINGS);
  const settings = settingsName === undefined ? undefined : pkg.xml(settingsName).root;
  const out = emptyNotes();
  out.footnotes.numbering = parseNoteNumbering(settings && child(settings, 'w:footnotePr')) ?? {};
  out.endnotes.numbering = parseNoteNumbering(settings && child(settings, 'w:endnotePr')) ?? {};
  collect(pkg, RelType.FOOTNOTES, 'fn:', out.footnotes, diagnostics);
  collect(pkg, RelType.ENDNOTES, 'en:', out.endnotes, diagnostics);
  return out;
}

function collect(
  pkg: OpcPackage,
  relType: string,
  idPrefix: string,
  into: NotePart,
  diagnostics: DiagnosticSink,
): void {
  const part = pkg.partNameByRelType(relType);
  if (part === undefined) return;
  into.part = part;
  for (const note of parseNotes(pkg.xml(part), diagnostics, part, idPrefix)) {
    const content: NoteContent = { id: note.id, part, blocks: note.blocks, resolved: [] };
    if (note.type === 'separator') into.separator = content;
    else if (note.type === 'continuationSeparator') into.continuationSeparator = content;
    // continuationNotice（「接下页」提示）默认是空的，且只有跨页续排才出现 —— 那条还没做
    else if (note.type === 'normal') into.notes[note.id] = content;
  }
}

/** 全部要级联的块，调用方逐条 `resolveBlocks` 后写回 `resolved` */
export function allNoteContents(notes: LoadedNotes): NoteContent[] {
  const out: NoteContent[] = [];
  for (const part of [notes.footnotes, notes.endnotes]) {
    out.push(...Object.values(part.notes));
    if (part.separator !== undefined) out.push(part.separator);
    if (part.continuationSeparator !== undefined) out.push(part.continuationSeparator);
  }
  return out;
}
