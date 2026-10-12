/**
 * 批注（`comments.xml`）：谁、什么时候、说了什么、批的是哪一段字。
 *
 * 与脚注同理分两处放：
 *
 * - **内容**按 `w:id` 摊成一张表（`LoadedComments.comments`），不挂在正文节点上 ——
 *   批注罩着的字会被编辑，内容只有一份
 * - **罩着哪些字**是正文 run 上的标记（`RunNode.comments`），跟着编辑走：拆 run、拆段自动带着，
 *   前面插了字不用映射。起止位置用的时候现算（`listComments`）
 *
 * 回复与「已解决」不在 `comments.xml` 里，在 Word 2013 起的 `commentsExtended.xml`：
 * 它按批注**最后一段的 `w14:paraId`** 认批注（不按 `w:id`），`paraIdParent` 指向被回复的那条。
 * 没有这个部件（WPS、老版 Word）时一律当顶层、未解决。
 *
 * **画不画由宿主定**：Word 的批注框住在页边，打开它整页要缩小、重排 —— 这一层只给数据，
 * 宿主拿 `range` 交给 `view.decorate()` / `view.overlay()` 自己画（api.md §8）。
 * 批注的编辑（增删改、回复、标记解决）没做，回写原样照搬 `comments.xml`。
 */
import type { DiagnosticSink } from '@uw/core';
import type { OpcPackage } from '@uw/ooxml';
import { attr, children, RelType } from '@uw/ooxml';
import type { Block, Body, DocumentBody, PropSet, RunNode } from './nodes.ts';
import { paragraphText, walkBlocks, walkParagraphs } from './nodes.ts';
import { contentLength, runEnd, runStart } from './order.ts';
import { parseComments } from './parse-body.ts';
import type { DocPosition, DocRange } from './position.ts';
import { runText } from './revisions.ts';

/** 一条批注解析完的样子 */
export interface CommentContent {
  /** `w:id`，与正文 run 上的 `comments` 标记对得上 */
  id: string;
  author?: string;
  initials?: string;
  /** `w:date` 原样（Word 写的是本地时间却带着 `Z`，不换算） */
  date?: string;
  /** 被回复的那条的 `w:id`（`commentsExtended.xml` 的 `paraIdParent`）；顶层批注缺席 */
  parent?: string;
  /** 「已解决」（`w15:done="1"`） */
  done: boolean;
  /** 批注正文。节点 id 带 `cm:` 前缀，与正文、页眉页脚、注都不撞车 */
  blocks: Block[];
}

export interface LoadedComments {
  /** `comments.xml` 的部件名；文档没有批注部件时缺席 */
  part?: string;
  /** 按 `w:id` 索引，顺序同文件 */
  comments: Record<string, CommentContent>;
}

export function emptyComments(): LoadedComments {
  return { comments: {} };
}

/** 解析批注部件与它的扩展部件，缺哪个都不是错误 */
export function parseCommentParts(pkg: OpcPackage, diagnostics: DiagnosticSink): LoadedComments {
  const part = pkg.partNameByRelType(RelType.COMMENTS);
  if (part === undefined) return emptyComments();
  const parsed = parseComments(pkg.xml(part), diagnostics, part, 'cm:');
  const extended = commentsExtended(pkg);
  // paraId → w:id：扩展部件里的父子关系是按 paraId 写的
  const byParaId = new Map<string, string>();
  for (const c of parsed) if (c.paraId !== undefined) byParaId.set(c.paraId, c.id);

  const comments: Record<string, CommentContent> = {};
  for (const c of parsed) {
    const out: CommentContent = { id: c.id, done: false, blocks: c.blocks };
    if (c.author !== undefined) out.author = c.author;
    if (c.initials !== undefined) out.initials = c.initials;
    if (c.date !== undefined) out.date = c.date;
    const ex = c.paraId === undefined ? undefined : extended.get(c.paraId);
    if (ex !== undefined) {
      out.done = ex.done;
      const parent = ex.parentParaId === undefined ? undefined : byParaId.get(ex.parentParaId);
      // 指向自己或指不到的父：当顶层，不让面板里出现一条挂在空气上的回复
      if (parent !== undefined && parent !== c.id) out.parent = parent;
    }
    comments[c.id] = out;
  }
  return { part, comments };
}

interface CommentEx {
  parentParaId?: string;
  done: boolean;
}

/** `commentsExtended.xml`：paraId → 父批注的 paraId + 是否已解决 */
function commentsExtended(pkg: OpcPackage): Map<string, CommentEx> {
  const out = new Map<string, CommentEx>();
  const part = pkg.partNameByRelType(RelType.COMMENTS_EXTENDED);
  if (part === undefined) return out;
  for (const el of children(pkg.xml(part).root)) {
    if (el.name !== 'w15:commentEx') continue;
    const paraId = attr(el, 'w15:paraId');
    if (paraId === undefined) continue;
    const done = attr(el, 'w15:done');
    const ex: CommentEx = { done: done === '1' || done === 'true' };
    const parent = attr(el, 'w15:paraIdParent');
    if (parent !== undefined) ex.parentParaId = parent;
    out.set(paraId, ex);
  }
  return out;
}

/** 给批注面板的一条：批注本身 + 它此刻罩着正文的哪一段 */
export interface CommentSpan extends Omit<CommentContent, 'blocks'> {
  /** 批注正文的纯文字，段与段之间 `\n` */
  text: string;
  /**
   * 罩着的字（被批注的原文），段与段之间 `\n`，被删除的修订不算。点批注（范围是空的）为空串
   */
  quote: string;
  /**
   * 罩着的范围：罩着的头一个有内容的 run 开头到末一个的结尾。点批注是折叠的范围，落在旁边那个字上。
   * 罩着的字全删光了且旁边也没字（或连引用标记所在的 run 也没了），或批注挂在脚注里，答 undefined
   */
  range?: DocRange;
}

/**
 * 全文的批注，按**锚点的文档序**（同一处起头的按 `w:id` 在文件里的先后）；
 * 在正文里找不到锚点的排在最后，按文件顺序。
 *
 * 吃**可编辑的那棵树**：标记只挂在那里（级联树不带）。
 */
export function listComments(body: Body, loaded: LoadedComments): CommentSpan[] {
  const anchors = commentAnchors(body);
  const order = Object.keys(loaded.comments);
  const spans = Object.values(loaded.comments).map((c): CommentSpan => {
    const { blocks, ...rest } = c;
    const span: CommentSpan = { ...rest, text: blocksText(blocks), quote: '' };
    const anchor = anchors.get(c.id);
    if (anchor?.start !== undefined && anchor.end !== undefined) {
      span.range = { start: anchor.start, end: anchor.end };
      span.quote = anchor.quote;
    } else if (anchor?.point !== undefined) {
      span.range = { start: anchor.point, end: anchor.point };
    }
    return span;
  });
  const rank = (s: CommentSpan): number => anchors.get(s.id)?.rank ?? Number.POSITIVE_INFINITY;
  return spans.sort((a, b) => rank(a) - rank(b) || order.indexOf(a.id) - order.indexOf(b.id));
}

interface Anchor {
  /** 首个带标记的 run 在文档序里是第几个 —— 排序用 */
  rank: number;
  /** 罩着的、**有内容**的头一个与末一个 run 的两端；只罩着空 run（点批注）时缺席 */
  start?: DocPosition;
  end?: DocPosition;
  /** 点批注落在哪：标记所在空 run 旁边那个有内容的字（前一个的结尾，没有就后一个的开头） */
  point?: DocPosition;
  quote: string;
  /** 上一次往 quote 里写字是在哪一段：换段时补 `\n` */
  paragraph: string;
}

/**
 * 为什么只认**有内容**的 run：引用标记所在的 run（`w:commentReference`）在树上是个空 run，
 * 布局里没有它的片段，落在它身上的位置命中测试与光标都答不出来 —— 范围的终点、点批注的锚
 * 都会画不出（`view.overlay` 找不到锚点就藏起来）。所以两端取罩着的有内容的 run，
 * 一个都没有时取旁边那个字
 */
function commentAnchors<S extends PropSet>(body: DocumentBody<S>): Map<string, Anchor> {
  const out = new Map<string, Anchor>();
  let rank = 0;
  for (const p of walkParagraphs(body)) {
    // 本段里上一个有内容的 run 的结尾；点批注先挂在它身上，没有的话等下一个有内容的 run
    let before: DocPosition | undefined;
    const waiting: Anchor[] = [];
    for (const run of p.runs) {
      rank++;
      const filled = run.content.some((c) => contentLength(c) > 0);
      if (filled) {
        for (const a of waiting.splice(0)) a.point = runStart(run);
      }
      for (const id of run.comments ?? []) {
        let hit = out.get(id);
        if (hit === undefined) {
          hit = { rank, quote: '', paragraph: p.id };
          out.set(id, hit);
        }
        if (!filled) {
          if (hit.start === undefined && hit.point === undefined) {
            if (before !== undefined) hit.point = before;
            else waiting.push(hit);
          }
          continue;
        }
        hit.start ??= runStart(run);
        hit.end = runEnd(run);
        if (hit.paragraph !== p.id) {
          // 范围从上一段的空 run（段末的界桩）起头时，quote 不该以换行开头
          if (hit.quote !== '') hit.quote += '\n';
          hit.paragraph = p.id;
        }
        if (run.revision?.kind !== 'delete') hit.quote += runText(run as RunNode<S>);
      }
      if (filled) before = runEnd(run);
    }
  }
  return out;
}

function blocksText(blocks: readonly Block[]): string {
  const lines: string[] = [];
  for (const b of walkBlocks(blocks)) if (b.kind === 'paragraph') lines.push(paragraphText(b));
  return lines.join('\n');
}
