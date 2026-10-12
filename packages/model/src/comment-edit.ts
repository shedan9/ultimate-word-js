/**
 * 批注的编辑（Word「审阅」里的新建 / 回复 / 改 / 解决 / 删除）用到的纯函数；命令本身在 text-transaction.ts。
 *
 * 一条批注在模型里是两半（comments.ts）：内容在 `Body.comments`，罩着哪些字是 run 上的 `comments` 标记。
 * 两半都在可编辑的树上，所以事务的草稿与撤销快照把它们一起带着 —— 这是把内容从 `LoadedDocument`
 * 挪到 `Body` 上的原因。
 *
 * 三处容易搞错：
 * ① **新 id 不复用删掉的**：回写按 `w:id` 认「原文里有没有这条」（comments-writer.ts），删掉 3 号再新建一条
 *    还叫 3 号，回写会把它当成原来那条、留着原来的界桩，标记却挪了地方。所以编号接着
 *    「加载时见过的最大 id」往下走、只增不减（`commentIdSeed`，编辑器建起来时取一次）；
 * ② **回复罩着与顶层那条相同的字**：Word 写回复时在同一处再写一对界桩。批注线程只有一层，
 *    回复的回复挂到顶层上（`threadRoot`）；
 * ③ **解决的是整条线程**：Word 的「解决」按钮在线程上，顶层与回复的 `w15:done` 一起改。
 *    ⚠️ 没有真值（本机没有 Word）：只从 Word 存出来的文件里看到过「整条线程都是 done」，
 *    没有对过「只解决一条回复」的界面行为。
 */
import type { CommentContent } from './comments.ts';
import type { Block, Body, Paragraph, Run, RunContent } from './nodes.ts';
import { walkParagraphs } from './nodes.ts';

/** 新批注 / 回复的输入。`text` 里 `\n` 分段、`\t` 写成制表位 */
export interface CommentDraft {
  text: string;
  author?: string;
  initials?: string;
  /** `w:date` 原样写进去。Word 写的是本地时间却带着 `Z`（comments.ts），门面按这个惯例补 */
  date?: string;
}

/**
 * 下一个可用的批注 id：内容表的 key 与 run 上标记里出现过的数字 id 的最大值 + 1。
 * 标记也要看：文件里可能有界桩指着 `comments.xml` 里没有的 id（批注被别的工具删了一半），
 * 新批注占了那个号，回写时那对孤儿界桩就成了它的
 */
export function commentIdSeed(body: Body): number {
  let max = -1;
  const see = (id: string): void => {
    const n = Number(id);
    if (Number.isInteger(n) && n > max) max = n;
  };
  for (const id of Object.keys(body.comments ?? {})) see(id);
  for (const p of walkParagraphs(body)) for (const r of p.runs) for (const id of r.comments ?? []) see(id);
  return max + 1;
}

/** 批注正文：一行一段，没有字的行是空段落。节点 id 由调用方给（要带 `cm:` 前缀，与正文不撞） */
export function commentBlocks(text: string, newId: () => string): Block[] {
  return text.split('\n').map((line): Paragraph => {
    const content: RunContent[] = [];
    line.split('\t').forEach((part, i) => {
      if (i > 0) content.push({ kind: 'tab' });
      if (part !== '') content.push({ kind: 'text', text: part });
    });
    const runs: Run[] = content.length ? [{ kind: 'run', id: newId(), props: {}, content }] : [];
    return { kind: 'paragraph', id: newId(), props: {}, runs };
  });
}

export function newComment(
  id: string,
  draft: CommentDraft,
  blocks: Block[],
  parent?: string,
): CommentContent {
  const out: CommentContent = { id, done: false, blocks };
  if (draft.author !== undefined) out.author = draft.author;
  if (draft.initials !== undefined) out.initials = draft.initials;
  if (draft.date !== undefined) out.date = draft.date;
  if (parent !== undefined) out.parent = parent;
  return out;
}

export function checkCommentDraft(draft: CommentDraft): void {
  if (draft === null || typeof draft !== 'object') throw new TypeError('批注需要 { text } 对象');
  if (typeof draft.text !== 'string') throw new TypeError('批注正文必须是字符串');
  for (const key of ['author', 'initials', 'date'] as const) {
    const v = draft[key];
    if (v !== undefined && typeof v !== 'string') throw new TypeError(`批注的 ${key} 必须是字符串`);
  }
}

/** 线程的顶层那条。指不到的父（文件里写坏了）当顶层 */
export function threadRoot(comments: Readonly<Record<string, CommentContent>>, id: string): string {
  const parent = comments[id]?.parent;
  return parent !== undefined && comments[parent] !== undefined ? parent : id;
}

/** 一条线程：顶层 + 直接挂在它下面的回复，顺序同文件 */
export function threadOf(comments: Readonly<Record<string, CommentContent>>, id: string): string[] {
  const root = threadRoot(comments, id);
  return Object.values(comments)
    .filter((c) => c.id === root || c.parent === root)
    .map((c) => c.id);
}
