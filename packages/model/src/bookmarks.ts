/**
 * 书签 → 它起点所在的段落。
 *
 * 书签本身挂在段落上（`ParagraphNode.bookmarks`），这里只把它摊成一张查表给 PAGEREF 用。
 * 它吃的是**可编辑**的那棵树：级联树不带书签（排版用不上，带上还会进段落缓存的键），
 * 而段落 id 在两棵树上一致，所以布局拿着这张表照样查得到页。
 *
 * 同名书签按文档序**先到先得** —— Word 不许重名，真出现（拼接出来的文档）时它认第一个。
 */
import type { Body, NodeId } from './nodes.ts';
import { walkParagraphs } from './nodes.ts';

export type BookmarkTargets = ReadonlyMap<string, NodeId>;

export function bookmarkTargets(body: Body): BookmarkTargets {
  const out = new Map<string, NodeId>();
  for (const p of walkParagraphs(body)) {
    for (const name of p.bookmarks ?? []) if (!out.has(name)) out.set(name, p.id);
  }
  return out;
}
