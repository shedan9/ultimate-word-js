/**
 * 编辑后的批注（`Body.comments`）→ `comments.xml` + `commentsExtended.xml`，**以原文为底打补丁**。
 *
 * 与正文同一个思路（body-writer.ts）：没动过的 `w:comment` 原样吐回（它带着 `w:rsid`、我们不认识的格式），
 * 删掉的摘掉，正文改了的重写段落，新的追加在末尾。三处容易写错：
 *
 * ① **重写正文时留住末段的 `w14:paraId`**：`commentsExtended.xml`（回复挂在谁下面、是否已解决）与
 *    `commentsIds.xml` 都按它认批注，换了等于把这条批注的回复与解决状态弄丢；
 * ② **首段开头的 `w:annotationRef`** 是 Word 在批注框里画「[张三1]」那个号的地方，重写时照原样放回去，
 *    新批注照文件里已有的批注抄一个（连同它的字符样式）。段落格式、文字格式同样照抄 —— 中文版 Word 的
 *    「批注文字」样式 id 是本地化的（`a3` 之类），自己写 `CommentText` 指不到任何样式；
 * ③ **一旦有回复或已解决就要写扩展部件**，扩展部件要求**每条**批注的末段都有 `w14:paraId` ——
 *    WPS 存的文件没有，这时给它们补上（全文唯一，`< 0x80000000`，见 `ParaIds`）。
 *
 * 删批注时顺手清掉 `commentsIds.xml` / `commentsExtensible.xml` 里指空的条目（Word 2019 / 365 写的，
 * 按 paraId → durableId 串起来）；新批注不往里写 —— 两个部件都是可选的，Word 缺了会自己补。
 */
import type { CommentContent, RunContent } from '@uw/model';
import type { XmlDocument, XmlElement, XmlNode } from '@uw/ooxml';
import { attr, child, children } from '@uw/ooxml';
import { W_NS } from './numbering-writer.ts';
import { el, same } from './xml-edit.ts';

const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const DECLARATION = { version: '1.0', encoding: 'UTF-8', standalone: 'yes' } as const;

/**
 * 全文唯一的段落 id。Word 要求 `w14:paraId` 在整份文档（正文、页眉页脚、注、批注）里不重复、
 * 且小于 0x80000000。从一个固定的数往上数、跳过用过的 —— 不用随机数，同一次编辑导出两遍逐字节相同
 */
export class ParaIds {
  readonly #used: Set<string>;
  #next = 0x40000000;
  constructor(used: Iterable<string>) {
    this.#used = new Set([...used].map((id) => id.toUpperCase()));
  }
  fresh(): string {
    let id: string;
    do id = (this.#next++).toString(16).toUpperCase().padStart(8, '0');
    while (this.#used.has(id));
    this.#used.add(id);
    return id;
  }
}

export interface CommentsWriteInput {
  /** 原来的 `comments.xml`；文档原来没有批注时缺席 */
  base: XmlDocument | undefined;
  /** 原来的 `commentsExtended.xml` */
  baseExtended: XmlDocument | undefined;
  /** 加载时的批注（同一个包重新解析的） */
  original: Readonly<Record<string, CommentContent>>;
  /** 编辑后的批注（`body.comments ?? {}`） */
  now: Readonly<Record<string, CommentContent>>;
  paraIds: ParaIds;
}

export interface CommentsWriteResult {
  /** 要写的 `comments.xml`；没变时缺席 */
  comments?: XmlDocument;
  /** 要写的 `commentsExtended.xml`；不需要或没变时缺席 */
  extended?: XmlDocument;
  /** 删掉的批注末段的 paraId，给 `pruneCommentIds` 清扩展部件 */
  removedParaIds: string[];
}

export function writeComments(input: CommentsWriteInput): CommentsWriteResult {
  const { base, baseExtended, original, now, paraIds } = input;
  if (same(original, now)) return { removedParaIds: [] };

  const elements = new Map<string, XmlElement>();
  for (const c of base === undefined ? [] : children(base.root, 'w:comment')) {
    const id = attr(c, 'w:id');
    if (id !== undefined && !elements.has(id)) elements.set(id, c);
  }
  // 新批注抄格式的模板：文件里第一条批注
  const template = elements.values().next().value;
  const needExtended =
    baseExtended !== undefined || Object.values(now).some((c) => c.parent !== undefined || c.done);

  const paraIdOf = new Map<string, string>();
  const removedParaIds: string[] = [];
  const write = (c: CommentContent): XmlElement => {
    const src = elements.get(c.id);
    const had = src === undefined ? undefined : lastParaId(src);
    const paraId = had ?? (needExtended ? paraIds.fresh() : undefined);
    if (paraId !== undefined) paraIdOf.set(c.id, paraId);
    const was = original[c.id];
    if (src !== undefined && was !== undefined && same(was, c) && paraId === had) return src;
    const attrs: Record<string, string> = { ...(src?.attrs ?? { 'w:id': c.id }) };
    // `w:author` 是必填属性（`CT_TrackChange`），没给作者也得写一个空的
    attrs['w:author'] = c.author ?? attrs['w:author'] ?? '';
    if (c.date !== undefined) attrs['w:date'] = c.date;
    if (c.initials !== undefined) attrs['w:initials'] = c.initials;
    const blocks =
      src !== undefined && was !== undefined && same(was.blocks, c.blocks)
        ? withLastParaId(src.children, paraId)
        : commentParagraphs(c, src ?? template, paraId);
    return el('w:comment', attrs, blocks);
  };

  const out: XmlNode[] = [];
  for (const node of base?.root.children ?? []) {
    if (node.kind !== 'element' || node.name !== 'w:comment') {
      out.push(node);
      continue;
    }
    const id = attr(node, 'w:id');
    const c = id === undefined ? undefined : now[id];
    if (c !== undefined) out.push(write(c));
    else {
      const gone = lastParaId(node);
      if (gone !== undefined) removedParaIds.push(gone);
    }
  }
  for (const c of Object.values(now)) if (!elements.has(c.id)) out.push(write(c));

  const root = withNamespaces(
    base?.root ?? el('w:comments', { 'xmlns:w': W_NS }),
    paraIdOf.size ? { w14: W14_NS } : {},
  );
  const result: CommentsWriteResult = {
    comments: {
      declaration: base === undefined ? DECLARATION : base.declaration,
      root: { ...root, children: out },
    },
    removedParaIds,
  };
  if (needExtended) {
    const extended = writeExtended(
      baseExtended,
      now,
      paraIdOf,
      new Set(removedParaIds.map((p) => p.toUpperCase())),
      new Set(Object.keys(original)),
    );
    if (extended !== undefined) result.extended = extended;
  }
  return result;
}

function lastParaId(comment: XmlElement): string | undefined {
  const last = children(comment, 'w:p').at(-1);
  return last === undefined ? undefined : attr(last, 'w14:paraId');
}

/** 正文没变、只差末段补一个 paraId（扩展部件要用）时，只改那一段的属性 */
function withLastParaId(nodes: readonly XmlNode[], paraId: string | undefined): XmlNode[] {
  if (paraId === undefined) return [...nodes];
  const at = nodes.findLastIndex((n) => n.kind === 'element' && n.name === 'w:p');
  const last = nodes[at];
  if (last === undefined || last.kind !== 'element' || last.attrs['w14:paraId'] === paraId) return [...nodes];
  return nodes.map((n, i) => (i === at ? { ...last, attrs: { ...last.attrs, 'w14:paraId': paraId } } : n));
}

/**
 * 按模型重写批注的段落。格式照模板（这条批注自己原来的样子，新批注照文件里第一条）：
 * 段落属性取首段的 `w:pPr`，`w:annotationRef` 那个 run 整个照抄，文字 run 的 `w:rPr` 取第一个有字的 run 的
 */
function commentParagraphs(
  c: CommentContent,
  template: XmlElement | undefined,
  paraId: string | undefined,
): XmlNode[] {
  const firstP = template === undefined ? undefined : child(template, 'w:p');
  const pPr = firstP === undefined ? undefined : child(firstP, 'w:pPr');
  const runs = template === undefined ? [] : children(template, 'w:p').flatMap((p) => children(p, 'w:r'));
  const refRun =
    runs.find((r) => child(r, 'w:annotationRef') !== undefined) ?? el('w:r', {}, [el('w:annotationRef')]);
  const textRPr = runs
    .find((r) => child(r, 'w:t') !== undefined)
    ?.children.find((n): n is XmlElement => n.kind === 'element' && n.name === 'w:rPr');
  const paragraphs = c.blocks.filter((b) => b.kind === 'paragraph');
  // 一段都没有的批注 Word 认为损坏（`CT_Comment` 至少要一个块）
  const list = paragraphs.length ? paragraphs : [{ kind: 'paragraph' as const, id: '', props: {}, runs: [] }];
  return list.map((p, i) => {
    const kids: XmlNode[] = [];
    if (pPr !== undefined) kids.push(pPr);
    if (i === 0) kids.push(refRun);
    for (const r of p.runs) {
      const items = r.content.flatMap(contentElement);
      if (items.length) kids.push(el('w:r', {}, textRPr === undefined ? items : [textRPr, ...items]));
    }
    const attrs: Record<string, string> =
      i === list.length - 1 && paraId !== undefined ? { 'w14:paraId': paraId } : {};
    return el('w:p', attrs, kids);
  });
}

function contentElement(c: RunContent): XmlElement[] {
  if (c.kind === 'text')
    return c.text === ''
      ? []
      : [
          el('w:t', /^\s|\s$/.test(c.text) ? { 'xml:space': 'preserve' } : {}, [
            { kind: 'text', text: c.text },
          ]),
        ];
  if (c.kind === 'tab') return [el('w:tab')];
  if (c.kind === 'break') return [el('w:br')];
  return [];
}

/** 根元素上补命名空间声明，并把它们登记进 `mc:Ignorable`（老版 Word 遇到不认识又没声明可忽略的前缀会报损坏） */
function withNamespaces(root: XmlElement, extra: Readonly<Record<string, string>>): XmlElement {
  const prefixes = Object.keys(extra).filter((p) => root.attrs[`xmlns:${p}`] === undefined);
  if (prefixes.length === 0) return root;
  const attrs = { ...root.attrs };
  for (const p of prefixes) attrs[`xmlns:${p}`] = extra[p] as string;
  if (attrs['xmlns:mc'] === undefined) attrs['xmlns:mc'] = MC_NS;
  const ignorable = new Set((attrs['mc:Ignorable'] ?? '').split(/\s+/).filter(Boolean));
  for (const p of prefixes) ignorable.add(p);
  attrs['mc:Ignorable'] = [...ignorable].join(' ');
  return { ...root, attrs };
}

/**
 * `commentsExtended.xml`：每条批注一个 `w15:commentEx`（paraId + 父的 paraId + 是否已解决）。
 * 原有的条目按原顺序、没变就原样留着；指不到任何批注的条目也留着（不是我们的，不替人删）
 */
function writeExtended(
  base: XmlDocument | undefined,
  now: Readonly<Record<string, CommentContent>>,
  paraIdOf: ReadonlyMap<string, string>,
  removed: ReadonlySet<string>,
  existed: ReadonlySet<string>,
): XmlDocument | undefined {
  const want = new Map<string, Record<string, string>>();
  /** 原有的、条目缺席也没关系的批注（顶层、未解决 —— 缺席就是这个意思）：不为它们补条目，部件才能原样照搬 */
  const optional = new Set<string>();
  for (const c of Object.values(now)) {
    const paraId = paraIdOf.get(c.id);
    if (paraId === undefined) continue;
    if (existed.has(c.id) && c.parent === undefined && !c.done) optional.add(paraId.toUpperCase());
    const attrs: Record<string, string> = { 'w15:paraId': paraId };
    const parent = c.parent === undefined ? undefined : paraIdOf.get(c.parent);
    if (parent !== undefined) attrs['w15:paraIdParent'] = parent;
    attrs['w15:done'] = c.done ? '1' : '0';
    want.set(paraId.toUpperCase(), attrs);
  }
  let changed = base === undefined;
  const out: XmlNode[] = [];
  const seen = new Set<string>();
  for (const node of base?.root.children ?? []) {
    if (node.kind !== 'element' || node.name !== 'w15:commentEx') {
      out.push(node);
      continue;
    }
    const key = (attr(node, 'w15:paraId') ?? '').toUpperCase();
    const attrs = want.get(key);
    if (attrs === undefined) {
      // 指着的批注被删了 —— 摘掉；指着别的段落（不是我们认得的批注）—— 留着
      if (removed.has(key)) changed = true;
      else out.push(node);
      continue;
    }
    seen.add(key);
    const done = attr(node, 'w15:done') === '1' || attr(node, 'w15:done') === 'true';
    if (
      done === (attrs['w15:done'] === '1') &&
      attr(node, 'w15:paraIdParent') === attrs['w15:paraIdParent']
    ) {
      out.push(node);
      continue;
    }
    changed = true;
    const next = { ...node.attrs };
    delete next['w15:paraIdParent'];
    out.push({ ...node, attrs: { ...next, ...attrs } });
  }
  for (const [key, attrs] of want) {
    if (seen.has(key) || (base !== undefined && optional.has(key))) continue;
    changed = true;
    out.push(el('w15:commentEx', attrs));
  }
  if (!changed) return undefined;
  const root =
    base?.root ?? el('w15:commentsEx', { 'xmlns:mc': MC_NS, 'xmlns:w15': W15_NS, 'mc:Ignorable': 'w15' });
  return {
    declaration: base === undefined ? DECLARATION : base.declaration,
    root: { ...root, children: out },
  };
}

/**
 * 删批注后清 `commentsIds.xml`（paraId → durableId）与 `commentsExtensible.xml`（durableId → …）里指空的条目。
 * 返回改过的两份（没有要清的就缺席）
 */
export function pruneCommentIds(
  ids: XmlDocument | undefined,
  extensible: XmlDocument | undefined,
  removedParaIds: readonly string[],
): { ids?: XmlDocument; extensible?: XmlDocument } {
  if (ids === undefined || removedParaIds.length === 0) return {};
  const removed = new Set(removedParaIds.map((p) => p.toUpperCase()));
  const durable = new Set<string>();
  const kept = ids.root.children.filter((n) => {
    if (n.kind !== 'element' || n.name !== 'w16cid:commentId') return true;
    if (!removed.has((attr(n, 'w16cid:paraId') ?? '').toUpperCase())) return true;
    const d = attr(n, 'w16cid:durableId');
    if (d !== undefined) durable.add(d);
    return false;
  });
  if (kept.length === ids.root.children.length) return {};
  const out: { ids?: XmlDocument; extensible?: XmlDocument } = {
    ids: { ...ids, root: { ...ids.root, children: kept } },
  };
  if (extensible !== undefined) {
    const rest = extensible.root.children.filter(
      (n) =>
        n.kind !== 'element' ||
        n.name !== 'w16cex:commentExtensible' ||
        !durable.has(attr(n, 'w16cex:durableId') ?? ''),
    );
    if (rest.length !== extensible.root.children.length)
      out.extensible = { ...extensible, root: { ...extensible.root, children: rest } };
  }
  return out;
}
