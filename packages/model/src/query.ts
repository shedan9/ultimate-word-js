/**
 * 按结构找（api.md §7 的 `doc.query()`）—— 类 CSS 选择器 → 节点列表。
 *
 * 够用即止，不是完整 CSS：类型（`paragraph` / `run` / `table` / `row` / `cell`）、
 * 属性（`[styleId=Heading1]`、`[key]`）、位置伪类（`:first-child` / `:last-child` /
 * `:nth-child(n)`）、后代（空格）与直接子（`>`）。没有逗号分组、没有 `*`、
 * 没有 `~` / `+` —— 需要的时候再加，先摆一堆语法只会让「哪些能用」说不清。
 *
 * 树的父子关系（这决定了 `>` 与 `:first-child` 说的是谁）：
 * `table > row > cell > (paragraph | table)`，`paragraph > run`；
 * 节（section）**不是节点**，顶层的块直接当根。`:nth-child` 数的是父列表里的位置、
 * **不分类型**（与 CSS 一致）：一格里第一个块是表格时，它后面的段落不是 `:first-child`。
 *
 * api.md 里列的 `image` / `field` / `sdt` 三种类型**这里没有**，各有各的原因：
 * 图片是 run 内容的一个片段（`ObjectContent`），不是节点、没有 id，查出来也没法当锚点；
 * 域不在树上（`LoadedDocument.fields` 单独一份，因为它跨段落）；
 * 内容控件（`w:sdt`）解析时当成透明容器剥掉了（parse-body.ts），`tag` / `alias` 根本没留下。
 * 三种一律抛错而不是静默答空 —— 「查不到」与「查不了」是两回事。
 *
 * 属性值按**字符串**比（`String(props[key]) === value`），`[bold=true]` 因此也能写。
 * 读的是节点 `props` 上的字段名，两棵树（直接格式 / 级联完）都行；
 * 但要注意直接格式树上样式里的值看不见 —— `[bold=true]` 在那棵树上只命中直接加粗的。
 */
import type { BlockNode, DocumentBody, PropSet, RunNode, TableCellNode, TableRowNode } from './nodes.ts';

export type QueryNode<S extends PropSet> = BlockNode<S> | RunNode<S> | TableRowNode<S> | TableCellNode<S>;

const TYPES = new Set(['paragraph', 'run', 'table', 'row', 'cell']);
const UNSUPPORTED: Record<string, string> = {
  image: '图片是 run 内容的片段，不是节点',
  field: '域不在节点树上（见 LoadedDocument.fields）',
  sdt: '内容控件解析时已剥成透明容器',
};

interface Attr {
  key: string;
  value?: string;
}
type Pseudo = { kind: 'first' } | { kind: 'last' } | { kind: 'nth'; n: number };
interface Compound {
  type?: string;
  attrs: Attr[];
  pseudos: Pseudo[];
  /** 与**左边**那个 compound 的关系；最左边的没有 */
  combinator?: 'child' | 'descendant';
}

/** 遍历时的一个节点：它自己、在父列表里的下标、父列表的长度 */
interface Entry<S extends PropSet> {
  node: QueryNode<S>;
  index: number;
  count: number;
}

export function queryNodes<S extends PropSet>(body: DocumentBody<S>, selector: string): QueryNode<S>[] {
  const parts = parseSelector(selector);
  const out: QueryNode<S>[] = [];
  const chain: Entry<S>[] = [];
  const visitAll = (children: readonly QueryNode<S>[]) => {
    for (let i = 0; i < children.length; i++) visit(children[i] as QueryNode<S>, i, children.length);
  };
  const visit = (node: QueryNode<S>, index: number, count: number) => {
    chain.push({ node, index, count });
    if (matchesChain(parts, parts.length - 1, chain, chain.length - 1)) out.push(node);
    switch (node.kind) {
      case 'paragraph':
        visitAll(node.runs);
        break;
      case 'table':
        visitAll(node.rows);
        break;
      case 'row':
        visitAll(node.cells);
        break;
      case 'cell':
        visitAll(node.blocks);
        break;
      default:
        break;
    }
    chain.pop();
  };
  for (const section of body.sections) visitAll(section.blocks);
  return out;
}

/**
 * 从右往左配：`parts[pi]` 要配上 `chain[ci]`，再按 combinator 往祖先走。
 * 后代关系要回溯（`A B` 里的 A 可以是任何一个祖先），直接子只看紧邻的那个父节点。
 */
function matchesChain<S extends PropSet>(
  parts: Compound[],
  pi: number,
  chain: Entry<S>[],
  ci: number,
): boolean {
  if (ci < 0) return false;
  const part = parts[pi] as Compound;
  if (!matchesCompound(part, chain[ci] as Entry<S>)) return false;
  if (pi === 0) return true;
  const left = parts[pi - 1] as Compound;
  if (part.combinator === 'child') return matchesChain(parts, pi - 1, chain, ci - 1);
  for (let k = ci - 1; k >= 0; k--) {
    if (matchesCompound(left, chain[k] as Entry<S>) && matchesChain(parts, pi - 1, chain, k)) return true;
  }
  return false;
}

function matchesCompound<S extends PropSet>(part: Compound, entry: Entry<S>): boolean {
  if (part.type !== undefined && entry.node.kind !== part.type) return false;
  const props = entry.node.props as Record<string, unknown> | undefined;
  for (const attr of part.attrs) {
    const actual = props?.[attr.key];
    if (attr.value === undefined) {
      if (actual === undefined || actual === false) return false;
    } else if (actual === undefined || String(actual) !== attr.value) return false;
  }
  for (const pseudo of part.pseudos) {
    if (pseudo.kind === 'first' && entry.index !== 0) return false;
    if (pseudo.kind === 'last' && entry.index !== entry.count - 1) return false;
    if (pseudo.kind === 'nth' && entry.index !== pseudo.n - 1) return false;
  }
  return true;
}

// ── 解析 ──────────────────────────────────────────────────────────────────────

const IDENT = /^[A-Za-z_][\w-]*/;

export function parseSelector(selector: string): Compound[] {
  const src = selector.trim();
  if (src.length === 0) throw new RangeError('选择器不能为空');
  const parts: Compound[] = [];
  let pos = 0;
  let pending: Compound['combinator'];
  while (pos < src.length) {
    const ws = /^\s+/.exec(src.slice(pos));
    if (ws !== null) {
      pos += ws[0].length;
      if (pending === undefined && parts.length > 0) pending = 'descendant';
      continue;
    }
    if (src[pos] === '>') {
      if (parts.length === 0) throw new RangeError(`选择器「${selector}」不能以 > 开头`);
      pending = 'child';
      pos++;
      continue;
    }
    const { compound, next } = parseCompound(src, pos, selector);
    if (parts.length > 0) {
      if (pending === undefined) throw new RangeError(`选择器「${selector}」第 ${pos} 位缺少组合符`);
      compound.combinator = pending;
    }
    parts.push(compound);
    pending = undefined;
    pos = next;
  }
  if (pending === 'child') throw new RangeError(`选择器「${selector}」以 > 结尾`);
  return parts;
}

function parseCompound(src: string, start: number, whole: string): { compound: Compound; next: number } {
  let pos = start;
  const compound: Compound = { attrs: [], pseudos: [] };
  const type = IDENT.exec(src.slice(pos));
  if (type !== null) {
    const name = type[0];
    if (!TYPES.has(name)) {
      const why = UNSUPPORTED[name];
      throw new RangeError(
        why === undefined
          ? `选择器「${whole}」里不认识的类型「${name}」`
          : `选择器类型「${name}」不支持：${why}`,
      );
    }
    compound.type = name;
    pos += name.length;
  }
  let matched = type !== null;
  for (;;) {
    if (src[pos] === '[') {
      const end = src.indexOf(']', pos);
      if (end < 0) throw new RangeError(`选择器「${whole}」的 [ 没有闭合`);
      const inner = src.slice(pos + 1, end).trim();
      const eq = inner.indexOf('=');
      const key = (eq < 0 ? inner : inner.slice(0, eq)).trim();
      if (!IDENT.test(key) || IDENT.exec(key)?.[0] !== key)
        throw new RangeError(`选择器「${whole}」的属性名「${key}」不合法`);
      const attr: Attr = { key };
      if (eq >= 0) attr.value = unquote(inner.slice(eq + 1).trim());
      compound.attrs.push(attr);
      pos = end + 1;
      matched = true;
      continue;
    }
    if (src[pos] === ':') {
      const m = /^:(first-child|last-child|nth-child\((\d+)\))/.exec(src.slice(pos));
      if (m === null) throw new RangeError(`选择器「${whole}」第 ${pos} 位的伪类不认识`);
      if (m[1] === 'first-child') compound.pseudos.push({ kind: 'first' });
      else if (m[1] === 'last-child') compound.pseudos.push({ kind: 'last' });
      else {
        const n = Number(m[2]);
        if (n < 1) throw new RangeError(`选择器「${whole}」的 nth-child 从 1 起数`);
        compound.pseudos.push({ kind: 'nth', n });
      }
      pos += m[0].length;
      matched = true;
      continue;
    }
    break;
  }
  if (!matched) throw new RangeError(`选择器「${whole}」第 ${pos} 位无法解析`);
  return { compound, next: pos };
}

function unquote(value: string): string {
  const q = value[0];
  if ((q === '"' || q === "'") && value.length >= 2 && value.endsWith(q)) return value.slice(1, -1);
  return value;
}
