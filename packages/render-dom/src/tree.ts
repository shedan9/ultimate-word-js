/**
 * 渲染中间层：**纯数据的元素树**，不是 DOM。
 *
 * 为什么中间要隔这一层，而不是一路 `document.createElementNS` 画到底：
 *
 * 1. **能在 Node 里测。** 画得对不对是「这个片段的 x 是不是 89.53pt」这种数值问题，
 *    与浏览器毫无关系。隔了这一层，`@uw/render-dom` 的单测跑在纯 Node 环境里，
 *    不需要 jsdom / happy-dom，与仓库里其余六个包一致
 * 2. **截图回归的前提。** `serialize()` 出来的 SVG 可以直接落盘、可以用任何工具渲成位图
 *    与 Word 导出的 PDF 比对 —— 这条辅助手段是开发计划 §7 把 DOM 渲染器排在第一位的理由之一
 * 3. **与 Worker 化同构。** 树是可结构化克隆的（原则 1.1），将来「Worker 里排版 +
 *    主线程只负责贴 DOM」时，过界的就是这棵树
 *
 * 代价是多一次遍历。相对 DOM 写入本身的开销可以忽略，真成瓶颈时增量渲染要比的也正是
 * 这棵树的前后两版，那时它反而是必需的。
 */

/**
 * 一个元素。`children` 与 `text` **互斥**：有 `text` 的是叶子（`<text>` / `<tspan>`），
 * 其余一律用 `children`。属性值统一存成字符串，数字格式化在造树时就做完
 * （见 `fmt`）—— 序列化与建 DOM 两条路径都拿到同一份文本，不会一条路 3 位小数、
 * 另一条路 17 位。
 */
export interface RElement {
  tag: string;
  attrs: Record<string, string>;
  children: RElement[];
  text?: string;
}

export function el(tag: string, attrs: Record<string, string>, children: RElement[] = []): RElement {
  return { tag, attrs, children };
}

export function textEl(tag: string, attrs: Record<string, string>, text: string): RElement {
  return { tag, attrs, children: [], text };
}

/**
 * 数字 → 属性字符串，**保留 3 位小数**（0.001pt ≈ 0.00035px，远细于任何显示设备）。
 *
 * 定死位数不是为了好看，是为了让 golden file 稳定：浮点数直接 `String()` 会在不同
 * 平台上吐出 `12.300000000000001` 这种尾巴，快照测试会因此在 CI 上红。
 */
export function fmt(n: number): string {
  const r = Math.round(n * 1000) / 1000;
  // -0 与 0 在快照里是两个字符串，统一掉
  return Object.is(r, -0) ? '0' : String(r);
}

/** 一串数字 → SVG 的 x 列表（`<text x="x1 x2 x3 …">`） */
export function fmtList(ns: readonly number[]): string {
  let out = '';
  for (let i = 0; i < ns.length; i++) {
    if (i > 0) out += ' ';
    out += fmt(ns[i] as number);
  }
  return out;
}

const TEXT_ESCAPE = /[&<>]/g;
const ATTR_ESCAPE = /[&<>"]/g;
const ESCAPES: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' };

function escapeText(s: string): string {
  return s.replace(TEXT_ESCAPE, (c) => ESCAPES[c] as string);
}

function escapeAttr(s: string): string {
  return s.replace(ATTR_ESCAPE, (c) => ESCAPES[c] as string);
}

/**
 * 元素树 → 标记文本。
 *
 * 输出**不换行不缩进**：这东西的读者是浏览器与快照比对，不是人；加了缩进之后
 * `<text>` 里的空白会变成真的空格（SVG 不折叠空白），文字位置直接错。
 */
export function serialize(node: RElement): string {
  let out = `<${node.tag}`;
  for (const [k, v] of Object.entries(node.attrs)) out += ` ${k}="${escapeAttr(v)}"`;
  if (node.text === undefined && node.children.length === 0) return `${out}/>`;
  out += '>';
  if (node.text !== undefined) out += escapeText(node.text);
  for (const c of node.children) out += serialize(c);
  return `${out}</${node.tag}>`;
}
