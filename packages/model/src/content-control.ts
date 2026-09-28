/**
 * 内容控件（`w:sdt`）—— 模板填充（api.md §9 `doc.bindings`）的底子。
 *
 * 排版上它仍是透明容器（parse-body.ts 照旧把内容压平），这里只多留两样东西：
 * ① `sdtPr` 里的**身份**（`w:tag` / `w:alias` / 类型 / 锁 / 占位符标记），挂在 `Body.contentControls`；
 * ② **谁在控件里** —— 行内控件标在 run 上、块级控件标在段落 / 表格上（nodes.ts 的 `contentControl`）。
 *
 * 为什么「谁在控件里」不存成「控件 → 起止位置」：控件里的内容会被编辑（拆 run、拆段、删字），
 * 起止位置每次都得映射，而标记跟着节点走 —— 拆出来的 run / 段落是 `{ ...原节点 }`，自动带着。
 *
 * 嵌套靠 `parent` 串：run 只记最内层那个，判断「在不在控件 X 里」沿 parent 往外找。
 * 表格行 / 单元格那一层的 `w:sdt`（包着 `w:tr` / `w:tc`）不收 —— 它们是重复区块的外壳，
 * 填值填不进一整行，照旧当透明容器。
 */
import type { XmlElement } from '@uw/ooxml';
import { attr, child, children } from '@uw/ooxml';
import type { DocumentBody, NodeId, ParagraphNode, PropSet, RunNode } from './nodes.ts';
import { walkBlocks } from './nodes.ts';
import { runEnd, runStart } from './order.ts';
import { parseRunProps } from './parse-props.ts';
import type { DocPosition, DocRange } from './position.ts';
import type { RunProps } from './props.ts';
import { attrOf } from './xml-values.ts';

/**
 * 控件类型，取自 `sdtPr` 里的类型元素。没有类型元素的是**富文本**（Word 的默认）。
 * `gallery` 是 `w:docPartObj` / `w:docPartList`（目录、封面这类文档部件的外壳，不是给人填的），
 * 其余不认识的（公式、引文、`w15:repeatingSection`…）一律 `other`。
 */
export type ContentControlType =
  | 'richText'
  | 'text'
  | 'date'
  | 'dropDownList'
  | 'comboBox'
  | 'checkbox'
  | 'picture'
  | 'gallery'
  | 'other';

/** `w:lock`：`sdtLocked` 不许删控件、`contentLocked` 不许改内容、`sdtContentLocked` 两样都不许 */
export type ContentControlLock = 'sdtLocked' | 'contentLocked' | 'sdtContentLocked';

export interface ContentControl {
  id: NodeId;
  /** `w:tag`：给程序看的名字，模板填充按它认坑位 */
  tag?: string;
  /** `w:alias`：Word 界面上显示的标题 */
  alias?: string;
  type: ContentControlType;
  /** 包着 run 的是 `inline`，包着段落 / 表格的是 `block` */
  scope: 'inline' | 'block';
  /** 外层控件；最外层缺席 */
  parent?: NodeId;
  /**
   * `w:showingPlcHdr`：里面现在是占位文字（「单击或点击此处输入文字。」）。
   * 填了值要清掉 —— 留着的话 Word 把填进去的字当占位符，点一下整段选中、一打字就没了
   */
  showingPlaceholder: boolean;
  lock?: ContentControlLock;
  /**
   * `sdtPr/w:rPr`：控件自己的字符格式。占位文字用的是 `PlaceholderText` 字符样式（灰字），
   * 填值时换成这一份 —— 与 Word 里点进占位符开始打字的结果一致
   */
  runProps?: RunProps;
  /** 下拉 / 组合框的选项。`value` 缺席时等于显示文字 */
  items?: { text: string; value: string }[];
  /** `w:text w:multiLine`：纯文本控件允不允许换行 */
  multiLine?: boolean;
  /**
   * 有没有 `w:dataBinding`（绑到 customXml 部件的某个 XPath）。
   * 绑定的控件在 Word 打开时**会被 customXml 里的值盖掉** —— 只改 `sdtContent` 等于白填，
   * 所以回写时内容一变就去掉绑定（serialize 的 body-writer.ts）
   */
  dataBound?: boolean;
}

const TYPE_ELEMENTS: readonly [string, ContentControlType][] = [
  ['w:text', 'text'],
  ['w:richText', 'richText'],
  ['w:date', 'date'],
  ['w:dropDownList', 'dropDownList'],
  ['w:comboBox', 'comboBox'],
  ['w14:checkbox', 'checkbox'],
  ['w:picture', 'picture'],
  ['w:docPartObj', 'gallery'],
  ['w:docPartList', 'gallery'],
  ['w:equation', 'other'],
  ['w:citation', 'other'],
  ['w:bibliography', 'other'],
  ['w:group', 'other'],
  ['w15:repeatingSection', 'other'],
  ['w15:repeatingSectionItem', 'other'],
];

const LOCKS: readonly ContentControlLock[] = ['sdtLocked', 'contentLocked', 'sdtContentLocked'];

/** `w:sdt` → 控件属性。`id` 由调用方给（按解析顺序生成，回写靠它对回原元素） */
export function parseContentControl(
  sdt: XmlElement,
  id: NodeId,
  scope: ContentControl['scope'],
  parent: NodeId | undefined,
): ContentControl {
  const pr = child(sdt, 'w:sdtPr');
  let type: ContentControlType = 'richText';
  const out: ContentControl = { id, type, scope, showingPlaceholder: false };
  if (parent !== undefined) out.parent = parent;
  if (pr === undefined) return out;
  for (const [name, t] of TYPE_ELEMENTS) {
    const e = child(pr, name);
    if (e === undefined) continue;
    type = t;
    if (t === 'text') {
      const multi = attr(e, 'w:multiLine');
      if (multi !== undefined && multi !== '0' && multi !== 'false') out.multiLine = true;
    } else if (t === 'dropDownList' || t === 'comboBox') {
      out.items = children(e, 'w:listItem').map((li) => {
        const text = attr(li, 'w:displayText') ?? attr(li, 'w:value') ?? '';
        return { text, value: attr(li, 'w:value') ?? text };
      });
    }
    break;
  }
  out.type = type;
  const tag = attrOf(child(pr, 'w:tag'), 'w:val');
  const alias = attrOf(child(pr, 'w:alias'), 'w:val');
  if (tag !== undefined) out.tag = tag;
  if (alias !== undefined) out.alias = alias;
  const plc = child(pr, 'w:showingPlcHdr');
  // 开关属性：`<w:showingPlcHdr/>` 就是 true，只有显式写 0 / false 才是关
  if (plc !== undefined && !['0', 'false', 'off'].includes(attr(plc, 'w:val') ?? ''))
    out.showingPlaceholder = true;
  const lock = LOCKS.find((l) => l === attrOf(child(pr, 'w:lock'), 'w:val'));
  if (lock !== undefined) out.lock = lock;
  const rPr = child(pr, 'w:rPr');
  if (rPr !== undefined) out.runProps = parseRunProps(rPr);
  if (child(pr, 'w:dataBinding') !== undefined) out.dataBound = true;
  return out;
}

/** `inner` 是不是 `target` 本身或在它里面（沿 parent 往外找） */
export function withinContentControl(
  controls: Readonly<Record<NodeId, ContentControl>>,
  inner: NodeId | undefined,
  target: NodeId,
): boolean {
  for (let id = inner, guard = 0; id !== undefined && guard < 64; id = controls[id]?.parent, guard++) {
    if (id === target) return true;
  }
  return false;
}

/** 控件在树上占的那一段：成员 run（文档序）与成员段落 */
export interface ContentControlSpan<S extends PropSet> {
  control: ContentControl;
  runs: RunNode<S>[];
  paragraphs: ParagraphNode<S>[];
}

/**
 * 按文档序列出每个控件罩着的 run 与段落。一个 run 属于它的最内层控件（行内标记优先，
 * 没有就看段落的块级标记）以及这个控件的全部外层。
 *
 * **内容全被删光的控件不列**：行内控件里的 run 删光了，回写时这个容器也不留
 * （serialize 的 body-writer.ts），它在文档里已经不存在了。
 */
export function contentControlSpans<S extends PropSet>(
  body: DocumentBody<S> & { contentControls?: Readonly<Record<NodeId, ContentControl>> },
): ContentControlSpan<S>[] {
  const controls = body.contentControls;
  if (controls === undefined) return [];
  const spans = new Map<NodeId, ContentControlSpan<S>>();
  const span = (id: NodeId) => {
    let s = spans.get(id);
    if (s === undefined) {
      const control = controls[id];
      if (control === undefined) return undefined;
      s = { control, runs: [], paragraphs: [] };
      spans.set(id, s);
    }
    return s;
  };
  const chain = (inner: NodeId | undefined): NodeId[] => {
    const out: NodeId[] = [];
    for (let id = inner, guard = 0; id !== undefined && guard < 64; id = controls[id]?.parent, guard++)
      out.push(id);
    return out;
  };
  for (const section of body.sections) {
    for (const block of walkBlocks(section.blocks)) {
      if (block.kind !== 'paragraph') continue;
      for (const id of chain(block.contentControl)) span(id)?.paragraphs.push(block);
      for (const run of block.runs) {
        for (const id of chain(run.contentControl ?? block.contentControl)) {
          const s = span(id);
          if (s === undefined) continue;
          s.runs.push(run);
          // 行内控件没有段落标记，但「它在哪一段」要有 —— 取成员 run 所在的段落
          if (s.control.scope === 'inline' && s.paragraphs.at(-1) !== block) s.paragraphs.push(block);
        }
      }
    }
  }
  return [...spans.values()].filter((s) => s.runs.length > 0 || s.control.scope === 'block');
}

/**
 * 控件内容的范围：首个成员 run 的开头到末个成员 run 的结尾；块级控件里一个 run 都没有时
 * 是首段的折叠位置（空段落的位置就是段落 id 的 {0, 0}）。
 */
export function rangeOfContentControl<S extends PropSet>(span: ContentControlSpan<S>): DocRange | undefined {
  const first = span.runs[0];
  const last = span.runs.at(-1);
  if (first !== undefined && last !== undefined) return { start: runStart(first), end: runEnd(last) };
  const p = span.paragraphs[0];
  if (p === undefined) return undefined;
  const at: DocPosition = { nodeId: p.id, contentIndex: 0, offset: 0 };
  return { start: at, end: { ...at } };
}

/**
 * 控件现在显示的纯文字：段落之间 `\n`，软换行 `\n`，制表位 `\t`。
 * 给 `bindings.list()` 的 `value` 用 —— 不是排版文字（域界桩、对象都不进来）。
 */
export function textOfContentControl<S extends PropSet>(span: ContentControlSpan<S>): string {
  const members = new Set(span.runs);
  return span.paragraphs
    .map((p) => {
      let out = '';
      for (const run of p.runs) {
        if (!members.has(run)) continue;
        for (const c of run.content) {
          if (c.kind === 'text') out += c.text;
          else if (c.kind === 'tab') out += '\t';
          else if (c.kind === 'break') out += '\n';
          else if (c.kind === 'symbol') out += c.char;
          else if (c.kind === 'noBreakHyphen') out += '-';
        }
      }
      return out;
    })
    .join('\n');
}
