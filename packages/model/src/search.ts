/**
 * 按文本找（api.md §7 的 `doc.find()`）—— 一串字 / 一个正则 → 命中的 `DocRange[]`。
 *
 * 三处容易搞反：
 * 1. **匹配跨 run、不跨段落**。Word 会把一个词切成好几个 run（拼写检查、字体切换、
 *    修订都会切），「签发人」三个字分在三个 run 里是常态，按 run 搜永远搜不到；
 *    而段落是 Word 查找的天然边界（`^p` 才能跨），一段的末尾与下一段的开头
 *    拼在一起匹配上是错的。所以是**逐段拼成一串再搜**，每个 UTF-16 单元记着它来自哪
 * 2. 拼出来的串与排版**看见的**一致，不是文件里**有的**：隐藏 run（`w:vanish`）不进串、
 *    域的界桩与指令不进串（它们不显示）、软连字符不进串（平时宽 0）；
 *    制表位与换行各算一个字符（`\t` / `\n`），对象算一个 U+FFFC —— 它**阻断**匹配，
 *    「签发」与「人」中间夹着一张图不该匹配「签发人」。
 *    求值过的域（`fieldValues` 里有的 run）整个跳过：它们显示的是算出来的页码，
 *    文件里存的旧值搜到了也指不到屏幕上任何一个字（布局层把它们的位置抹成 -1）
 * 3. 大小写：字符串默认**不分**（Word 的默认），`matchCase` 才分；正则跟自己的 flags 走。
 *    字符串一律转成正则跑同一条路 —— 用 `toLowerCase()` 比会改变 UTF-16 长度
 *    （`İ`.toLowerCase() 是两个单元），位置就对不上了
 *
 * 返回的 range 是半开区间：`end` 指最后一个命中单元**之后**，落在片段末尾时
 * `offset` 等于片段长度（`DocPosition` 允许这个值）。零长匹配（`/a*\/`）一律跳过，
 * 它们既画不出矩形也没法定位。
 */
import type { NodeId, ResolvedBody, ResolvedParagraph, ResolvedRun } from './nodes.ts';
import { walkParagraphs } from './nodes.ts';
import type { DocPosition, DocRange } from './position.ts';

export interface FindOptions {
  /** 最多返回几条，默认不限。整份长文档搜一个常见字，不限的话调用方自己都不想要 */
  limit?: number;
  /** 字符串模式是否区分大小写，默认不分。对正则无效（跟它自己的 `i` 走） */
  matchCase?: boolean;
  /** 求值过的域（`layoutDocumentWithFields()` 产出的 `fieldValues`），这些 run 整个跳过 */
  fieldValues?: ReadonlyMap<NodeId, string>;
}

/** 对象在拼出来的串里的占位符。它不会被任何正常的查找串匹配到，只用来阻断跨对象的匹配 */
const OBJECT_MARK = '￼';

/**
 * 吃的是**级联完**的那棵树（`LoadedDocument.resolved`），不是直接格式那棵：
 * 「隐藏不隐藏」要级联完才知道 —— `w:vanish` 常写在字符样式里，直接格式树上看不见。
 * 节点 id 两棵树一样，所以结果照样能指回可编辑的那棵。
 */
export function findText(
  body: ResolvedBody,
  pattern: string | RegExp,
  options: FindOptions = {},
): DocRange[] {
  const limit = options.limit ?? Number.POSITIVE_INFINITY;
  if (!(Number.isInteger(limit) || limit === Number.POSITIVE_INFINITY) || limit < 0) {
    throw new RangeError('limit 必须是非负整数');
  }
  const re = toRegExp(pattern, options.matchCase === true);
  const out: DocRange[] = [];
  if (limit === 0) return out;
  for (const p of walkParagraphs(body)) {
    const { text, positions } = flattenParagraph(p, options.fieldValues);
    if (text.length === 0) continue;
    re.lastIndex = 0;
    for (;;) {
      const m = re.exec(text);
      if (m === null) break;
      if (m[0].length === 0) {
        // 零长匹配不推进 lastIndex，不手动跳一格会原地打转
        re.lastIndex++;
        continue;
      }
      const first = positions[m.index] as DocPosition;
      const last = positions[m.index + m[0].length - 1] as DocPosition;
      out.push({ start: first, end: { ...last, offset: last.offset + 1 } });
      if (out.length >= limit) return out;
    }
  }
  return out;
}

/**
 * 字符串 → 转义后的正则；正则 → 带 `g` 的副本。
 * 不直接用调用方的正则对象：它的 `lastIndex` 是共享的可变状态，两次调用会互相踩。
 */
function toRegExp(pattern: string | RegExp, matchCase: boolean): RegExp {
  if (typeof pattern === 'string') {
    if (pattern.length === 0) throw new RangeError('查找串不能为空');
    return new RegExp(pattern.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), matchCase ? 'g' : 'gi');
  }
  // `y`（sticky）会让 exec 只认 lastIndex 那一点，与「找全部」的语义冲突，去掉
  const flags = pattern.flags.replace('y', '');
  return new RegExp(pattern.source, flags.includes('g') ? flags : `${flags}g`);
}

/** 一段的可见文字 + 每个 UTF-16 单元的模型位置，两者等长 */
function flattenParagraph(
  p: ResolvedParagraph,
  fieldValues: ReadonlyMap<NodeId, string> | undefined,
): { text: string; positions: DocPosition[] } {
  let text = '';
  const positions: DocPosition[] = [];
  const push = (run: ResolvedRun, ci: number, s: string) => {
    for (let i = 0; i < s.length; i++) positions.push({ nodeId: run.id, contentIndex: ci, offset: i });
    text += s;
  };
  for (const run of p.runs) {
    if (run.props.hidden) continue;
    if (fieldValues?.has(run.id)) continue;
    for (let ci = 0; ci < run.content.length; ci++) {
      const c = run.content[ci] as ResolvedRun['content'][number];
      switch (c.kind) {
        case 'text':
          push(run, ci, c.text);
          break;
        case 'tab':
          push(run, ci, '\t');
          break;
        case 'break':
          push(run, ci, '\n');
          break;
        case 'symbol':
          // 只取第一个码点：布局层也只画第一个（items.ts 的 symbol 分支）
          push(run, ci, String.fromCodePoint(c.char.codePointAt(0) ?? 0xfffd));
          break;
        case 'noBreakHyphen':
          push(run, ci, '-');
          break;
        case 'object':
          push(run, ci, OBJECT_MARK);
          break;
        // 软连字符平时不显示；域界桩与指令从不显示
        default:
          break;
      }
    }
  }
  return { text, positions };
}
