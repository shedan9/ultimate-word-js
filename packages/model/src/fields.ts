/**
 * 域（field）—— 把一串扁平的界桩还原成「指令 + 结果」。
 *
 * OOXML 里域**不是一个元素**，而是散在 run 序列里的几颗界桩：
 * `fldChar=begin` → 若干 `w:instrText` → `fldChar=separate` → 结果 run → `fldChar=end`。
 * 解析层（parse-body.ts）刻意只保留位置不做配对，因为配对必须在**整份 body** 上做：
 * 一个 TOC 域能横跨几十个段落，按段落分开扫就永远配不上。
 *
 * 这里**不做求值**。PAGE 要页码、TOC 要目录，都得等分页（Phase 3）；而 Word 存盘时
 * 已经把上次算出来的结果写在 separate 与 end 之间，直接显示就是「打开即所见」——
 * 这也是为什么本阶段不做求值也能正确渲染。
 *
 * 域还有个**压缩写法** `w:fldSimple`：整个域一个元素，域代码在属性里、结果就是子 run。
 * 解析层把它压平成 run 上的标记（`RunNode.fieldSimple`），这里一并收成 `kind: 'simple'`
 * 的区 —— 不收的话，凡是第三方生成器写出来的 HYPERLINK 都会平白失效。
 *
 * 三处容易搞反：
 * 1. **没有 separate 的域什么都不显示**。不是「显示指令」，是整段空 ——
 *    Word 里 `{ PAGE }` 从未刷新过就是这个样子
 * 2. **嵌套域的 instrText 归内层**（靠栈分家）。于是外层的指令文字里**缺一块** ——
 *    `IF { PAGE } = 1` 的外层只看得到 `IF  = 1`。补那一块是求值期的事（要先算内层），
 *    在这里替换会把「结构还原」和「求值」搅在一起
 * 3. 结果区按 **run 粒度**给（`resultRuns`）。界桩在真实文件里总是独占一个 run，
 *    所以这个近似不会咬人；真遇到「一个 run 里既有界桩又有结果文字」，
 *    那个 run 会被整个漏掉 —— 宁可漏也不要把半个 run 标成链接
 */
import type { DiagnosticSink } from '@uw/core';
import type { DocumentBody, NodeId, PropSet, RunNode } from './nodes.ts';
import { walkParagraphs } from './nodes.ts';

/** 域开关：`\o "1-3"` → `{ name: 'o', value: '1-3' }` */
export interface FieldSwitch {
  /** `\` 后面那**一个**字符，原样保留大小写。查开关用 `fieldSwitch()`，它不分大小写 */
  name: string;
  value?: string;
}

export interface FieldInstruction {
  /** 域类型，已大写（`PAGE` / `HYPERLINK` / `TOC`）。指令为空时是 `''` */
  type: string;
  /** 类型之后的位置参数，引号已剥、转义已还原 */
  args: string[];
  switches: FieldSwitch[];
}

/** 界桩在文档里的位置。三段式是因为 run id 之外还要知道它在哪个段落、run 里第几个片段 */
export interface FieldPoint {
  paragraphId: NodeId;
  runId: NodeId;
  /** 在 `run.content` 里的下标 */
  contentIndex: number;
}

export interface FieldRegion {
  /** `complex` 是 `w:fldChar` 三桩那种，`simple` 是 `w:fldSimple` 压缩写法 */
  kind: 'complex' | 'simple';
  instruction: FieldInstruction;
  /** 拼起来的原始指令文字。Word 会把一条指令切碎成好几个 `w:instrText`，这里已经接回去 */
  instructionText: string;
  /** 简单域没有界桩，这三个都缺席 —— 它的范围就是 `resultRuns` */
  begin?: FieldPoint;
  /** 缺席 = 这个域**没有结果**，什么都不显示 */
  separate?: FieldPoint;
  /** 缺席 = 文件里少了 end（已记诊断），当作域一直开到文档末尾但结果区为空 */
  end?: FieldPoint;
  /** 嵌套深度，最外层是 0 */
  depth: number;
  /** 结果区里的 run id，按文档顺序。复杂域缺 separate / end 时为空 */
  resultRuns: NodeId[];
}

// ── 指令解析 ──────────────────────────────────────────────────────────────────

interface Token {
  text: string;
  /** 未加引号、且以 `\` 开头 —— 引号里的 `\l` 是普通文字，不是开关 */
  isSwitch: boolean;
}

const WS = new Set([' ', '\t', '\n', '\r']);

/**
 * 指令分词。
 *
 * 引号内 `\` 一律是转义（`\"` 得到引号本身）；引号外的 `\` **原样留着** ——
 * `INCLUDEPICTURE C:\\pic\\a.png` 里的反斜杠是路径的一部分，剥掉就打不开了。
 */
function tokenize(instr: string): Token[] {
  const out: Token[] = [];
  let i = 0;
  while (i < instr.length) {
    const ch = instr.charAt(i);
    if (WS.has(ch)) {
      i++;
      continue;
    }
    if (ch === '"') {
      let text = '';
      i++;
      while (i < instr.length) {
        const c = instr.charAt(i);
        if (c === '\\' && i + 1 < instr.length) {
          text += instr.charAt(i + 1);
          i += 2;
          continue;
        }
        i++;
        if (c === '"') break;
        text += c;
      }
      out.push({ text, isSwitch: false });
      continue;
    }
    let text = '';
    while (i < instr.length) {
      const c = instr.charAt(i);
      if (WS.has(c) || c === '"') break;
      text += c;
      i++;
    }
    out.push({ text, isSwitch: text.startsWith('\\') && text.length > 1 });
  }
  return out;
}

/**
 * 指令文字 → 结构。
 *
 * **已知的简化**：「开关后面那个普通词是它的值」是个启发式。Word 真正的做法是每种域
 * 自带一张「哪些开关带值」的表（`\h` 不带、`\o` 带），照那张表才能百分百分对。
 * 现实里带值的开关后面必跟值、不带值的后面必跟另一个开关或结尾，所以启发式够用；
 * 真出现「不带值的开关后面直接跟位置参数」时，那个参数会被误当成开关的值。
 * 补的办法是加一张按域类型分的开关表，等真有域被这条坑到再加，不预先摆着。
 */
export function parseFieldInstruction(text: string): FieldInstruction {
  const args: string[] = [];
  const switches: FieldSwitch[] = [];
  let type = '';
  // 上一个还等着收值的开关
  let pending: FieldSwitch | undefined;

  for (const tok of tokenize(text)) {
    if (tok.isSwitch) {
      // `\o1-3` 这种不带空格的写法 Word 也认，所以第二个字符之后的部分直接当值
      const rest = tok.text.slice(2);
      const sw: FieldSwitch =
        rest === '' ? { name: tok.text.charAt(1) } : { name: tok.text.charAt(1), value: rest };
      switches.push(sw);
      pending = rest === '' ? sw : undefined;
      continue;
    }
    if (type === '') {
      // 域类型永远排在最前，不会被上一个开关抢走
      type = tok.text.toUpperCase();
      pending = undefined;
      continue;
    }
    if (pending !== undefined) {
      pending.value = tok.text;
      pending = undefined;
      continue;
    }
    args.push(tok.text);
  }
  return { type, args, switches };
}

/** 按名字取开关。**不分大小写** —— Word 里 `\O` 与 `\o` 是同一个开关 */
export function fieldSwitch(instr: FieldInstruction, name: string): FieldSwitch | undefined {
  const want = name.toLowerCase();
  return instr.switches.find((s) => s.name.toLowerCase() === want);
}

// ── 扫描 ──────────────────────────────────────────────────────────────────────

interface StreamRun<S extends PropSet> {
  paragraphId: NodeId;
  run: RunNode<S>;
}

interface Frame {
  /** 按 begin 的先后编号：域是嵌套结束的，不排一下顺序输出就成了内层在前 */
  order: number;
  depth: number;
  begin: FieldPoint;
  instr: string[];
  separate?: FieldPoint;
  separateRun?: number;
}

/**
 * 扫描整份 body，把域还原出来。
 *
 * 不平衡的界桩（多一个 end、少一个 end）记诊断后继续 —— 域坏了不该让整份文档白屏
 * （原则 1.5）。少 end 的那个仍然吐出来，因为回写时还得知道它在哪儿。
 */
export function scanFields<S extends PropSet>(
  body: DocumentBody<S>,
  diagnostics?: DiagnosticSink,
): FieldRegion[] {
  // 先拉平成一条 run 流：域跨段落，甚至能从表格外开始、在格子里结束，
  // 按文档顺序走一遍才是唯一能配对的视角
  const stream: StreamRun<S>[] = [];
  for (const p of walkParagraphs(body)) {
    for (const run of p.runs) stream.push({ paragraphId: p.id, run });
  }

  const stack: Frame[] = [];
  const done: { order: number; region: FieldRegion }[] = [];
  let order = 0;
  // 简单域：同一个 id 的连续 run 属于同一个域，第一次见到时建区、之后往里追加
  const simple = new Map<NodeId, FieldRegion>();

  for (let i = 0; i < stream.length; i++) {
    const entry = stream[i];
    if (entry === undefined) continue;

    const fld = entry.run.fieldSimple;
    if (fld !== undefined) {
      const seen = simple.get(fld.id);
      if (seen === undefined) {
        const region: FieldRegion = {
          kind: 'simple',
          instruction: parseFieldInstruction(fld.instr),
          instructionText: fld.instr,
          depth: stack.length,
          resultRuns: [entry.run.id],
        };
        simple.set(fld.id, region);
        done.push({ order: order++, region });
      } else {
        seen.resultRuns.push(entry.run.id);
      }
    }

    for (let j = 0; j < entry.run.content.length; j++) {
      const c = entry.run.content[j];
      if (c === undefined) continue;

      if (c.kind === 'fieldInstruction') {
        // 归**栈顶**那个域：内层域的指令不该混进外层（见文件头第 2 条）
        const top = stack[stack.length - 1];
        if (top !== undefined) top.instr.push(c.text);
        continue;
      }
      if (c.kind !== 'fieldChar') continue;

      const point: FieldPoint = { paragraphId: entry.paragraphId, runId: entry.run.id, contentIndex: j };
      if (c.charType === 'begin') {
        stack.push({ order: order++, depth: stack.length, begin: point, instr: [] });
      } else if (c.charType === 'separate') {
        const top = stack[stack.length - 1];
        if (top === undefined) {
          diagnostics?.warn('field-unbalanced', '遇到没有 begin 的 w:fldChar separate，已忽略');
        } else if (top.separate === undefined) {
          // 一个域只认第一个 separate；多出来的是坏文件，忽略比抛错更接近 Word
          top.separate = point;
          top.separateRun = i;
        }
      } else {
        const top = stack.pop();
        if (top === undefined) {
          diagnostics?.warn('field-unbalanced', '遇到没有 begin 的 w:fldChar end，已忽略');
          continue;
        }
        done.push({ order: top.order, region: finish(top, point, i, stream) });
      }
    }
  }

  for (const top of stack) {
    diagnostics?.warn('field-unclosed', `域 ${top.instr.join('').trim() || '(空指令)'} 缺少 w:fldChar end`);
    done.push({ order: top.order, region: finish(top, undefined, undefined, stream) });
  }

  return done.sort((a, b) => a.order - b.order).map((d) => d.region);
}

function finish<S extends PropSet>(
  frame: Frame,
  end: FieldPoint | undefined,
  endRun: number | undefined,
  stream: readonly StreamRun<S>[],
): FieldRegion {
  const instructionText = frame.instr.join('');
  const region: FieldRegion = {
    kind: 'complex',
    instruction: parseFieldInstruction(instructionText),
    instructionText,
    begin: frame.begin,
    depth: frame.depth,
    resultRuns: resultRuns(frame.separateRun, endRun, stream),
  };
  if (frame.separate !== undefined) region.separate = frame.separate;
  if (end !== undefined) region.end = end;
  return region;
}

/**
 * 结果区的 run。
 *
 * 取的是**严格夹在**两颗界桩之间的整 run —— 界桩所在的那两个 run 自己不算。
 * 真实文件里界桩独占一个 run，所以这与「界桩之间的全部内容」是同一件事。
 */
function resultRuns<S extends PropSet>(
  separateRun: number | undefined,
  endRun: number | undefined,
  stream: readonly StreamRun<S>[],
): NodeId[] {
  if (separateRun === undefined || endRun === undefined) return [];
  const out: NodeId[] = [];
  for (let i = separateRun + 1; i < endRun; i++) {
    const entry = stream[i];
    if (entry !== undefined) out.push(entry.run.id);
  }
  return out;
}

// ── HYPERLINK ────────────────────────────────────────────────────────────────

/**
 * HYPERLINK 域给出的链接。
 *
 * 与 `w:hyperlink` 容器那条路的区别只在**地址从哪儿来**：容器给的是关系 id（要查 rels
 * 才知道地址），域把地址**字面写在指令里**。所以这里多一个 `url`，两条路最终都落在
 * `RunNode.hyperlink` 上，渲染层不必认识「域」这个概念。
 */
export interface FieldHyperlink {
  url?: string;
  anchor?: string;
}

/**
 * 从扫描结果里挑出 HYPERLINK 域，铺成「run id → 链接」。
 *
 * 嵌套时内层赢：域按 begin 顺序排，内层排在外层之后，后写的覆盖先写的。
 */
export function fieldHyperlinks(regions: readonly FieldRegion[]): Map<NodeId, FieldHyperlink> {
  const out = new Map<NodeId, FieldHyperlink>();
  for (const f of regions) {
    if (f.instruction.type !== 'HYPERLINK') continue;
    const link: FieldHyperlink = {};
    const url = f.instruction.args[0];
    if (url !== undefined && url !== '') link.url = url;
    // `\l` 是书签名。`HYPERLINK \l "top"` 这种没有 url 的写法是文档内跳转
    const anchor = fieldSwitch(f.instruction, 'l')?.value;
    if (anchor !== undefined && anchor !== '') link.anchor = anchor;
    if (link.url === undefined && link.anchor === undefined) continue;
    for (const id of f.resultRuns) out.set(id, link);
  }
  return out;
}
