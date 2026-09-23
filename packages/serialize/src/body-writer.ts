/**
 * 可编辑的正文树 → `document.xml`，**以原文为底打补丁**。
 *
 * 模型是有损的：书签、批注范围、拼写标记、被删除的修订、内容控件的外壳、`w14:paraId`、
 * 不认识的段落属性……解析时都跳过了（parse-body.ts 的 IGNORED / TRANSPARENT）。
 * 从模型重新生成整份正文会把它们全丢掉，所以这里反过来：**走原文的 XML 树**，
 * 在模型节点对得上的位置用模型的内容，对不上的（不认识的）原样保留。
 *
 * 对应关系靠 `BodySources`：同一个部件重新解析一遍，id 与加载时一字不差，于是
 * 可编辑树上的每个原有节点都能找回自己的元素。三档处理，越往下动得越多：
 *
 * 1. **节点与原来结构相等** → 原元素整个吐回（绝大多数段落走这条，逐字节不变）
 * 2. **有原元素但改过** → 以原元素为底：属性容器按字段补丁（props-writer.ts），
 *    子节点按原来的顺序走一遍、run 换成模型里的样子，其余（书签、`w:del`…）照抄
 * 3. **没有原元素**（拆段 / 拆 run 新造的）→ 以**它从哪儿拆出来的那个**为模板：
 *    属性没变就沿用模板的 `w:pPr` / `w:rPr`，于是模板里我们不认识的格式（边框、底纹）
 *    也跟着过去 —— 与 Word 里回车继承上一段格式是同一个行为
 *
 * 编辑不会改变节点的相对顺序（事务只有插入 / 删除 / 拆 / 合，没有移动），所以两边
 * 按顺序对齐即可：原文里的节点在模型里还在就输出它，不在了（被合段吃掉）就跳过，
 * 紧跟在它后面的新节点就是从它拆出来的。
 */
import type { DiagnosticSink } from '@uw/core';
import type {
  Block,
  Body,
  BodySources,
  NodeId,
  Paragraph,
  Run,
  RunContent,
  Table,
  TableCell,
  TableRow,
  TableWidth,
} from '@uw/model';
import type { XmlDocument, XmlElement, XmlNode } from '@uw/ooxml';
import { child } from '@uw/ooxml';
import { PPR_ORDER, patchParaProps, patchRunProps } from './props-writer.ts';
import { el, insertOrdered, patchAttrs, same, withoutChildren } from './xml-edit.ts';

/** run 外面那几层会压平进 run 的容器（parse-body.ts 的 `collectRuns`） */
const RUN_CONTAINERS = new Set([
  'w:hyperlink',
  'w:fldSimple',
  'w:smartTag',
  'w:bdo',
  'w:dir',
  'w:ins',
  'w:moveTo',
  'w:sdt',
]);

/** 重新生成 run 内容时丢掉的：Word 的分页缓存，内容一改就过期了，留着反而误导 */
const STALE_IN_RUN = new Set(['w:lastRenderedPageBreak']);

/** 按顺序消费的一列模型节点 */
interface Cursor<T extends { id: NodeId }> {
  items: readonly T[];
  i: number;
  ids: ReadonlySet<NodeId>;
}

function cursor<T extends { id: NodeId }>(items: readonly T[]): Cursor<T> {
  return { items, i: 0, ids: new Set(items.map((x) => x.id)) };
}

/** run 身上那两个压平的外层容器标记 */
interface Marks {
  link: Run['hyperlink'];
  field: Run['fieldSimple'];
}

const NO_MARKS: Marks = { link: undefined, field: undefined };

function sameMarks(run: Run, marks: Marks): boolean {
  return same(run.hyperlink, marks.link) && same(run.fieldSimple, marks.field);
}

export interface BodyWriteInput {
  /** 原文的 `document.xml` */
  original: XmlDocument;
  /** 同一份原文重新解析出来的树与源元素表（`parseBody(…, sources)`） */
  originalBody: Body;
  sources: BodySources;
  /** 编辑后的树 */
  body: Body;
  diagnostics?: DiagnosticSink;
  part?: string;
}

export function writeBody(input: BodyWriteInput): XmlDocument {
  return new BodyWriter(input).write();
}

class BodyWriter {
  readonly #in: BodyWriteInput;
  /** 元素 → 它解析出来的节点 id（`BodySources.nodes` 的反查） */
  readonly #idOf = new Map<XmlElement, NodeId>();
  readonly #orig = new Map<NodeId, Block | Run>();
  /** 非末节最后一段 → 它该带的 `w:sectPr`。末节的在 `w:body` 末尾，原样留着 */
  readonly #sectFor = new Map<NodeId, XmlElement>();
  /** 图片 / 符号 / 域界桩的源元素池：拆 run 之后它们搬进了新 run，得凭内容找回原元素 */
  #pool: Map<string, XmlElement[]> | undefined;
  readonly #used = new Set<XmlElement>();

  constructor(input: BodyWriteInput) {
    this.#in = input;
    for (const [id, e] of input.sources.nodes) this.#idOf.set(e, id);
    for (const section of input.originalBody.sections) collect(section.blocks, this.#orig);
    const sections = input.body.sections;
    for (let s = 0; s < sections.length - 1; s++) {
      const section = sections[s];
      const last = section?.blocks.at(-1);
      const sectPr = section === undefined ? undefined : input.sources.nodes.get(section.id);
      if (last !== undefined && sectPr !== undefined) this.#sectFor.set(last.id, sectPr);
    }
  }

  write(): XmlDocument {
    const { original, body } = this.#in;
    const bodyEl = child(original.root, 'w:body');
    if (bodyEl === undefined) throw new Error('原文没有 <w:body>，无从回写');
    const cur = cursor(body.sections.flatMap((s) => s.blocks));
    const out = this.#blocks(bodyEl.children, cur);
    // 剩下的新块（原文末尾那段被拆出来的）要落在末节 `w:sectPr` **之前** —— 它必须是 body 的最后一个子元素
    const rest = this.#rest(cur);
    const last = out.at(-1);
    const tail = last?.kind === 'element' && last.name === 'w:sectPr' ? out.length - 1 : out.length;
    const children = [...out.slice(0, tail), ...rest, ...out.slice(tail)];
    const root = {
      ...original.root,
      children: original.root.children.map((c) => (c === bodyEl ? { ...bodyEl, children } : c)),
    };
    return { declaration: original.declaration, root };
  }

  // ── 块 ──────────────────────────────────────────────────────────────────────

  #blocks(nodes: readonly XmlNode[], cur: Cursor<Block>): XmlNode[] {
    const out: XmlNode[] = [];
    for (const node of nodes) {
      if (node.kind !== 'element') {
        out.push(node);
        continue;
      }
      const id = this.#idOf.get(node);
      if ((node.name === 'w:p' || node.name === 'w:tbl') && id !== undefined) {
        if (!cur.ids.has(id)) continue; // 合段吃掉了
        let template: Block | undefined;
        while (cur.items[cur.i]?.id !== id) {
          const b = cur.items[cur.i] as Block;
          if (this.#orig.has(b.id)) throw new Error(`回写时块顺序对不上：${b.id} 出现在 ${id} 之前`);
          out.push(this.#block(b, template));
          cur.i++;
        }
        template = cur.items[cur.i++] as Block;
        out.push(this.#block(template, undefined));
        // 紧跟着的新块是从它拆出来的，以它为模板
        for (let b = cur.items[cur.i]; b !== undefined && !this.#orig.has(b.id); b = cur.items[cur.i]) {
          out.push(this.#block(b, template));
          cur.i++;
        }
      } else if (node.name === 'w:sdt') {
        // 块级内容控件：外壳照抄，里面的块照常对齐（拆出来的新段留在控件里面）
        const content = child(node, 'w:sdtContent');
        if (content === undefined) out.push(node);
        else {
          const inner = { ...content, children: this.#blocks(content.children, cur) };
          out.push({ ...node, children: node.children.map((c) => (c === content ? inner : c)) });
        }
      } else out.push(node);
    }
    return out;
  }

  /** 走完原文还剩下的模型节点：只可能是新的（原有的都对上了），挨着最后一个输出 */
  #rest(cur: Cursor<Block>): XmlNode[] {
    const out: XmlNode[] = [];
    let template: Block | undefined = cur.items[cur.i - 1];
    for (; cur.i < cur.items.length; cur.i++) {
      const b = cur.items[cur.i] as Block;
      out.push(this.#block(b, template));
      template = b;
    }
    return out;
  }

  #block(b: Block, template: Block | undefined): XmlElement {
    if (b.kind === 'table') return this.#table(b);
    return this.#paragraph(b, template?.kind === 'paragraph' ? template : undefined);
  }

  // ── 段落 ────────────────────────────────────────────────────────────────────

  #paragraph(p: Paragraph, template: Paragraph | undefined): XmlElement {
    const src = this.#in.sources.nodes.get(p.id);
    const orig = this.#orig.get(p.id);
    const wantSect = this.#sectFor.get(p.id);
    const srcPPr = src === undefined ? undefined : child(src, 'w:pPr');
    const srcSect = srcPPr === undefined ? undefined : child(srcPPr, 'w:sectPr');
    if (src !== undefined && same(orig, p) && srcSect === wantSect) return src;

    // 新段落以模板的 pPr 为底：属性没变就整个沿用，模板里我们不认识的格式也跟过来
    const base = src ?? (template === undefined ? undefined : this.#in.sources.nodes.get(template.id));
    const basePPr = base === undefined ? undefined : child(base, 'w:pPr');
    let pPr = patchParaProps(basePPr, p.props);
    // 分节符跟着「本节最后一段」走：拆了末段，它要搬到新的末段上去
    if (pPr !== undefined) {
      const children = pPr.children.filter((c) => c.kind !== 'element' || c.name !== 'w:sectPr');
      pPr = children.length === 0 && Object.keys(pPr.attrs).length === 0 ? undefined : { ...pPr, children };
    }
    if (wantSect !== undefined) {
      pPr = { ...(pPr ?? el('w:pPr')), children: insertOrdered(pPr?.children ?? [], wantSect, PPR_ORDER) };
    }

    const runs = cursor(p.runs);
    const home = new Set<NodeId>();
    if (src !== undefined) for (const r of runsIn(src, this.#idOf)) home.add(r);
    const inner = src === undefined ? [] : this.#runs(src.children, runs, NO_MARKS, home);
    const children: XmlNode[] = [];
    if (pPr !== undefined) children.push(pPr);
    // 新段落的第一个 run 是从模板段落的末 run 拆出来的（拆段时右半截取新 id，text-transaction.ts）
    const runTemplate = runs.items[runs.i - 1] ?? (src === undefined ? template?.runs.at(-1) : undefined);
    children.push(...inner, ...this.#looseRuns(runs.items.slice(runs.i), NO_MARKS, runTemplate));
    // 新段落不抄模板的属性：`w14:paraId` 要求全文唯一，抄过去就是两段同一个 id
    return el('w:p', src?.attrs ?? {}, children);
  }

  /**
   * 按原文顺序走段落的子节点。`home` 是**原来就在这一段里**的 run —— 只有它们能当锚；
   * 合段搬过来的 run 虽然也有源元素，但它的位置在别的段落里，这里当新 run 处理。
   */
  #runs(nodes: readonly XmlNode[], cur: Cursor<Run>, marks: Marks, home: ReadonlySet<NodeId>): XmlNode[] {
    const out: XmlNode[] = [];
    const anchored = (r: Run) => home.has(r.id);
    /** 跟在刚输出的那个后面的新 run，外层标记相同才留在这一层容器里 */
    const trailing = () => {
      for (
        let r = cur.items[cur.i];
        r !== undefined && !anchored(r) && sameMarks(r, marks);
        r = cur.items[cur.i]
      ) {
        out.push(this.#run(r, cur.items[cur.i - 1]));
        cur.i++;
      }
    };
    for (const node of nodes) {
      if (node.kind !== 'element') {
        out.push(node);
        continue;
      }
      if (node.name === 'w:pPr') continue;
      if (node.name === 'w:r') {
        const id = this.#idOf.get(node);
        if (id === undefined || !cur.ids.has(id)) continue; // 整个删掉了
        const before: Run[] = [];
        const prev = cur.items[cur.i - 1];
        while (cur.items[cur.i]?.id !== id) {
          const r = cur.items[cur.i] as Run;
          if (anchored(r)) throw new Error(`回写时 run 顺序对不上：${r.id} 出现在 ${id} 之前`);
          before.push(r);
          cur.i++;
        }
        out.push(...this.#looseRuns(before, marks, prev));
        out.push(this.#run(cur.items[cur.i++] as Run, undefined));
        trailing();
      } else if (RUN_CONTAINERS.has(node.name)) {
        // 进容器之前，先把挂在外面这一层的新 run 放下 —— 它们在容器之前
        trailing();
        const content = node.name === 'w:sdt' ? child(node, 'w:sdtContent') : node;
        if (content === undefined) {
          out.push(node);
          continue;
        }
        const inner = this.#containerMarks(node, marks);
        const start = cur.i;
        const children = this.#runs(content.children, cur, inner, home);
        // 容器里原有的 run 全删光了，容器也不留（空超链接在 Word 里是个点不中的幽灵）
        if (hasModelRun(content.children, this.#idOf) && cur.i === start) continue;
        const rebuilt = { ...content, children };
        out.push(
          content === node
            ? rebuilt
            : { ...node, children: node.children.map((c) => (c === content ? rebuilt : c)) },
        );
        trailing();
      } else out.push(node); // 书签、批注范围、w:del、w:proofErr……原样
    }
    return out;
  }

  #containerMarks(node: XmlElement, marks: Marks): Marks {
    if (node.name === 'w:hyperlink') {
      const link: NonNullable<Run['hyperlink']> = {};
      const relId = node.attrs['r:id'];
      const anchor = node.attrs['w:anchor'];
      if (relId !== undefined) link.relId = relId;
      if (anchor !== undefined) link.anchor = anchor;
      return { ...marks, link };
    }
    if (node.name === 'w:fldSimple') {
      const id = this.#idOf.get(node);
      return id === undefined ? marks : { ...marks, field: { id, instr: node.attrs['w:instr'] ?? '' } };
    }
    return marks;
  }

  /**
   * 没有锚的 run：按外层标记分组，标记与所在层不同的包一层 `w:hyperlink` / `w:fldSimple`。
   * 合段把一段带链接的文字搬到另一段末尾时走这条 —— 链接得跟着过去。
   */
  #looseRuns(runs: readonly Run[], marks: Marks, template: Run | undefined): XmlNode[] {
    const out: XmlNode[] = [];
    let prev = template;
    for (let i = 0; i < runs.length; ) {
      const first = runs[i] as Run;
      const group: XmlElement[] = [];
      for (
        ;
        i < runs.length &&
        same(runs[i]?.hyperlink, first.hyperlink) &&
        same(runs[i]?.fieldSimple, first.fieldSimple);
        i++
      ) {
        const r = runs[i] as Run;
        group.push(this.#run(r, prev));
        prev = r;
      }
      let nodes: XmlNode[] = group;
      if (first.fieldSimple !== undefined && !same(first.fieldSimple, marks.field)) {
        const src = this.#in.sources.nodes.get(first.fieldSimple.id);
        nodes = [el('w:fldSimple', src?.attrs ?? { 'w:instr': first.fieldSimple.instr }, nodes)];
      }
      if (first.hyperlink !== undefined && !same(first.hyperlink, marks.link)) {
        const attrs: Record<string, string> = {};
        if (first.hyperlink.relId !== undefined) attrs['r:id'] = first.hyperlink.relId;
        if (first.hyperlink.anchor !== undefined) attrs['w:anchor'] = first.hyperlink.anchor;
        nodes = [el('w:hyperlink', attrs, nodes)];
      }
      out.push(...nodes);
    }
    return out;
  }

  // ── run ─────────────────────────────────────────────────────────────────────

  #run(run: Run, template: Run | undefined): XmlElement {
    const src = this.#in.sources.nodes.get(run.id);
    const orig = this.#orig.get(run.id) as Run | undefined;
    const sameContent = orig !== undefined && same(orig.content, run.content);
    if (src !== undefined && orig !== undefined && sameContent && same(orig.props, run.props)) return src;

    const base = src ?? (template === undefined ? undefined : this.#in.sources.nodes.get(template.id));
    const rPr = patchRunProps(base === undefined ? undefined : child(base, 'w:rPr'), run.props);
    const content =
      src !== undefined && sameContent
        ? src.children.filter((c) => c.kind !== 'element' || c.name !== 'w:rPr')
        : this.#content(run, src, orig);
    return el('w:r', src?.attrs ?? {}, rPr === undefined ? content : [rPr, ...content]);
  }

  /**
   * 重新生成 run 的内容。文字、制表位、换行这些照模型写；图片、符号、域界桩**找回原元素** ——
   * 模型只存了它们的一部分（图片只有外框与引用，没有 `a:graphic` 的其余部分），重新生成是有损的。
   * 原 run 里不进模型的子元素（脚注引用、批注引用）按「在所有内容之前 / 之后」放回去。
   */
  #content(run: Run, src: XmlElement | undefined, orig: Run | undefined): XmlNode[] {
    const origins = src === undefined ? [] : (this.#in.sources.content.get(run.id) ?? []);
    const lead: XmlNode[] = [];
    const tail: XmlNode[] = [];
    if (src !== undefined) {
      const firstOrigin = src.children.findIndex((c) => c.kind === 'element' && origins.includes(c));
      src.children.forEach((c, i) => {
        if (c.kind !== 'element' || c.name === 'w:rPr' || origins.includes(c) || STALE_IN_RUN.has(c.name))
          return;
        (firstOrigin >= 0 && i < firstOrigin ? lead : tail).push(c);
      });
    }
    let scan = 0;
    const own = (c: RunContent): XmlElement | undefined => {
      if (orig === undefined) return undefined;
      for (let j = scan; j < orig.content.length; j++) {
        const e = origins[j];
        if (e !== undefined && !this.#used.has(e) && same(orig.content[j], c)) {
          scan = j + 1;
          this.#used.add(e);
          return e;
        }
      }
      return undefined;
    };
    const out: XmlNode[] = [];
    for (const c of run.content) {
      const node = this.#item(c, own);
      if (node !== undefined) out.push(node);
    }
    return [...lead, ...out, ...tail];
  }

  #item(c: RunContent, own: (c: RunContent) => XmlElement | undefined): XmlNode | undefined {
    switch (c.kind) {
      case 'text':
        // 删空的文字片段是占槽位的（text-transaction.ts），文件里不需要它
        if (c.text === '') return undefined;
        return el('w:t', /^\s|\s$/.test(c.text) ? { 'xml:space': 'preserve' } : {}, [
          { kind: 'text', text: c.text },
        ]);
      case 'tab':
        return el('w:tab');
      case 'break':
        return el('w:br', c.breakType === 'line' ? {} : { 'w:type': c.breakType });
      case 'noBreakHyphen':
        return el('w:noBreakHyphen');
      case 'softHyphen':
        return el('w:softHyphen');
      case 'fieldInstruction':
        return el('w:instrText', { 'xml:space': 'preserve' }, [{ kind: 'text', text: c.text }]);
      case 'symbol':
      case 'fieldChar':
      case 'object': {
        const found = own(c) ?? this.#fromPool(c);
        if (found !== undefined) return found;
        if (c.kind === 'symbol') {
          const code = c.char.codePointAt(0) ?? 0;
          return el('w:sym', {
            'w:font': c.font,
            'w:char': code.toString(16).toUpperCase().padStart(4, '0'),
          });
        }
        if (c.kind === 'fieldChar') return el('w:fldChar', { 'w:fldCharType': c.charType });
        // 图形只有原元素才写得出来；编辑不会凭空造图，走到这里说明对应关系断了 —— 留痕，不静默丢
        this.#in.diagnostics?.warn('serialize-object-lost', '回写时找不到图形的原始 XML，这个对象没有写出', {
          ...(this.#in.part === undefined ? {} : { part: this.#in.part }),
        });
        return undefined;
      }
    }
  }

  #fromPool(c: RunContent): XmlElement | undefined {
    if (this.#pool === undefined) {
      this.#pool = new Map();
      for (const [id, origins] of this.#in.sources.content) {
        const run = this.#orig.get(id) as Run | undefined;
        run?.content.forEach((item, j) => {
          const e = origins[j];
          if (
            e === undefined ||
            (item.kind !== 'object' && item.kind !== 'symbol' && item.kind !== 'fieldChar')
          )
            return;
          const key = stableKey(item);
          const list = this.#pool?.get(key);
          if (list) list.push(e);
          else this.#pool?.set(key, [e]);
        });
      }
    }
    const list = this.#pool.get(stableKey(c));
    const e = list?.find((x) => !this.#used.has(x));
    if (e !== undefined) this.#used.add(e);
    return e;
  }

  // ── 表格 ──────────────────────────────────────────────────────────────────
  // 行与列都可以插、可以删（`insertRow` / `deleteRows` / `insertColumn` / `deleteColumns`），
  // 跟着变的只有结构：`w:tblGrid`、`w:tblW`、行的 `w:gridBefore` / `w:gridAfter`、格的 `w:tcW` /
  // `w:gridSpan` / `w:vMerge`。表格 / 行 / 格的其余属性还改不了，原样照抄

  #table(t: Table): XmlElement {
    const src = this.#in.sources.nodes.get(t.id);
    if (src === undefined) throw new Error(`回写时找不到表格 ${t.id} 的原文（编辑不会新建表格）`);
    const orig = this.#orig.get(t.id) as Table | undefined;
    if (same(orig, t)) return src;
    const templates = this.#rowTemplates(t.rows);
    const cur = cursor(t.rows);
    const out = this.#tableChildren(src.children, cur, templates).map((n) => {
      if (n.kind !== 'element') return n;
      if (n.name === 'w:tblPr' && !same(orig?.props.width, t.props.width))
        return patchChild(n, 'w:tblW', t.props.width, TBLPR_ORDER);
      if (n.name === 'w:tblGrid' && !same(orig?.grid, t.grid)) {
        const cols = t.grid.map((w) => el('w:gridCol', { 'w:w': String(Math.round(w)) }));
        // `w:tblGridChange`（修订）等排在 gridCol 后面的原样留着
        return { ...n, children: [...cols, ...withoutChildren(n.children, new Set(['w:gridCol']))] };
      }
      return n;
    });
    // 剩下的新行跟在最后一个 `w:tr` 后面（表格末尾一般就是它，扩展元素排在它后面的也不挪）
    const rest = t.rows.slice(cur.i).map((r) => this.#newRow(r, templates));
    const last = out.findLastIndex((n) => n.kind === 'element' && (n.name === 'w:tr' || n.name === 'w:sdt'));
    const children = [...out.slice(0, last + 1), ...rest, ...out.slice(last + 1)];
    return { ...src, children };
  }

  /**
   * 每个新行照着哪个原有行抄 XML。新行是照着它的上一行（下方插入）或下一行（上方插入）抄的，
   * 只看顺序分不出是哪一个 —— 「在 B 上方插」与「在 A 下方插」得到的模型顺序一模一样，
   * 而表头行下方插的那一行与 B 的 `trPr` 完全不同。所以拿**结构与属性**比：新行是模板的精确拷贝
   * （table-edit.ts 的 `withInsertedRow`），与它一致的那个邻居就是模板；两边都一致时任取，XML 本来就相同。
   */
  #rowTemplates(rows: readonly TableRow[]): Map<NodeId, TableRow> {
    const out = new Map<NodeId, TableRow>();
    const isOrig = (r: TableRow | undefined) => r !== undefined && this.#in.sources.nodes.has(r.id);
    const shape = (r: TableRow) => [r.props, r.propsEx, r.cells.map((c) => [c.props, c.gridSpan])];
    rows.forEach((row, i) => {
      if (isOrig(row)) return;
      let prev: TableRow | undefined;
      for (let j = i - 1; j >= 0 && prev === undefined; j--) if (isOrig(rows[j])) prev = rows[j];
      let next: TableRow | undefined;
      for (let j = i + 1; j < rows.length && next === undefined; j++) if (isOrig(rows[j])) next = rows[j];
      const want = shape(row);
      const pick = [prev, next].find((r) => r !== undefined && same(shape(r), want)) ?? prev ?? next;
      if (pick !== undefined) out.set(row.id, pick);
    });
    return out;
  }

  /** 原有行在**原文**里的样子（模型里的它可能已经被改过 vMerge / 格数） */
  #origRow(row: TableRow): TableRow | undefined {
    const table = [...this.#orig.values()].find(
      (b): b is Table => b.kind === 'table' && b.rows.some((r) => r.id === row.id),
    );
    return table?.rows.find((r) => r.id === row.id);
  }

  /**
   * 第 `index` 格照着哪个**原有**格抄 XML：它自己是原有的就是它；插列新造的格照同一行里与它属性
   * 一致的邻居（table-edit.ts 的 `withInsertedColumn` 照目标那一侧的邻居抄，宽度另算，所以比的时候
   * 不看 `width`），邻居都是新格时退到这一行原文里的格。新行里的格由调用方先换成模板行里同一下标的格再来问。
   */
  #cellTemplate(row: TableRow, index: number): TableCell | undefined {
    const cell = row.cells[index];
    if (cell === undefined) return undefined;
    const own = this.#origCell(cell.id);
    if (own !== undefined) return own;
    const key = (c: TableCell) => [{ ...c.props, width: undefined }, c.vMerge];
    const near: TableCell[] = [];
    for (let d = 1; d < row.cells.length; d++)
      for (const j of [index - d, index + d]) {
        const o = row.cells[j] === undefined ? undefined : this.#origCell((row.cells[j] as TableCell).id);
        if (o !== undefined) near.push(o);
      }
    // 同一行里原有的格全删光了（先插列再删掉原有列）：退到这一行在原文里的格，新格当初就是照其中一个抄的
    near.push(...(this.#origRow(row)?.cells ?? []));
    return near.find((o) => same(key(o), key(cell))) ?? near[0];
  }

  #tableChildren(
    nodes: readonly XmlNode[],
    cur: Cursor<TableRow>,
    templates: ReadonlyMap<NodeId, TableRow>,
  ): XmlNode[] {
    const out: XmlNode[] = [];
    const isNew = (r: TableRow | undefined) => r !== undefined && !this.#in.sources.nodes.has(r.id);
    for (const node of nodes) {
      if (node.kind !== 'element') {
        out.push(node);
        continue;
      }
      if (node.name === 'w:sdt') {
        out.push(this.#inSdt(node, (c) => this.#tableChildren(c, cur, templates)));
        continue;
      }
      const id = node.name === 'w:tr' ? this.#idOf.get(node) : undefined;
      if (id === undefined) {
        out.push(node);
        continue;
      }
      if (!cur.ids.has(id)) continue; // 删掉的行
      while (cur.items[cur.i]?.id !== id) {
        const r = cur.items[cur.i] as TableRow;
        if (!isNew(r)) throw new Error(`回写时行顺序对不上：${r.id} 出现在 ${id} 之前`);
        out.push(this.#newRow(r, templates));
        cur.i++;
      }
      const row = cur.items[cur.i++] as TableRow;
      out.push(this.#existingRow(node, row));
      for (let r = cur.items[cur.i]; isNew(r); r = cur.items[cur.i]) {
        out.push(this.#newRow(r as TableRow, templates));
        cur.i++;
      }
    }
    return out;
  }

  /** 原有的行：删掉的格不写、新格插在模型里的位置、`trPr` 的跳过列数跟着改 */
  #existingRow(node: XmlElement, row: TableRow): XmlElement {
    const before = this.#origRow(row);
    const cells = cursor(row.cells);
    let children = this.#rowChildren(node.children, row, cells);
    if (cells.i < row.cells.length) {
      // 原有的格一个不剩（删列后又插列）或新格在末尾：跟在最后一个格后面
      const rest = row.cells.slice(cells.i).map((_, k) => this.#newCell(row, cells.i + k, row));
      const last = children.findLastIndex(
        (n) => n.kind === 'element' && (n.name === 'w:tc' || n.name === 'w:sdt'),
      );
      const at = last >= 0 ? last + 1 : children.length;
      children = [...children.slice(0, at), ...rest, ...children.slice(at)];
    }
    if (
      before !== undefined &&
      (before.props.gridBefore !== row.props.gridBefore || before.props.gridAfter !== row.props.gridAfter)
    ) {
      const trPr = child(node, 'w:trPr');
      let next = trPr ?? el('w:trPr');
      for (const name of ['gridBefore', 'gridAfter'] as const) {
        const v = row.props[name];
        next = replaceChild(
          next,
          `w:${name}`,
          v === undefined ? undefined : el(`w:${name}`, { 'w:val': String(v) }),
          TRPR_ORDER,
        );
      }
      const empty = next.children.length === 0 && Object.keys(next.attrs).length === 0;
      if (trPr !== undefined) children = children.flatMap((c) => (c === trPr ? (empty ? [] : [next]) : [c]));
      else if (!empty) {
        const ex = children.findIndex((c) => c.kind === 'element' && c.name === 'w:tblPrEx');
        children = [...children.slice(0, ex + 1), next, ...children.slice(ex + 1)];
      }
    }
    return { ...node, children };
  }

  #rowChildren(nodes: readonly XmlNode[], row: TableRow, cur: Cursor<TableCell>): XmlNode[] {
    const out: XmlNode[] = [];
    const isNew = (c: TableCell | undefined) => c !== undefined && !this.#in.sources.nodes.has(c.id);
    for (const node of nodes) {
      if (node.kind !== 'element') {
        out.push(node);
        continue;
      }
      if (node.name === 'w:sdt') {
        out.push(this.#inSdt(node, (c) => this.#rowChildren(c, row, cur)));
        continue;
      }
      const id = node.name === 'w:tc' ? this.#idOf.get(node) : undefined;
      if (id === undefined) {
        out.push(node);
        continue;
      }
      if (!cur.ids.has(id)) continue; // 删掉的列
      while (cur.items[cur.i]?.id !== id) {
        if (!isNew(cur.items[cur.i]))
          throw new Error(`回写时格顺序对不上：${cur.items[cur.i]?.id} 出现在 ${id} 之前`);
        out.push(this.#newCell(row, cur.i++, row));
      }
      const cell = cur.items[cur.i++] as TableCell;
      const blocks = cursor(cell.blocks);
      let children = [...this.#blocks(node.children, blocks), ...this.#rest(blocks)];
      const tcPr = child(node, 'w:tcPr');
      const next = patchTcPr(tcPr, cell, this.#origCell(id));
      if (next !== tcPr)
        children =
          tcPr === undefined
            ? next === undefined
              ? children
              : [next, ...children]
            : children.flatMap((c) => (c === tcPr ? (next === undefined ? [] : [next]) : [c]));
      out.push({ ...node, children });
      while (isNew(cur.items[cur.i])) out.push(this.#newCell(row, cur.i++, row));
    }
    return out;
  }

  /**
   * 新格的 XML：`w:tcPr` 抄模板格（`#cellTemplate`）再按字段改结构那几项，格内段落以模板格的首段为模板 ——
   * 模型不认识的格底纹扩展、段落边框跟着走。`tplRow` 是新格抄的那一行：原有行里就是它自己，
   * 新行里是它的模板行（新行照模板行逐格抄，同一下标就是同一格）。
   */
  #newCell(row: TableRow, index: number, tplRow: TableRow): XmlElement {
    const cell = row.cells[index] as TableCell;
    const tpl = this.#cellTemplate(tplRow, index);
    const tplSrc = tpl === undefined ? undefined : this.#in.sources.nodes.get(tpl.id);
    const tcPr = patchTcPr(tplSrc === undefined ? undefined : child(tplSrc, 'w:tcPr'), cell, tpl);
    let template: Block | undefined = tpl?.blocks.find((b) => b.kind === 'paragraph');
    const blocks = cell.blocks.map((b) => {
      const e = this.#block(b, template);
      template = b;
      return e;
    });
    return el('w:tc', {}, [...(tcPr === undefined ? [] : [tcPr]), ...blocks]);
  }

  /**
   * 新行的 XML：`w:tr` 的属性与 `w:tblPrEx` / `w:trPr` 抄模板行，每格照模板行同一下标的格抄（`#newCell`）。
   * `w14:paraId` / `w14:textId` 不抄，它们要求全文唯一。
   */
  #newRow(row: TableRow, templates: ReadonlyMap<NodeId, TableRow>): XmlElement {
    const tplRow = templates.get(row.id);
    const tplSrc = tplRow === undefined ? undefined : this.#in.sources.nodes.get(tplRow.id);
    const head =
      tplSrc?.children.filter(
        (c): c is XmlElement => c.kind === 'element' && (c.name === 'w:tblPrEx' || c.name === 'w:trPr'),
      ) ?? [];
    const cells = row.cells.map((_, i) => this.#newCell(row, i, tplRow ?? row));
    const attrs = Object.fromEntries(
      Object.entries(tplSrc?.attrs ?? {}).filter(([k]) => k !== 'w14:paraId' && k !== 'w14:textId'),
    );
    return el('w:tr', attrs, [...head, ...cells]);
  }

  #origCells: Map<NodeId, TableCell> | undefined;
  #origCell(id: NodeId): TableCell | undefined {
    if (this.#origCells === undefined) {
      this.#origCells = new Map();
      for (const b of this.#orig.values())
        if (b.kind === 'table') for (const r of b.rows) for (const c of r.cells) this.#origCells.set(c.id, c);
    }
    return this.#origCells.get(id);
  }

  #inSdt(node: XmlElement, map: (children: readonly XmlNode[]) => XmlNode[]): XmlElement {
    const content = child(node, 'w:sdtContent');
    if (content === undefined) return node;
    const inner = { ...content, children: map(content.children) };
    return { ...node, children: node.children.map((c) => (c === content ? inner : c)) };
  }
}

function collect(blocks: readonly Block[], out: Map<NodeId, Block | Run>): void {
  for (const b of blocks) {
    out.set(b.id, b);
    if (b.kind === 'paragraph') for (const r of b.runs) out.set(r.id, r);
    else for (const row of b.rows) for (const cell of row.cells) collect(cell.blocks, out);
  }
}

/** 段落元素里（含压平的容器里）原有的 run id */
function runsIn(p: XmlElement, idOf: ReadonlyMap<XmlElement, NodeId>): NodeId[] {
  const out: NodeId[] = [];
  const walk = (nodes: readonly XmlNode[]) => {
    for (const n of nodes) {
      if (n.kind !== 'element') continue;
      if (n.name === 'w:r') {
        const id = idOf.get(n);
        if (id !== undefined) out.push(id);
      } else if (RUN_CONTAINERS.has(n.name))
        walk(n.name === 'w:sdt' ? (child(n, 'w:sdtContent')?.children ?? []) : n.children);
    }
  };
  walk(p.children);
  return out;
}

function hasModelRun(nodes: readonly XmlNode[], idOf: ReadonlyMap<XmlElement, NodeId>): boolean {
  return nodes.some(
    (n) =>
      n.kind === 'element' &&
      ((n.name === 'w:r' && idOf.has(n)) ||
        (RUN_CONTAINERS.has(n.name) &&
          hasModelRun(n.name === 'w:sdt' ? (child(n, 'w:sdtContent')?.children ?? []) : n.children, idOf))),
  );
}

/** 键序无关的 JSON，给内容片段当池的 key */
function stableKey(v: unknown): string {
  if (typeof v !== 'object' || v === null) return JSON.stringify(v);
  if (Array.isArray(v)) return `[${v.map(stableKey).join(',')}]`;
  const o = v as Record<string, unknown>;
  return `{${Object.keys(o)
    .filter((k) => o[k] !== undefined)
    .sort()
    .map((k) => `${JSON.stringify(k)}:${stableKey(o[k])}`)
    .join(',')}}`;
}

/** `w:tcPr` 的 schema 顺序（§17.4.66），插 `w:vMerge` 要落在对的位置 */
const TCPR_ORDER = [
  'w:cnfStyle',
  'w:tcW',
  'w:gridSpan',
  'w:hMerge',
  'w:vMerge',
  'w:tcBorders',
  'w:shd',
  'w:noWrap',
  'w:tcMar',
  'w:textDirection',
  'w:tcFitText',
  'w:vAlign',
  'w:hideMark',
  'w:headers',
  'w:cellIns',
  'w:cellDel',
  'w:cellMerge',
  'w:tcPrChange',
];

/** `w:tblPr` 的 schema 顺序（§17.4.60），改 `w:tblW` 用 */
const TBLPR_ORDER = [
  'w:tblStyle',
  'w:tblpPr',
  'w:tblOverlap',
  'w:bidiVisual',
  'w:tblStyleRowBandSize',
  'w:tblStyleColBandSize',
  'w:tblW',
  'w:jc',
  'w:tblCellSpacing',
  'w:tblInd',
  'w:tblBorders',
  'w:shd',
  'w:tblLayout',
  'w:tblCellMar',
  'w:tblLook',
  'w:tblCaption',
  'w:tblDescription',
  'w:tblPrChange',
];

/** `w:trPr` 的子元素顺序。规范里它是 choice 不限次数，但 Word 自己按这个顺序写，照着插最稳 */
const TRPR_ORDER = [
  'w:cnfStyle',
  'w:divId',
  'w:gridBefore',
  'w:gridAfter',
  'w:wBefore',
  'w:wAfter',
  'w:cantSplit',
  'w:trHeight',
  'w:tblHeader',
  'w:tblCellSpacing',
  'w:jc',
  'w:hidden',
  'w:ins',
  'w:del',
  'w:trPrChange',
];

/** 换掉（`node` 为 undefined 就删掉）容器里名为 `name` 的子元素，新元素按 schema 顺序插 */
function replaceChild(
  parent: XmlElement,
  name: string,
  node: XmlElement | undefined,
  order: readonly string[],
): XmlElement {
  const kept = withoutChildren(parent.children, new Set([name]));
  return { ...parent, children: node === undefined ? kept : insertOrdered(kept, node, order) };
}

/** 宽度元素（`w:tblW` / `w:tcW`）按新值改 `w:w` / `w:type`，别的属性原样；没有宽度就删掉 */
function patchChild(
  parent: XmlElement,
  name: string,
  width: TableWidth | undefined,
  order: readonly string[],
): XmlElement {
  if (width === undefined) return replaceChild(parent, name, undefined, order);
  const old = child(parent, name);
  const attrs = patchAttrs(old?.attrs ?? {}, ['w:w', 'w:type'], {
    'w:w': String(Math.round(width.value)),
    'w:type': width.type,
  });
  return replaceChild(parent, name, { ...(old ?? el(name)), attrs }, order);
}

/**
 * 按字段改 `w:tcPr` 里结构的三项（`w:tcW` / `w:gridSpan` / `w:vMerge`），只动与 `base`（XML 来自的那一格）
 * 不同的几项，其余子元素原样。没有改动就原样返回同一个对象；结果为空就不要这个容器。
 */
function patchTcPr(
  tcPr: XmlElement | undefined,
  cell: TableCell,
  base: TableCell | undefined,
): XmlElement | undefined {
  let next = tcPr ?? el('w:tcPr');
  if (base?.vMerge !== cell.vMerge)
    next = replaceChild(
      next,
      'w:vMerge',
      cell.vMerge === 'none'
        ? undefined
        : el('w:vMerge', cell.vMerge === 'restart' ? { 'w:val': 'restart' } : {}),
      TCPR_ORDER,
    );
  if ((base?.gridSpan ?? 1) !== cell.gridSpan)
    next = replaceChild(
      next,
      'w:gridSpan',
      cell.gridSpan > 1 ? el('w:gridSpan', { 'w:val': String(cell.gridSpan) }) : undefined,
      TCPR_ORDER,
    );
  if (!same(base?.props.width, cell.props.width))
    next = patchChild(next, 'w:tcW', cell.props.width, TCPR_ORDER);
  if (next === tcPr) return tcPr;
  if (next.children.length === 0 && Object.keys(next.attrs).length === 0) return undefined;
  return next;
}
