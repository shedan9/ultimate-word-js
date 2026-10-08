/**
 * `UwDocument` —— 内容 + 布局，不关心屏幕（api.md §2 ①）。
 *
 * 加载一次、排版一次；`pageCount` 是文档的固有属性，不是某个视图的。`find` / `query` /
 * `compare` / `rangeOf` 全是**同步**的：布局结果已经在内存里，没有理由返回 Promise。
 *
 * 它只是把 `@uw/model` / `@uw/layout` 的低层入口按 api.md 的形状收拢，**不另起算法**：
 * 查找是 `findText`（吃级联完的树，自动带上求值过的域），选择器是 `queryNodes`
 * （走可编辑的那棵树），文档序是 `order.ts` 的那一套。两处这一层自己补的：
 * ① `rangeOf` 认 `NodeId`，所以要有一张 id → 节点的表（懒建，`query` 之外的入口才用得上）；
 * ② `compare` 遇到不在树上的 run **抛错**而不是答 undefined —— 门面的调用方拿到的位置
 *    都是这份文档自己吐出来的，比不了只能是拿错了文档。
 *
 * 事件（api.md §11）只在**加载之后**的变化上触发：`load()` 返回之前没人来得及挂监听者，
 * 那一趟的排版结果就是 `layout` / `pageCount`，诊断就是 `diagnostics` 的初值。
 */
import type { Diagnostic } from '@uw/core';
import type { DocumentLayout } from '@uw/layout';
import { indentCharUnit, paragraphPageNumbers } from '@uw/layout';
import type {
  DirectProps,
  DocPosition,
  DocRange,
  FindOptions,
  LoadedDocument,
  NodeId,
  QueryNode,
  ResolvedRun,
  RevisionSpan,
  RevisionView,
  RunOrder,
  TextChangeSet,
  TextEditor,
  TextMatch,
  TextTransaction,
  TextTransactionOptions,
  TocPlan,
} from '@uw/model';
import {
  bookmarkTargets,
  buildRunOrder,
  compareDocPositions,
  createTextEditor,
  expandReplacement,
  findMatches,
  findText,
  fragmentOfRange,
  listRevisions,
  paragraphStyleNames,
  paragraphsOfRange,
  planTableOfContents,
  queryNodes,
  rangeEditable,
  rangeOfNode,
  runPropsAtInsertion,
  runPropsOfRange,
  textOfRange,
  tocFields,
  walkBlocks,
  walkParagraphs,
} from '@uw/model';
import type { OpcPackage } from '@uw/ooxml';
import { imageHrefResolver } from '@uw/render-dom';
import { serializeDocx } from '@uw/serialize';
import { Bindings } from './bindings.ts';
import { Emitter } from './events.ts';
import type { Disposable, ElementHit, UwView, ViewOptions } from './view.ts';
import { createView } from './view.ts';

/** `query()` 答的节点：段落 / run / 表格 / 行 / 格，直接格式那棵树上的 */
export type DocNode = QueryNode<DirectProps>;

export type { FindOptions, RevisionSpan, RevisionView, TextMatch };

/** `replaceAll` 的结果。跳过的是落在域（页码、目录、超链接域的显示文字）里的命中，事务改不了它们 */
export interface ReplaceResult {
  replaced: number;
  skipped: number;
}

/** `updateTableOfContents` 的结果。跳过的目录照旧显示（原因如「不支持图表目录（\c）」） */
export interface TocUpdateResult {
  updated: number;
  skipped: { reason: string }[];
}

/** docx 的 MIME，`toDocx()` 的 Blob 带着它，下载时浏览器才知道扩展名 */
export const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

export interface DocumentEvents {
  /**
   * 事务 / 撤销 / 重做之后的那一趟排版完成（含域求值的全部迭代）。`duration` 是级联 + 排版的
   * 毫秒数，`iterations` 是域求值排了几趟（1 = 没有要迭代的域）。
   */
  'layout:done': { pageCount: number; duration: number; iterations: number };
  /**
   * 模型变了，在 `layout:done` 之前派发（两者同步相继，此时 `doc.layout` 已经是新的）。
   * `source` 区分事务与撤销 / 重做 —— 「标记未保存」只看有没有这个事件，协同同步才要分清。
   */
  'document:change': { changeSet: TextChangeSet; source: 'tx' | 'undo' | 'redo' };
  /** 重排时新发现的内容问题（比如插进来的文字用了缺失的字体）。同一条不重复报 */
  diagnostic: Diagnostic;
}

interface Reflowed {
  loaded: LoadedDocument;
  layout: DocumentLayout;
  values: ReadonlyMap<NodeId, string>;
  /** 域求值排了几趟；缺省按 1 */
  passes?: number;
  /** 这一趟新记下的诊断（诊断表是追加式的，只要增量） */
  diagnostics?: readonly Diagnostic[];
}

export interface UwDocumentInit {
  loaded: LoadedDocument;
  /** 原包。回写只重写改过的部件，其余条目从这里逐字节照搬；没有它就不能 `toDocx()` */
  pkg?: OpcPackage;
  reflow?: (body: LoadedDocument['body']) => Reflowed;
  /** 换修订视图：把下一趟 `reflow` 要用的页眉页脚 / 注 / 文本框按新视图重新级联（见 load.ts） */
  setRevisionView?: (view: RevisionView) => void;
  layout: DocumentLayout;
  /** 与 `layout` 自洽的域求值结果（run id → 显示的文字），查找时用它跳过旧值 */
  fieldValues: ReadonlyMap<NodeId, string>;
  diagnostics: readonly Diagnostic[];
}

export class UwDocument {
  /**
   * 排版结果。**不在稳定性承诺内**（api.md §16）：它是流水线的中间数据，增量排版与
   * Worker 化都会动它的形状。暴露出来是给调试台与保真度工具用的，产品代码别依赖它
   */
  #layout: DocumentLayout;
  get layout(): DocumentLayout {
    return this.#layout;
  }
  /**
   * 解析与排版期记下的内容问题（结构性错误早在 `load()` 就抛了）。编辑后重排新发现的
   * 会追加在后面（去重），与 `diagnostic` 事件报的是同一批
   */
  get diagnostics(): readonly Diagnostic[] {
    return this.#diagnostics;
  }
  readonly #diagnostics: Diagnostic[];
  readonly #diagnosticKeys: Set<string>;
  readonly #events = new Emitter<DocumentEvents>();
  #loaded: LoadedDocument;
  #fieldValues: ReadonlyMap<NodeId, string>;
  readonly #editor: TextEditor;
  readonly #reflow: UwDocumentInit['reflow'];
  readonly #setRevisionView: UwDocumentInit['setRevisionView'];
  readonly #pkg: OpcPackage | undefined;
  #prepared: (Reflowed & { duration: number }) | undefined;
  readonly #listeners = new Set<(change: TextChangeSet) => void>();
  readonly #views = new Set<(layout: DocumentLayout) => void>();
  #order: RunOrder | undefined;
  #nodes: Map<NodeId, DocNode> | undefined;
  #resolvedRuns: Map<NodeId, ResolvedRun> | undefined;
  #imageHref: ((id: string) => string | undefined) | undefined;
  /** 模板填充（api.md §9）：按内容控件的 tag 填值，`apply()` 一次提交 */
  readonly bindings: Bindings;

  constructor(init: UwDocumentInit) {
    this.#loaded = init.loaded;
    this.#layout = init.layout;
    this.#editor = createTextEditor(init.loaded.body, {
      validate: (body) => {
        if (!this.#reflow) return;
        const t0 = performance.now();
        const result = this.#reflow(body);
        this.#prepared = { ...result, duration: performance.now() - t0 };
      },
    });
    this.#reflow = init.reflow;
    this.#setRevisionView = init.setRevisionView;
    this.#pkg = init.pkg;
    this.#fieldValues = init.fieldValues;
    this.#diagnostics = [...init.diagnostics];
    this.#diagnosticKeys = new Set(this.#diagnostics.map(diagnosticKey));
    this.bindings = new Bindings({ body: () => this.#editor.body, tx: (callback) => this.tx(callback) });
  }

  /** 挂一个事件监听者，见 `DocumentEvents`。视图上的事件在 `UwView.on` */
  on<K extends keyof DocumentEvents>(type: K, listener: (payload: DocumentEvents[K]) => void): Disposable {
    return this.#events.on(type, listener);
  }

  get canUndo(): boolean {
    return this.#editor.canUndo;
  }
  get canRedo(): boolean {
    return this.#editor.canRedo;
  }
  tx(
    callback: (tx: TextTransaction) => undefined,
    options?: TextTransactionOptions,
  ): TextChangeSet | undefined {
    if (!this.#reflow) throw new Error('文档未配置编辑重排');
    return this.#changed(this.#editor.tx(callback, options), 'tx');
  }
  undo(): TextChangeSet | undefined {
    return this.#changed(this.#editor.undo(), 'undo');
  }
  redo(): TextChangeSet | undefined {
    return this.#changed(this.#editor.redo(), 'redo');
  }
  #changed(
    change: TextChangeSet | undefined,
    source: DocumentEvents['document:change']['source'],
  ): TextChangeSet | undefined {
    if (!change || !this.#reflow) return change;
    const result = this.#prepared;
    if (!result) throw new Error('缺少预排版结果');
    this.#prepared = undefined;
    const fresh = this.#adopt(result);
    for (const listener of this.#listeners) listener(change);
    this.#events.emit('document:change', { changeSet: change, source });
    this.#events.emit('layout:done', {
      pageCount: this.pageCount,
      duration: result.duration,
      iterations: result.passes ?? 1,
    });
    for (const d of fresh) this.#events.emit('diagnostic', d);
    return change;
  }

  /**
   * 修订显示哪一版（`LoadOptions.revisions`，默认 `final`）。见 `setRevisionView`
   */
  get revisionView(): RevisionView {
    return this.#loaded.revisions.view;
  }

  /**
   * 换修订视图并重排：`final` 最终状态（被删的字不占位）/ `markup` 所有标记（作者色，插入加下划线、
   * 删除加删除线并**占位**，页边画改动竖线）/ `original` 原始状态（插入的字不占位）。
   *
   * 不是模型修改：不进撤销栈、不派发 `document:change`，只派发 `layout:done`。
   * 选区与装饰的位置照旧有效（模型没动），但落在这一版看不见的字上的那些画不出来
   */
  setRevisionView(view: RevisionView): void {
    if (view === this.revisionView) return;
    if (!this.#reflow || !this.#setRevisionView) throw new Error('文档未配置重排，不能切换修订视图');
    this.#setRevisionView(view);
    const t0 = performance.now();
    const result = this.#reflow(this.#editor.body);
    const fresh = this.#adopt(result);
    this.#events.emit('layout:done', {
      pageCount: this.pageCount,
      duration: performance.now() - t0,
      iterations: result.passes ?? 1,
    });
    for (const d of fresh) this.#events.emit('diagnostic', d);
  }

  /**
   * 文档里的修订，文档序；同一段里相邻、作者 / 时间 / 种类都相同的合成一处。
   * 与这一刻显示哪一版无关 —— 答的是文件里记着什么（给修订面板用）
   */
  revisions(): RevisionSpan[] {
    return listRevisions(this.#editor.body);
  }

  /** 收下一趟重排的结果、刷新视图，返回新发现的诊断（事件由调用方按自己的顺序派发） */
  #adopt(result: Reflowed): Diagnostic[] {
    this.#loaded = result.loaded;
    this.#layout = result.layout;
    this.#fieldValues = result.values;
    this.#order = undefined;
    this.#nodes = undefined;
    this.#resolvedRuns = undefined;
    const fresh: Diagnostic[] = [];
    for (const d of result.diagnostics ?? []) {
      const key = diagnosticKey(d);
      if (this.#diagnosticKeys.has(key)) continue;
      this.#diagnosticKeys.add(key);
      this.#diagnostics.push(d);
      fresh.push(d);
    }
    // 视图先刷新、再告诉宿主：监听者里去读 `view.selection` / `view.rectsOf` 拿到的是新布局上的
    for (const update of this.#views) update(this.layout);
    return fresh;
  }

  /**
   * 导出 docx（api.md §13）。**round-trip 安全**：没编辑过的文档每个部件与原文件逐字节相同；
   * 编辑过的只重写正文（与新建列表时的编号定义），模型不认识的 XML（书签、修订、边框…）原样留着。
   *
   * 返回 Promise 是 api.md 定的形状 —— 现在的实现是同步的（几 MB 的 zip 压缩是毫秒级），
   * 留着异步是为了将来挪进 Worker 不改签名。
   */
  async toDocx(): Promise<Blob> {
    if (this.#pkg === undefined) throw new Error('文档没有原包，无法导出 docx');
    const bytes = serializeDocx(this.#pkg, this.#editor.body);
    return new Blob([bytes as Uint8Array<ArrayBuffer>], { type: DOCX_MIME });
  }

  get pageCount(): number {
    return this.layout.pages.length;
  }

  /**
   * 按文本找。跨 run、不跨段落；字符串默认不分大小写（Word 的默认），正则跟自己的 flags。
   * 求值过的域自动跳过 —— 它们显示的是算出来的页码，文件里存的旧值搜到了也画不到屏幕上
   */
  find(pattern: string | RegExp, options: Omit<FindOptions, 'fieldValues'> = {}): DocRange[] {
    return findText(this.#loaded.resolved, pattern, { ...options, fieldValues: this.#fieldValues });
  }

  /**
   * 全部替换：一个事务、一次重排、一个撤销单元（与 `bindings.apply` 同理，逐处 `tx` 就是 N 次重排）。
   * 查找规则同 `find`；新文字取每处**首字**的格式（`tx.replaceText`），`\t` 写成制表位、换行写成软换行。
   * 正则查找时字符串替换按 `String.prototype.replace` 展开 `$1` / `$<name>` / `$&`；
   * 字符串查找时替换串**原样**写入 —— 查找串本身就不当正则，`$` 也不该有特殊意思（Word 的查找框同理）。
   * 命中落在域里的跳过并计数，不让一处失败把整批回滚。
   */
  replaceAll(
    pattern: string | RegExp,
    replacement: string | ((match: TextMatch) => string),
    options: Omit<FindOptions, 'fieldValues' | 'limit'> = {},
  ): ReplaceResult {
    const matches = findMatches(this.#loaded.resolved, pattern, {
      ...options,
      fieldValues: this.#fieldValues,
    });
    const body = this.#editor.body;
    const usable = matches.filter((m) => rangeEditable(body, m.range));
    const text = (m: TextMatch): string => {
      if (typeof replacement === 'function') {
        const out = replacement(m);
        if (typeof out !== 'string') throw new TypeError('替换函数必须返回字符串');
        return out;
      }
      return typeof pattern === 'string' ? replacement : expandReplacement(replacement, m);
    };
    // 先把替换文字全算出来：回调抛错时一个字都还没改
    const planned = usable.map((m) => ({ range: m.range, text: text(m) }));
    if (planned.length) {
      // 倒着换：删除不挪槽位、插入只动同一 run 里后面的片段，所以前面的命中位置一直有效
      this.tx((t) => {
        for (const p of planned.reverse()) t.replaceText(p.range, p.text);
      });
    }
    return { replaced: planned.length, skipped: matches.length - usable.length };
  }

  /**
   * 更新目录 —— Word 的「更新目录 → 更新整个目录」：按现在的标题重新生成每个 TOC 域的条目
   * （标题增删改之后目录跟着变），一个事务、一个撤销单元。条目的页码只是初值（导出的 docx 里存它），
   * 显示时 PAGEREF 仍按书签所在页重算。标题没有 `_Toc` 书签的顺手补一个，缺 `toc N` 样式的补定义。
   *
   * **不会自动发生**：Word 打开、编辑都不动目录，要用户点。宿主想「保存前刷新」就在保存前调它。
   * 图表目录（`\c`）等不支持的写法跳过、照旧显示，原因在 `skipped` 里
   */
  updateTableOfContents(): TocUpdateResult {
    const loaded = this.#loaded;
    const body = this.#editor.body;
    const inBody = new Set<NodeId>();
    for (const p of walkParagraphs(body)) inBody.add(p.id);
    const styleNames = paragraphStyleNames(loaded.cascade.styles, body.styles);
    const normalStyleId =
      loaded.cascade.styles.all().find((s) => s.type === 'paragraph' && s.isDefault)?.id ?? '';
    const pages = paragraphPageNumbers(this.#layout, loaded.resolved);
    const assigned = new Map<NodeId, string>();
    const plans: TocPlan[] = [];
    const skipped: TocUpdateResult['skipped'] = [];
    for (const region of tocFields(loaded.fields)) {
      if (!inBody.has(region.begin?.paragraphId ?? '')) continue;
      const result = planTableOfContents(loaded.resolved, body, loaded.fields, region, {
        styleNames,
        normalStyleId,
        pageText: (id) => pages.get(id),
        assigned,
      });
      if (!result.ok) {
        skipped.push({ reason: result.reason });
        continue;
      }
      // 后面的目录看得见前面补的书签与样式，不重名、不重复补
      for (const b of result.plan.bookmarks) assigned.set(b.paragraphId, b.name);
      for (const s of result.plan.styles) styleNames.set(s.id, s.name);
      plans.push(result.plan);
    }
    if (plans.length) {
      this.tx((t) => {
        for (const plan of plans) {
          for (const s of plan.styles) t.addStyle(s);
          for (const b of plan.bookmarks) t.addBookmark(b.paragraphId, b.name);
          t.replaceFieldResult(plan.field, plan.paragraphs);
        }
      });
    }
    return { updated: plans.length, skipped };
  }

  /** 按结构找：`paragraph[styleId=Heading1]`、`table > row:first-child cell`。支持的语法见 api.md §7 */
  query(selector: string): DocNode[] {
    return queryNodes(this.#loaded.body, selector);
  }

  /** 文档序。两个位置都得是这份文档里的 run，否则抛错 */
  compare(a: DocPosition, b: DocPosition): -1 | 0 | 1 {
    this.#order ??= buildRunOrder(this.#loaded.body);
    const r = compareDocPositions(this.#order, a, b);
    if (r === undefined) throw new Error(`位置不在这份文档里：${a.nodeId} / ${b.nodeId}`);
    return r;
  }

  /**
   * 一个节点覆盖的 range（首 run 开头到末 run 结尾）。空段落返回段落 id 的折叠范围。
   * 不存在的 id 返回 undefined
   */
  rangeOf(node: NodeId | DocNode): DocRange | undefined {
    const target = typeof node === 'string' ? this.#nodeById(node) : node;
    return target === undefined ? undefined : rangeOfNode(target);
  }

  /**
   * 书签起点所在段落的 range —— 目录条目「跳转」的落点：`click:element` 给的 `href` 是
   * `#_Toc…`，去掉 `#` 交给它，再把结果交给 `view.scrollTo()`（视图不替宿主跳转，api.md §11）。
   * 只精确到**段落**（书签只记了起点在哪一段，见 `ParagraphNode.bookmarks`）；没有这个书签答 undefined
   */
  rangeOfBookmark(name: string): DocRange | undefined {
    const target = bookmarkTargets(this.#loaded.body).get(name);
    return target === undefined ? undefined : this.rangeOf(target);
  }

  /** 挂到一个容器上（选择器或元素），**先清空容器**。一份文档可以挂多个视图 */
  mount(target: string | Element, options: ViewOptions = {}): UwView {
    const container = typeof target === 'string' ? globalThis.document?.querySelector(target) : target;
    if (container === null || container === undefined) {
      throw new Error(`挂载目标不存在：${typeof target === 'string' ? target : '(element)'}`);
    }
    // 图片 → data URI 的编码每份文档只做一次，挂几个视图共用一份缓存
    this.#imageHref ??= imageHrefResolver(this.#loaded.images);
    if (options.mode === 'edit' && !this.#reflow) throw new Error('文档未配置编辑重排');
    const owner = this;
    const editor: TextEditor = {
      get body() {
        return owner.#editor.body;
      },
      get canUndo() {
        return owner.canUndo;
      },
      get canRedo() {
        return owner.canRedo;
      },
      tx: (callback, opts) => this.tx(callback, opts),
      undo: () => this.undo(),
      redo: () => this.redo(),
      breakHistory: () => this.#editor.breakHistory(),
    };
    return createView(
      container,
      this.layout,
      options,
      this.#imageHref,
      {
        editor,
        text: (range) => textOfRange(this.#loaded.resolved, range, this.#fieldValues),
        fragment: (range) => fragmentOfRange(this.#loaded.resolved, range, this.#fieldValues),
        format: (range, pending) =>
          pending
            ? runPropsAtInsertion(this.#loaded.cascade, this.#editor.body, range.start, pending)
            : runPropsOfRange(this.#loaded.resolved, range),
        paragraphFormat: (range) =>
          paragraphsOfRange(this.#loaded.resolved, range).map((p) => ({ ...p, charUnit: indentCharUnit(p) })),
        tabStop: this.#loaded.cascade.settings.defaultTabStop,
        styles: this.#loaded.cascade.styles
          .all()
          .filter((s) => s.type === 'paragraph')
          .map((s) => ({ id: s.id, name: s.name, next: s.next, isDefault: s.isDefault })),
        subscribe: (listener) => {
          this.#listeners.add(listener);
          return () => {
            this.#listeners.delete(listener);
          };
        },
      },
      {
        subscribe: (update) => {
          this.#views.add(update);
          return () => {
            this.#views.delete(update);
          };
        },
      },
      (position) => this.#elementAt(position),
    );
  }

  /**
   * 点到的位置落在什么「元素」上（`click:element` 的数据）。现在只认超链接：
   * 图片的绘制层不接指针事件、布局索引里也没有对象的矩形；内容控件（run 上的标记）
   * 还没接 —— 点坑位要答哪一层控件（嵌套时）还没定。
   */
  #elementAt(position: DocPosition): ElementHit | undefined {
    if (this.#resolvedRuns === undefined) {
      const map = new Map<NodeId, ResolvedRun>();
      for (const section of this.#loaded.resolved.sections)
        for (const block of walkBlocks(section.blocks))
          if (block.kind === 'paragraph') for (const run of block.runs) map.set(run.id, run);
      this.#resolvedRuns = map;
    }
    const link = this.#resolvedRuns.get(position.nodeId)?.hyperlink;
    const node = this.#nodeById(position.nodeId);
    if (link === undefined || node === undefined) return undefined;
    // 容器那条路给的是关系 id，外部关系的 Target 原文就是 URL；书签锚点拼成 `#名字`
    let href = link.url;
    if (href === undefined && link.relId !== undefined && this.#pkg !== undefined) {
      const rel = this.#pkg.rels(this.#pkg.mainDocumentPartName()).byId(link.relId);
      if (rel?.targetMode === 'External') href = rel.rawTarget;
    }
    if (link.anchor !== undefined) href = `${href ?? ''}#${link.anchor}`;
    const hit: ElementHit = { kind: 'hyperlink', node };
    if (href !== undefined) hit.href = href;
    return hit;
  }

  #nodeById(id: NodeId): DocNode | undefined {
    if (this.#nodes === undefined) {
      const map = new Map<NodeId, DocNode>();
      for (const section of this.#loaded.body.sections) {
        for (const block of walkBlocks(section.blocks)) {
          map.set(block.id, block);
          if (block.kind === 'paragraph') for (const run of block.runs) map.set(run.id, run);
          else {
            for (const row of block.rows) {
              map.set(row.id, row);
              for (const cell of row.cells) map.set(cell.id, cell);
            }
          }
        }
      }
      this.#nodes = map;
    }
    return this.#nodes.get(id);
  }
}

function diagnosticKey(d: Diagnostic): string {
  return [d.severity, d.code, d.message, d.part ?? '', d.path ?? ''].join('\u0000');
}
