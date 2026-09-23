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
import { indentCharUnit } from '@uw/layout';
import type {
  DirectProps,
  DocPosition,
  DocRange,
  FindOptions,
  LoadedDocument,
  NodeId,
  QueryNode,
  ResolvedRun,
  RunOrder,
  TextChangeSet,
  TextEditor,
  TextTransaction,
  TextTransactionOptions,
} from '@uw/model';
import {
  buildRunOrder,
  compareDocPositions,
  createTextEditor,
  findText,
  fragmentOfRange,
  paragraphsOfRange,
  queryNodes,
  rangeOfNode,
  runPropsOfRange,
  textOfRange,
  walkBlocks,
} from '@uw/model';
import type { OpcPackage } from '@uw/ooxml';
import { imageHrefResolver } from '@uw/render-dom';
import { serializeDocx } from '@uw/serialize';
import { Emitter } from './events.ts';
import type { Disposable, ElementHit, UwView, ViewOptions } from './view.ts';
import { createView } from './view.ts';

/** `query()` 答的节点：段落 / run / 表格 / 行 / 格，直接格式那棵树上的 */
export type DocNode = QueryNode<DirectProps>;

export type { FindOptions };

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
  readonly #pkg: OpcPackage | undefined;
  #prepared: (Reflowed & { duration: number }) | undefined;
  readonly #listeners = new Set<(change: TextChangeSet) => void>();
  readonly #views = new Set<(layout: DocumentLayout) => void>();
  #order: RunOrder | undefined;
  #nodes: Map<NodeId, DocNode> | undefined;
  #resolvedRuns: Map<NodeId, ResolvedRun> | undefined;
  #imageHref: ((id: string) => string | undefined) | undefined;

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
    this.#pkg = init.pkg;
    this.#fieldValues = init.fieldValues;
    this.#diagnostics = [...init.diagnostics];
    this.#diagnosticKeys = new Set(this.#diagnostics.map(diagnosticKey));
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
        format: (range) => runPropsOfRange(this.#loaded.resolved, range),
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
   * 图片的绘制层不接指针事件、布局索引里也没有对象的矩形，内容控件解析时已经剥掉了。
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
