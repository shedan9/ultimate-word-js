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
 */
import type { Diagnostic } from '@uw/core';
import type { DocumentLayout } from '@uw/layout';
import type {
  DirectProps,
  DocPosition,
  DocRange,
  FindOptions,
  LoadedDocument,
  NodeId,
  QueryNode,
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
  queryNodes,
  rangeOfNode,
  walkBlocks,
} from '@uw/model';
import { imageHrefResolver } from '@uw/render-dom';
import type { UwView, ViewOptions } from './view.ts';
import { createView } from './view.ts';

/** `query()` 答的节点：段落 / run / 表格 / 行 / 格，直接格式那棵树上的 */
export type DocNode = QueryNode<DirectProps>;

export type { FindOptions };

export interface UwDocumentInit {
  loaded: LoadedDocument;
  reflow?: (body: LoadedDocument['body']) => {
    loaded: LoadedDocument;
    layout: DocumentLayout;
    values: ReadonlyMap<NodeId, string>;
  };
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
  /** 解析与排版期记下的内容问题（结构性错误早在 `load()` 就抛了） */
  readonly diagnostics: readonly Diagnostic[];
  #loaded: LoadedDocument;
  #fieldValues: ReadonlyMap<NodeId, string>;
  readonly #editor: TextEditor;
  readonly #reflow: UwDocumentInit['reflow'];
  #prepared: ReturnType<NonNullable<UwDocumentInit['reflow']>> | undefined;
  readonly #listeners = new Set<(change: TextChangeSet) => void>();
  readonly #views = new Set<(layout: DocumentLayout) => void>();
  #order: RunOrder | undefined;
  #nodes: Map<NodeId, DocNode> | undefined;
  #imageHref: ((id: string) => string | undefined) | undefined;

  constructor(init: UwDocumentInit) {
    this.#loaded = init.loaded;
    this.#layout = init.layout;
    this.#editor = createTextEditor(init.loaded.body, {
      validate: (body) => {
        this.#prepared = this.#reflow?.(body);
      },
    });
    this.#reflow = init.reflow;
    this.#fieldValues = init.fieldValues;
    this.diagnostics = init.diagnostics;
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
    return this.#changed(this.#editor.tx(callback, options));
  }
  undo(): TextChangeSet | undefined {
    return this.#changed(this.#editor.undo());
  }
  redo(): TextChangeSet | undefined {
    return this.#changed(this.#editor.redo());
  }
  #changed(change: TextChangeSet | undefined): TextChangeSet | undefined {
    if (!change || !this.#reflow) return change;
    const result = this.#prepared;
    if (!result) throw new Error('缺少预排版结果');
    this.#prepared = undefined;
    this.#loaded = result.loaded;
    this.#layout = result.layout;
    this.#fieldValues = result.values;
    this.#order = undefined;
    this.#nodes = undefined;
    for (const update of this.#views) update(this.layout);
    for (const listener of this.#listeners) listener(change);
    return change;
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
    );
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
