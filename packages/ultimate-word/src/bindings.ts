/**
 * `doc.bindings` —— 模板填充（api.md §9）。
 *
 * 坑位就是 OOXML 原生的内容控件（`w:sdt`），按 `w:tag` 认（没有 tag 的按 `w:alias`）。
 * 填完导回 docx，坑位仍是坑位（外壳原样留着），在 Word 里可以接着改 —— 不会变成一段死文字。
 *
 * `set` 只是记下来，`apply` 才提交：一次 `apply` = 一个事务 = 一次重排 = 一个撤销单元
 * （api.md §10 的四件事对齐），填 20 个字段不会排 20 遍。
 *
 * 它不另起算法：列出坑位是 `@uw/model` 的 `contentControlSpans`，填值是事务命令
 * `fillContentControl`（换内容、换掉占位符格式、清占位标记都在那里）。这一层只管
 * 「名字 → 控件」与「攒着一起提交」。
 */
import type { Body, ContentControlType, DocRange, NodeId, TextChangeSet, TextTransaction } from '@uw/model';
import { contentControlSpans, rangeOfContentControl, textOfContentControl } from '@uw/model';

export interface BindingInfo {
  /** 坑位名：`w:tag`，没有 tag 时是 `w:alias`。同名的几个控件一起填 */
  key: string;
  tag?: string;
  alias?: string;
  type: ContentControlType;
  /** 行内（包着一段文字）还是块级（包着整段 / 整表） */
  scope: 'inline' | 'block';
  /** 现在显示的文字（段落之间、软换行都是 `\n`）。占位状态下是占位文字 */
  value: string;
  /** 里面还是占位文字（`w:showingPlcHdr`），还没人填过 */
  placeholder: boolean;
  /** 锁了内容（`w:lock` 为 contentLocked / sdtContentLocked），填不进去 */
  locked: boolean;
  /** 下拉框 / 组合框的选项（显示文字） */
  options?: string[];
  /** 控件内容的范围，可以拿去 `decorate` / `scrollTo` */
  range: DocRange | undefined;
}

interface BindingsHost {
  body: () => Body;
  tx: (callback: (tx: TextTransaction) => undefined) => TextChangeSet | undefined;
}

/** 能填文字的类型；复选框、图片、文档部件（目录外壳）只列出来，不能 `set` */
const FILLABLE: ReadonlySet<ContentControlType> = new Set([
  'richText',
  'text',
  'date',
  'dropDownList',
  'comboBox',
]);

export class Bindings {
  readonly #host: BindingsHost;
  readonly #pending = new Map<string, string>();

  constructor(host: BindingsHost) {
    this.#host = host;
  }

  /** 文档里的坑位，按文档序。没有 tag 也没有 alias 的控件不算坑位（没法按名字填） */
  list(): BindingInfo[] {
    return this.#spans().map(({ key, span }) => {
      const c = span.control;
      const info: BindingInfo = {
        key,
        type: c.type,
        scope: c.scope,
        value: textOfContentControl(span),
        placeholder: c.showingPlaceholder,
        locked: c.lock === 'contentLocked' || c.lock === 'sdtContentLocked',
        range: rangeOfContentControl(span),
      };
      if (c.tag !== undefined) info.tag = c.tag;
      if (c.alias !== undefined) info.alias = c.alias;
      if (c.items !== undefined) info.options = c.items.map((i) => i.text);
      return info;
    });
  }

  /**
   * 记下一个坑位的值，`apply()` 时才写进文档。**名字不存在就抛** —— 拼错字段名
   * 静默不填，是模板填充里最难发现的一类错。换行写成软换行、`\t` 写成制表位
   */
  set(key: string, value: string): this {
    if (typeof value !== 'string') throw new TypeError(`坑位「${key}」的值必须是字符串`);
    const targets = this.#spans().filter((s) => s.key === key);
    if (!targets.length) throw new RangeError(`文档里没有坑位：${key}`);
    const unfillable = targets.find((s) => !FILLABLE.has(s.span.control.type));
    if (unfillable) throw new Error(`坑位「${key}」是 ${unfillable.span.control.type} 控件，不能填文字`);
    this.#pending.set(key, value);
    return this;
  }

  /**
   * 一次记多个。与 `set` 不同，**文档里没有的名字跳过**：数据对象常常带着模板用不上的字段
   * （同一份数据套好几份模板），为此逐个挑字段太啰嗦。返回值里的 `skipped` 就是跳过的那些
   */
  setMany(values: Readonly<Record<string, string>>): { skipped: string[] } {
    const keys = new Set(this.#spans().map((s) => s.key));
    const skipped: string[] = [];
    for (const [key, value] of Object.entries(values)) {
      if (keys.has(key)) this.set(key, value);
      else skipped.push(key);
    }
    return { skipped };
  }

  /** 还没提交的值 */
  get pending(): ReadonlyMap<string, string> {
    return this.#pending;
  }

  /** 丢掉还没提交的值 */
  discard(): void {
    this.#pending.clear();
  }

  /**
   * 一次性提交：一个事务、一次重排、一个撤销单元。任何一个坑位填不进去（锁了、下拉框没有这个选项）
   * 整批回滚并抛错，攒着的值**留着**，改好了再 `apply`。值与现状完全相同时不产生事务，返回 undefined
   */
  apply(): TextChangeSet | undefined {
    if (!this.#pending.size) return undefined;
    // 倒着填：嵌套的两个坑位都给了值时，先填内层、再让外层把它整个盖掉（与 Word 里覆盖外层内容一致）；
    // 正着填的话外层先把内层删掉，内层就找不到自己了
    const jobs: [NodeId, string][] = this.#spans()
      .filter((s) => this.#pending.has(s.key))
      .map((s) => [s.span.control.id, this.#pending.get(s.key) as string] as [NodeId, string])
      .reverse();
    const change = this.#host.tx((t) => {
      for (const [id, value] of jobs) t.fillContentControl(id, value);
      return undefined;
    });
    this.#pending.clear();
    return change;
  }

  #spans() {
    return contentControlSpans(this.#host.body()).flatMap((span) => {
      const key = span.control.tag ?? span.control.alias;
      return key === undefined ? [] : [{ key, span }];
    });
  }
}
