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
import {
  contentControlSpans,
  formatContentControlDate,
  isContentControlLocked,
  rangeOfContentControl,
  textOfContentControl,
} from '@uw/model';

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
  /** 复选框的勾选状态；value 仍是显示符号，避免把文字接口改成混合类型 */
  checked?: boolean;
  /** 日期选择器的 OOXML 日期时间值；value 仍是模板格式化后的显示文字 */
  fullDate?: string;
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

/** 日历日用明确的对象区分显示文字，JSON 数据不受 Date 可变性与本机时区影响。 */
export interface DateBindingValue {
  date: string;
}

/** 文字填字符串，复选框填布尔值，日期填 { date: 'YYYY-MM-DD' }。 */
export type BindingValue = string | boolean | DateBindingValue;

/** 能填文字的类型；复选框走独立的状态命令，图片与文档部件（目录外壳）只列出 */
const FILLABLE: ReadonlySet<ContentControlType> = new Set([
  'richText',
  'text',
  'date',
  'dropDownList',
  'comboBox',
]);

export class Bindings {
  readonly #host: BindingsHost;
  readonly #pending = new Map<string, BindingValue>();

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
        locked: isContentControlLocked(this.#host.body().contentControls ?? {}, c.id),
        range: rangeOfContentControl(span),
      };
      if (c.tag !== undefined) info.tag = c.tag;
      if (c.alias !== undefined) info.alias = c.alias;
      if (c.items !== undefined) info.options = c.items.map((i) => i.text);
      if (c.checkbox !== undefined) info.checked = c.checkbox.checked;
      if (c.date?.fullDate !== undefined) info.fullDate = c.date.fullDate;
      return info;
    });
  }

  /**
   * 记下一个坑位的值，`apply()` 时才写进文档。**名字不存在就抛** —— 拼错字段名
   * 静默不填，是模板填充里最难发现的一类错。换行写成软换行、`\t` 写成制表位
   */
  set(key: string, value: BindingValue): this {
    const dateValue = value !== null && typeof value === 'object' && typeof value.date === 'string';
    if (typeof value !== 'string' && typeof value !== 'boolean' && !dateValue)
      throw new TypeError(`坑位「${key}」的值必须是字符串、布尔值或日期对象`);
    const targets = this.#spans().filter((s) => s.key === key);
    if (!targets.length) throw new RangeError(`文档里没有坑位：${key}`);
    for (const { span } of targets) {
      const type = span.control.type;
      if (type === 'checkbox') {
        if (typeof value !== 'boolean') throw new TypeError(`坑位「${key}」是 checkbox 控件，需要布尔值`);
      } else if (type === 'date' && dateValue) {
        formatContentControlDate(span.control, (value as DateBindingValue).date);
      } else {
        if (!FILLABLE.has(type)) throw new Error(`坑位「${key}」是 ${type} 控件，不能填写`);
        if (typeof value !== 'string') throw new TypeError(`坑位「${key}」是 ${type} 控件，需要字符串`);
      }
    }
    this.#pending.set(key, dateValue ? Object.freeze({ date: (value as DateBindingValue).date }) : value);
    return this;
  }

  /**
   * 一次记多个。与 `set` 不同，**文档里没有的名字跳过**：数据对象常常带着模板用不上的字段
   * （同一份数据套好几份模板），为此逐个挑字段太啰嗦。返回值里的 `skipped` 就是跳过的那些
   */
  setMany(values: Readonly<Record<string, BindingValue>>): { skipped: string[] } {
    const keys = new Set(this.#spans().map((s) => s.key));
    const skipped: string[] = [];
    for (const [key, value] of Object.entries(values)) {
      if (keys.has(key)) this.set(key, value);
      else skipped.push(key);
    }
    return { skipped };
  }

  /** 还没提交的值 */
  get pending(): ReadonlyMap<string, BindingValue> {
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
    const jobs: [NodeId, BindingValue][] = this.#spans()
      .filter((s) => this.#pending.has(s.key))
      .map((s) => [s.span.control.id, this.#pending.get(s.key) as BindingValue] as [NodeId, BindingValue])
      .reverse();
    const change = this.#host.tx((t) => {
      for (const [id, value] of jobs) {
        if (typeof value === 'boolean') t.setContentControlChecked(id, value);
        else if (typeof value === 'string') t.fillContentControl(id, value);
        else t.setContentControlDate(id, value.date);
      }
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
