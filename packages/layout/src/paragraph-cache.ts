/**
 * 段落几何不含页内 y，可以跨事务复用；分页与表格装配仍每次重算。
 *
 * 缓存项**深冻结后直接共享**，不再每次命中都 `structuredClone`：54 页的文档打一个字，
 * 排版 22ms 里有一半花在克隆上（段落本身一行都没重排）。冻结同样保证「调用方改不动缓存」，
 * 只是违规写入从悄悄污染下一帧变成当场抛 TypeError（ES 模块是严格模式）——
 * 排版下游一律展开复制（`joinParagraphFrames`、表格拆行），从来不原地改段落几何。
 */
import type { TextMeasurer } from '@uw/fonts';
import type { ResolvedParagraph } from '@uw/model';
import type { LayoutParagraphOptions } from './paragraph.ts';
import type { ParagraphLayout } from './types.ts';

export class ParagraphLayoutCache {
  readonly #capacity: number;
  readonly #entries = new Map<string, ParagraphLayout>();
  /**
   * 冻结段落 → 它的序列化串。级联备忘（`@uw/model` 的 `ResolveCache`）让没改的段落跨事务
   * 是同一个冻结对象，于是键里最贵的那一截（每段约 3.6KB 的级联属性）只序列化一次。
   * **只认冻结的**：可写的段落可能被原地改过，按身份记会拿旧串命中旧结果
   */
  #paragraphKeys = new WeakMap<ResolvedParagraph, string>();
  #measurer: TextMeasurer | undefined;
  #revision: number | undefined;

  constructor(capacity = 2048) {
    if (!Number.isSafeInteger(capacity) || capacity < 1) throw new RangeError('段落缓存容量必须是正整数');
    this.#capacity = capacity;
  }

  get size(): number {
    return this.#entries.size;
  }

  clear(): void {
    this.#entries.clear();
    this.#paragraphKeys = new WeakMap();
    this.#measurer = undefined;
    this.#revision = undefined;
  }

  /** 由 layoutParagraph 调用；无版本的度量器可能原地变化，保守地跳过缓存。 */
  getOrCreate(
    paragraph: ResolvedParagraph,
    options: LayoutParagraphOptions,
    compute: () => ParagraphLayout,
  ): ParagraphLayout {
    const revision = options.measurer.revision;
    if (this.#measurer !== options.measurer || this.#revision !== revision) {
      this.clear();
      this.#measurer = options.measurer;
      this.#revision = revision;
    }
    if (revision === undefined) return compute();
    // 保留完整序列化值避免哈希碰撞。域表只取本段的 run，其他页码变化不使正文失效。
    const key = `${this.#paragraphKey(paragraph)}\u0000${JSON.stringify([
      options.contentWidth,
      options.settings,
      options.docGrid,
      options.defaultFont,
      options.objectRules,
      options.scriptRules,
      options.widthRules,
      paragraph.runs.map((run) => options.fieldValues?.get(run.id)),
    ])}`;
    const hit = this.#entries.get(key);
    if (hit !== undefined) {
      this.#entries.delete(key);
      this.#entries.set(key, hit);
      return hit;
    }
    // 未命中也冻结：同一段落的结果是否可写不该取决于它是不是第一次排
    const result = deepFreeze(compute());
    this.#entries.set(key, result);
    if (this.#entries.size > this.#capacity) {
      const oldest = this.#entries.keys().next();
      if (!oldest.done) this.#entries.delete(oldest.value);
    }
    return result;
  }

  #paragraphKey(paragraph: ResolvedParagraph): string {
    if (!Object.isFrozen(paragraph)) return JSON.stringify(paragraph);
    let key = this.#paragraphKeys.get(paragraph);
    if (key === undefined) {
      key = JSON.stringify(paragraph);
      this.#paragraphKeys.set(paragraph, key);
    }
    return key;
  }
}

function deepFreeze<T>(value: T): T {
  if (typeof value !== 'object' || value === null || Object.isFrozen(value)) return value;
  for (const child of Object.values(value)) deepFreeze(child);
  return Object.freeze(value);
}
