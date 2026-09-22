/** 段落几何不含页内 y，可以跨事务复用；分页与表格装配仍每次重算。 */
import type { TextMeasurer } from '@uw/fonts';
import type { ResolvedParagraph } from '@uw/model';
import type { LayoutParagraphOptions } from './paragraph.ts';
import type { ParagraphLayout } from './types.ts';

export class ParagraphLayoutCache {
  readonly #capacity: number;
  readonly #entries = new Map<string, ParagraphLayout>();
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
    const key = JSON.stringify([
      paragraph,
      options.contentWidth,
      options.settings,
      options.docGrid,
      options.defaultFont,
      options.objectRules,
      options.scriptRules,
      options.widthRules,
      paragraph.runs.map((run) => options.fieldValues?.get(run.id)),
    ]);
    const hit = this.#entries.get(key);
    if (hit !== undefined) {
      this.#entries.delete(key);
      this.#entries.set(key, hit);
      // 公开布局允许调用方修改；不能让上一帧或缓存项与下一帧共享可变片段。
      return structuredClone(hit);
    }
    const result = compute();
    this.#entries.set(key, structuredClone(result));
    if (this.#entries.size > this.#capacity) {
      const oldest = this.#entries.keys().next();
      if (!oldest.done) this.#entries.delete(oldest.value);
    }
    return result;
  }
}
