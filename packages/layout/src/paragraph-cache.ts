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
  /**
   * 两级：段落串 → 环境串 → 结果。**不拼成一个长键**：Map 查表要给键算哈希，
   * 现拼的串每次都是新对象、要把约 9KB（中文是双字节串）拉平重算一遍 ——
   * 实测每次 9.4µs，54 页的文档每次按键光查表就 4.7ms。外层直接拿 `#frozenKeys`
   * 里记住的那个串对象当键，V8 把哈希缓存在串对象上，命中只剩指针比较
   */
  readonly #entries = new Map<string, Map<string, ParagraphLayout>>();
  #size = 0;
  /**
   * 冻结对象 → 它的序列化串。级联备忘（`@uw/model` 的 `ResolveCache`）让没改的段落跨事务
   * 是同一个冻结对象，于是键里最贵的那一截（每段约 3.6KB 的级联属性）只序列化一次；
   * 文档设置（约 765 字节，一趟排版里每段都是同一个对象）同理。
   * **只认冻结的**：可写的对象可能被原地改过，按身份记会拿旧串命中旧结果。
   * 调用方要**深**冻结 —— `Object.isFrozen` 只看最外一层
   */
  #frozenKeys = new WeakMap<object, string>();
  /** 设置串 → 小编号，让内层键保持短（设置几乎不变，这张表只有一两项） */
  #settingsIds = new Map<string, number>();
  #measurer: TextMeasurer | undefined;
  #revision: number | undefined;

  constructor(capacity = 2048) {
    if (!Number.isSafeInteger(capacity) || capacity < 1) throw new RangeError('段落缓存容量必须是正整数');
    this.#capacity = capacity;
  }

  get size(): number {
    return this.#size;
  }

  clear(): void {
    this.#entries.clear();
    this.#size = 0;
    this.#frozenKeys = new WeakMap();
    this.#settingsIds.clear();
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
    const paragraphKey = this.#frozenKey(paragraph);
    const settingsKey = this.#frozenKey(options.settings);
    let settingsId = this.#settingsIds.get(settingsKey);
    if (settingsId === undefined) {
      settingsId = this.#settingsIds.size;
      this.#settingsIds.set(settingsKey, settingsId);
    }
    const envKey = JSON.stringify([
      settingsId,
      options.contentWidth,
      options.docGrid,
      options.defaultFont,
      options.objectRules,
      options.scriptRules,
      options.widthRules,
      // 同一个 run 不会既是域结果又带脚注号（域结果整个换掉 run），合成一列省一次序列化
      paragraph.runs.map((run) => options.fieldValues?.get(run.id) ?? options.noteLabels?.get(run.id)),
    ]);
    const bucket = this.#entries.get(paragraphKey);
    const hit = bucket?.get(envKey);
    if (bucket !== undefined && hit !== undefined) {
      // LRU 以段落为单位挪到队尾：同一段落的几种环境（不同宽度）一起冷一起热
      this.#entries.delete(paragraphKey);
      this.#entries.set(paragraphKey, bucket);
      return hit;
    }
    // 未命中也冻结：同一段落的结果是否可写不该取决于它是不是第一次排
    const result = deepFreeze(compute());
    if (bucket === undefined) this.#entries.set(paragraphKey, new Map([[envKey, result]]));
    else {
      bucket.set(envKey, result);
      this.#entries.delete(paragraphKey);
      this.#entries.set(paragraphKey, bucket);
    }
    this.#size++;
    this.#evict();
    return result;
  }

  /** 超容量时从最冷的段落开始，一项一项淘汰（容量数的是结果项，不是段落） */
  #evict(): void {
    while (this.#size > this.#capacity) {
      const oldest = this.#entries.entries().next();
      if (oldest.done) return;
      const [paragraphKey, bucket] = oldest.value;
      const first = bucket.keys().next();
      if (!first.done) bucket.delete(first.value);
      if (bucket.size === 0) this.#entries.delete(paragraphKey);
      this.#size--;
    }
  }

  #frozenKey(value: object): string {
    if (!Object.isFrozen(value)) return JSON.stringify(value);
    let key = this.#frozenKeys.get(value);
    if (key === undefined) {
      key = JSON.stringify(value);
      this.#frozenKeys.set(value, key);
    }
    return key;
  }
}

function deepFreeze<T>(value: T): T {
  if (typeof value !== 'object' || value === null || Object.isFrozen(value)) return value;
  for (const child of Object.values(value)) deepFreeze(child);
  return Object.freeze(value);
}
