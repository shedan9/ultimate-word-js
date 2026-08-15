/**
 * OPC 包 —— `@uw/ooxml` 的门面，也是整条流水线的第一个阶段产物。
 *
 * 职责边界要严：它只回答「这个包里有哪些部件、每个部件的字节 / XML 树 / 关系是什么」，
 * **不认识任何 WordprocessingML 语义**。`w:p` 是什么、样式怎么级联，全是 `@uw/model` 的事。
 * 这条线划清楚，Phase 8 回写 docx 时才能原样复用同一套部件与关系。
 *
 * 关于「为什么这里可以是类」：原则 1.1 要求的是**跨阶段边界的数据**可结构化克隆，
 * 而 ooxml 与 model 同在 Worker 一侧（架构 §9），`OpcPackage` 不过边界。
 * 它持有字节和惰性 XML 缓存，用类是合适的；真正要过边界的 `LayoutInput` / `LayoutResult`
 * 才必须是裸对象。
 */
import { UwError, UwErrorCode } from '@uw/core';
import type { ContentTypes } from './content-types.ts';
import { parseContentTypes } from './content-types.ts';
import { relsPartNameOf, toPartName, toZipEntryName } from './part-names.ts';
import type { Relationship, Relationships } from './rels.ts';
import { createRelationships, parseRelationships, RelType } from './rels.ts';
import type { XmlDocument } from './xml.ts';
import { parseXml } from './xml.ts';
import type { ZipEntries } from './zip.ts';
import { unzip } from './zip.ts';

/** 内容类型表自己不是部件，规范里就这么写死的一个名字 */
const CONTENT_TYPES_PART = '/[Content_Types].xml';

export interface OpcPart {
  /** 绝对部件名，带前导斜杠 */
  name: string;
  /** 查不到内容类型时为 undefined —— 不影响读取，只影响按类型找部件 */
  contentType: string | undefined;
  bytes: Uint8Array;
}

const decoder = new TextDecoder('utf-8');
/** 关系为空的部件很多（`.rels` 文件根本不存在），共用一个空表省得反复建 */
const EMPTY_RELS = createRelationships([]);

export class OpcPackage {
  readonly #entries: ZipEntries;
  readonly #contentTypes: ContentTypes;
  /** 惰性缓存：同一个部件的 XML 树只解析一次。styles.xml 动辄 40 KB，重复解析很贵 */
  readonly #xmlCache = new Map<string, XmlDocument>();
  readonly #relsCache = new Map<string, Relationships>();

  private constructor(entries: ZipEntries, contentTypes: ContentTypes) {
    this.#entries = entries;
    this.#contentTypes = contentTypes;
  }

  /**
   * 解包。
   *
   * 这里抛异常是对的（架构 §10）：不是 zip、没有内容类型表 —— 后面一步都走不下去，
   * 没有「渲染其余部分」这个选项。内容层面的问题（某个部件缺失、某个元素不认识）
   * 才走诊断。
   */
  static open(data: Uint8Array | ArrayBuffer): OpcPackage {
    const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
    const entries = unzip(bytes);

    const ctBytes = entries.get(toZipEntryName(CONTENT_TYPES_PART));
    if (ctBytes === undefined) {
      throw new UwError(
        UwErrorCode.NOT_AN_OPC_PACKAGE,
        '这是个 zip，但缺 [Content_Types].xml —— 不是 OOXML 包（.docx / .xlsx / .pptx）',
      );
    }
    const contentTypes = parseContentTypes(parseXml(decoder.decode(ctBytes), CONTENT_TYPES_PART));
    return new OpcPackage(entries, contentTypes);
  }

  /** 全部部件名，不含 `[Content_Types].xml`（它不是部件）。已排序，便于做快照测试 */
  partNames(): string[] {
    const out: string[] = [];
    for (const name of this.#entries.keys()) {
      const partName = toPartName(name);
      if (partName !== CONTENT_TYPES_PART) out.push(partName);
    }
    return out.sort();
  }

  has(partName: string): boolean {
    return this.#entries.has(toZipEntryName(partName));
  }

  part(partName: string): OpcPart | undefined {
    const bytes = this.#entries.get(toZipEntryName(partName));
    if (bytes === undefined) return undefined;
    return { name: partName, contentType: this.#contentTypes.lookup(partName), bytes };
  }

  /** 缺部件是结构性错误还是内容问题，取决于缺的是哪个 —— 所以由调用方决定，这里只抛 */
  requirePart(partName: string): OpcPart {
    const p = this.part(partName);
    if (p === undefined) {
      throw new UwError(UwErrorCode.PART_NOT_FOUND, `包里没有部件 ${partName}`, { part: partName });
    }
    return p;
  }

  text(partName: string): string {
    return decoder.decode(this.requirePart(partName).bytes);
  }

  xml(partName: string): XmlDocument {
    const cached = this.#xmlCache.get(partName);
    if (cached !== undefined) return cached;
    const doc = parseXml(this.text(partName), partName);
    this.#xmlCache.set(partName, doc);
    return doc;
  }

  /** 某部件的关系表。传空串取包级关系（`/_rels/.rels`）。没有 `.rels` 文件就是空表，不是错误 */
  rels(partName = ''): Relationships {
    const cached = this.#relsCache.get(partName);
    if (cached !== undefined) return cached;
    const relsPart = relsPartNameOf(partName);
    const rels = this.has(relsPart) ? parseRelationships(partName, this.xml(relsPart)) : EMPTY_RELS;
    this.#relsCache.set(partName, rels);
    return rels;
  }

  /** 沿关系跳到目标部件名。外部关系（超链接）返回 undefined —— 它不指向包内任何东西 */
  resolveRel(sourcePartName: string, relId: string): string | undefined {
    return this.rels(sourcePartName).byId(relId)?.target;
  }

  /**
   * 主文档部件名（`/word/document.xml`）。
   *
   * 走包级关系而不是硬编码路径：路径是约定俗成不是规范强制，
   * 别的工具生成的 docx 完全可以把主文档放在别处，而 officeDocument 关系一定在。
   */
  mainDocumentPartName(): string {
    const rel = this.rels().byType(RelType.OFFICE_DOCUMENT)[0];
    const target = rel?.target;
    if (target === undefined || !this.has(target)) {
      throw new UwError(
        UwErrorCode.NOT_A_WORD_DOCUMENT,
        '包级关系里找不到 officeDocument 主部件，这可能不是 .docx（也可能是 .xlsx / .pptx）',
      );
    }
    return target;
  }

  /** 主文档直属的关系（样式、编号、设置、主题、图片都挂在这上面） */
  mainDocumentRels(): Relationships {
    return this.rels(this.mainDocumentPartName());
  }

  /**
   * 按关系类型从主文档找一个部件名，找不到返回 undefined。
   * 样式表、编号表这些「可能没有」的部件用它取 —— 没有样式表不是错误，是没定义任何样式。
   */
  partNameByRelType(type: string): string | undefined {
    const target = this.mainDocumentRels().byType(type)[0]?.target;
    return target !== undefined && this.has(target) ? target : undefined;
  }
}

export type { Relationship, Relationships };
