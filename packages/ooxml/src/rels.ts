/**
 * 关系（`.rels`）。
 *
 * OOXML 里几乎所有跨部件引用都是间接的：`document.xml` 里写 `r:embed="rId7"`，
 * 要经 `/word/_rels/document.xml.rels` 才知道 rId7 是哪张图。所以关系表不是元数据，
 * 是解析主链路上的必经一环。
 */

import { resolveTarget } from './part-names.ts';
import type { XmlDocument } from './xml.ts';
import { attr, children } from './xml.ts';

export interface Relationship {
  id: string;
  /** 完整的类型 URI，见下面的 `RelType` */
  type: string;
  /** Target 属性原文。External 时只有它有意义（是个 URL） */
  rawTarget: string;
  /** 解析成的绝对部件名；`External` 时为 undefined */
  target: string | undefined;
  targetMode: 'Internal' | 'External';
}

/** 按 Id 与按 Type 两种查法都要 O(1)，所以索引建两份 */
export interface Relationships {
  byId(id: string): Relationship | undefined;
  byType(type: string): Relationship[];
  all(): Relationship[];
}

export function parseRelationships(sourcePartName: string, doc: XmlDocument): Relationships {
  const list: Relationship[] = [];
  for (const el of children(doc.root, 'Relationship')) {
    const id = attr(el, 'Id');
    const type = attr(el, 'Type');
    const rawTarget = attr(el, 'Target');
    if (id === undefined || type === undefined || rawTarget === undefined) continue;
    const external = attr(el, 'TargetMode') === 'External';
    list.push({
      id,
      type,
      rawTarget,
      // 外部目标是 URL，拿它去拼部件路径会得到一堆垃圾
      target: external ? undefined : resolveTarget(sourcePartName, rawTarget),
      targetMode: external ? 'External' : 'Internal',
    });
  }
  return createRelationships(list);
}

export function createRelationships(list: readonly Relationship[]): Relationships {
  const byId = new Map<string, Relationship>();
  const byType = new Map<string, Relationship[]>();
  for (const r of list) {
    byId.set(r.id, r);
    const bucket = byType.get(r.type);
    if (bucket) bucket.push(r);
    else byType.set(r.type, [r]);
  }
  return {
    byId: (id) => byId.get(id),
    byType: (type) => byType.get(type)?.slice() ?? [],
    all: () => list.slice(),
  };
}

const OFFICE_DOC = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** 主链路上用得到的关系类型。Phase 1 只需要前几个，后面按阶段加 */
export const RelType = {
  /** 包级关系：指向主文档部件，是所有解析的入口 */
  OFFICE_DOCUMENT: `${OFFICE_DOC}/officeDocument`,
  STYLES: `${OFFICE_DOC}/styles`,
  NUMBERING: `${OFFICE_DOC}/numbering`,
  SETTINGS: `${OFFICE_DOC}/settings`,
  THEME: `${OFFICE_DOC}/theme`,
  FONT_TABLE: `${OFFICE_DOC}/fontTable`,
  WEB_SETTINGS: `${OFFICE_DOC}/webSettings`,
  HEADER: `${OFFICE_DOC}/header`,
  FOOTER: `${OFFICE_DOC}/footer`,
  FOOTNOTES: `${OFFICE_DOC}/footnotes`,
  ENDNOTES: `${OFFICE_DOC}/endnotes`,
  IMAGE: `${OFFICE_DOC}/image`,
  HYPERLINK: `${OFFICE_DOC}/hyperlink`,
} as const;
