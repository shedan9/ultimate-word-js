/**
 * `[Content_Types].xml` —— OPC 包里唯一一个「不是部件的部件」。
 *
 * 查找规则是两层：先看 `Override`（按完整部件名精确匹配），没有再看 `Default`（按扩展名）。
 * 顺序不能反 —— docx 里 `xml` 扩展名的 Default 通常是 `application/xml`，
 * 而 `/word/document.xml` 靠 Override 才拿到 WordprocessingML 的类型。
 */

import { extensionOf } from './part-names.ts';
import type { XmlDocument } from './xml.ts';
import { attr, children } from './xml.ts';

export interface ContentTypes {
  /** 查不到返回 undefined —— 这是内容问题不是结构错误，交给上层记诊断 */
  lookup(partName: string): string | undefined;
}

export function parseContentTypes(doc: XmlDocument): ContentTypes {
  const byExtension = new Map<string, string>();
  const byPartName = new Map<string, string>();

  // `[Content_Types].xml` 用默认命名空间，元素名不带前缀
  for (const el of children(doc.root)) {
    const type = attr(el, 'ContentType');
    if (type === undefined) continue;
    if (el.name === 'Default') {
      const ext = attr(el, 'Extension');
      if (ext !== undefined) byExtension.set(ext.toLowerCase(), type);
    } else if (el.name === 'Override') {
      const part = attr(el, 'PartName');
      // 规范说部件名比较不区分大小写，统一小写存，查的时候也小写
      if (part !== undefined) byPartName.set(part.toLowerCase(), type);
    }
  }

  return {
    lookup(partName) {
      return byPartName.get(partName.toLowerCase()) ?? byExtension.get(extensionOf(partName));
    },
  };
}

/** 用得上的几个内容类型。判断部件种类时比猜文件名可靠 */
export const ContentType = {
  MAIN_DOCUMENT: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  STYLES: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
  NUMBERING: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
  SETTINGS: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
  FONT_TABLE: 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml',
  THEME: 'application/vnd.openxmlformats-officedocument.theme+xml',
} as const;
