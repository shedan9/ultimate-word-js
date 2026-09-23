/**
 * 回写 docx 的入口：原包 + 编辑后的正文树 → 新的 docx 字节。
 *
 * **只重写改过的部件**。没编辑过的文档回写出来，每个部件都与原包逐字节相同
 * （`OpcPackage.save` 照搬没提到的条目）—— 这是 Phase 8 DoD「加载 → 不编辑 → 导出，
 * Word 打开无修复提示」最稳的一种满足法：我们根本没碰它。
 *
 * 编辑过的文档目前最多动这几个部件：
 * - `document.xml` —— 正文变了（body-writer.ts，补丁式，未改的段落逐字节不变）
 * - `numbering.xml` —— 新建列表加了定义（numbering-writer.ts，只追加）
 * - `styles.xml` —— 套用文档里没有的内建样式补了定义（styles-writer.ts，只追加）
 * - 新建 numbering.xml / styles.xml 时，外加 `[Content_Types].xml` 与主文档的关系表各登记一条
 *
 * 页眉页脚、设置都不编辑，所以都不写。`docProps/app.xml` 里的字数 / 页数统计
 * 会过期，Word 打开时自己重算，不写是对的 —— 写了也只是我们算的数，不是 Word 的。
 */
import type { DiagnosticSink } from '@uw/core';
import { createDiagnosticSink } from '@uw/core';
import type { Body, BodySources } from '@uw/model';
import { parseBody, parseNumbering, parseStyles } from '@uw/model';
import type { OpcPackage, XmlDocument } from '@uw/ooxml';
import { ContentType, parseXml, RelType, relsPartNameOf, serializeXml } from '@uw/ooxml';
import { writeBody } from './body-writer.ts';
import { addedNumbering, writeNumbering } from './numbering-writer.ts';
import { addedStyles, writeStyles } from './styles-writer.ts';
import { el, same } from './xml-edit.ts';

export interface SerializeOptions {
  /** 回写期发现的内容问题（找不到图形原文之类）。缺省丢弃 */
  diagnostics?: DiagnosticSink;
}

const encoder = new TextEncoder();
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT_PART = '/[Content_Types].xml';

function bytesOf(doc: XmlDocument): Uint8Array {
  return encoder.encode(serializeXml(doc));
}

/**
 * `body` 必须来自**同一个包**的 `loadDocument(pkg)`（之后可以任意编辑）：
 * 对应关系靠节点 id，别的包解析出来的 id 对不上。
 */
export function serializeDocx(pkg: OpcPackage, body: Body, options: SerializeOptions = {}): Uint8Array {
  const changes = new Map<string, Uint8Array | null>();
  const main = pkg.mainDocumentPartName();
  const original = pkg.xml(main);
  const sources: BodySources = { nodes: new Map(), content: new Map() };
  // 重新解析的诊断在加载时已经报过一遍，这里丢掉
  const originalBody = parseBody(original, createDiagnosticSink(), main, sources);

  if (!same(originalBody.sections, body.sections)) {
    const doc = writeBody({
      original,
      originalBody,
      sources,
      body,
      part: main,
      ...(options.diagnostics === undefined ? {} : { diagnostics: options.diagnostics }),
    });
    changes.set(main, bytesOf(doc));
  }

  // 新建部件要登记内容类型与关系；两个部件都新建时两次登记改的是同一份清单，所以先攒着最后再写
  const pending = new Map<string, XmlDocument>();
  const numberingPart = pkg.partNameByRelType(RelType.NUMBERING);
  const baseNumbering = numberingPart === undefined ? undefined : pkg.xml(numberingPart);
  const added = addedNumbering(parseNumbering(baseNumbering, createDiagnosticSink()), body.numbering);
  if (added !== undefined) {
    const part = numberingPart ?? registerPart(pkg, main, pending, NUMBERING_PART);
    changes.set(part, bytesOf(writeNumbering(baseNumbering, added)));
  }

  const stylesPart = pkg.partNameByRelType(RelType.STYLES);
  const baseStyles = stylesPart === undefined ? undefined : pkg.xml(stylesPart);
  const newStyles = addedStyles(parseStyles(baseStyles, createDiagnosticSink()), body.styles);
  if (newStyles !== undefined) {
    const part = stylesPart ?? registerPart(pkg, main, pending, STYLES_PART);
    changes.set(part, bytesOf(writeStyles(baseStyles, newStyles)));
  }
  for (const [part, doc] of pending) changes.set(part, bytesOf(doc));
  return pkg.save(changes);
}

interface NewPart {
  /** 不带扩展名的惯用文件名（`numbering`），与主文档同目录 */
  base: string;
  contentType: string;
  relType: string;
}
const NUMBERING_PART: NewPart = {
  base: 'numbering',
  contentType: ContentType.NUMBERING,
  relType: RelType.NUMBERING,
};
const STYLES_PART: NewPart = { base: 'styles', contentType: ContentType.STYLES, relType: RelType.STYLES };

/**
 * 原包没有这个部件：新建一个，登记内容类型与关系。部件名取 Word 的惯例（`/word/numbering.xml`，
 * 主文档同目录），被占了就加序号 —— 包里可能留着一份没人引用的旧文件。
 * 清单与关系表改在 `pending` 里那一份上：同一次回写可能新建两个部件，各自从原包读会互相覆盖。
 */
function registerPart(
  pkg: OpcPackage,
  main: string,
  pending: Map<string, XmlDocument>,
  spec: NewPart,
): string {
  const dir = main.slice(0, main.lastIndexOf('/') + 1);
  let name = `${dir}${spec.base}.xml`;
  for (let n = 1; pkg.has(name); n++) name = `${dir}${spec.base}${n}.xml`;

  const ct = pending.get(CT_PART) ?? pkg.xml(CT_PART);
  const override = el('Override', { PartName: name, ContentType: spec.contentType });
  pending.set(CT_PART, { ...ct, root: { ...ct.root, children: [...ct.root.children, override] } });

  const relsPart = relsPartNameOf(main);
  const rels =
    pending.get(relsPart) ??
    (pkg.has(relsPart)
      ? pkg.xml(relsPart)
      : parseXml(
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${RELS_NS}"/>`,
        ));
  const taken = new Set(
    rels.root.children.flatMap((c) => (c.kind === 'element' && c.attrs.Id !== undefined ? [c.attrs.Id] : [])),
  );
  let id = 1;
  while (taken.has(`rId${id}`)) id++;
  const rel = el('Relationship', { Id: `rId${id}`, Type: spec.relType, Target: name.slice(dir.length) });
  pending.set(relsPart, { ...rels, root: { ...rels.root, children: [...rels.root.children, rel] } });
  return name;
}
