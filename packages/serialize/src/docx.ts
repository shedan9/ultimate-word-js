/**
 * 回写 docx 的入口：原包 + 编辑后的正文树 → 新的 docx 字节。
 *
 * **只重写改过的部件**。没编辑过的文档回写出来，每个部件都与原包逐字节相同
 * （`OpcPackage.save` 照搬没提到的条目）—— 这是 Phase 8 DoD「加载 → 不编辑 → 导出，
 * Word 打开无修复提示」最稳的一种满足法：我们根本没碰它。
 *
 * 编辑过的文档目前最多动三个部件：
 * - `document.xml` —— 正文变了（body-writer.ts，补丁式，未改的段落逐字节不变）
 * - `numbering.xml` —— 新建列表加了定义（numbering-writer.ts，只追加）
 * - 新建 numbering.xml 时，外加 `[Content_Types].xml` 与主文档的关系表各登记一条
 *
 * 页眉页脚、样式表、设置都不编辑，所以都不写。`docProps/app.xml` 里的字数 / 页数统计
 * 会过期，Word 打开时自己重算，不写是对的 —— 写了也只是我们算的数，不是 Word 的。
 */
import type { DiagnosticSink } from '@uw/core';
import { createDiagnosticSink } from '@uw/core';
import type { Body, BodySources } from '@uw/model';
import { parseBody, parseNumbering } from '@uw/model';
import type { OpcPackage, XmlDocument } from '@uw/ooxml';
import { ContentType, parseXml, RelType, relsPartNameOf, serializeXml } from '@uw/ooxml';
import { writeBody } from './body-writer.ts';
import { addedNumbering, writeNumbering } from './numbering-writer.ts';
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

  const numberingPart = pkg.partNameByRelType(RelType.NUMBERING);
  const baseNumbering = numberingPart === undefined ? undefined : pkg.xml(numberingPart);
  const added = addedNumbering(parseNumbering(baseNumbering, createDiagnosticSink()), body.numbering);
  if (added !== undefined) {
    const part = numberingPart ?? registerNumberingPart(pkg, main, changes);
    changes.set(part, bytesOf(writeNumbering(baseNumbering, added)));
  }
  return pkg.save(changes);
}

/**
 * 原包没有 numbering.xml：新建一个，登记内容类型与关系。部件名取 Word 的惯例
 * `/word/numbering.xml`（主文档同目录），被占了就加序号 —— 包里可能留着一份没人引用的旧文件。
 */
function registerNumberingPart(
  pkg: OpcPackage,
  main: string,
  changes: Map<string, Uint8Array | null>,
): string {
  const dir = main.slice(0, main.lastIndexOf('/') + 1);
  let name = `${dir}numbering.xml`;
  for (let n = 1; pkg.has(name); n++) name = `${dir}numbering${n}.xml`;

  const ct = pkg.xml(CT_PART);
  const override = el('Override', { PartName: name, ContentType: ContentType.NUMBERING });
  changes.set(CT_PART, bytesOf({ ...ct, root: { ...ct.root, children: [...ct.root.children, override] } }));

  const relsPart = relsPartNameOf(main);
  const rels = pkg.has(relsPart)
    ? pkg.xml(relsPart)
    : parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${RELS_NS}"/>`);
  const taken = new Set(
    pkg
      .rels(main)
      .all()
      .map((r) => r.id),
  );
  let id = 1;
  while (taken.has(`rId${id}`)) id++;
  const rel = el('Relationship', { Id: `rId${id}`, Type: RelType.NUMBERING, Target: name.slice(dir.length) });
  changes.set(relsPart, bytesOf({ ...rels, root: { ...rels.root, children: [...rels.root.children, rel] } }));
  return name;
}
