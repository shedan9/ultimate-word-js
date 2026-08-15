/**
 * OPC 容器测试，跑在真实 fixture 上。
 *
 * 用 `gongwen-01.docx`（Word 自己生成的公文）而不是手搓的迷你 zip：这一层的坑
 * 全在真实文件的细节里 —— 相对 Target、`..` 回退、Override 覆盖 Default、
 * 部件名大小写。手搓样本正好把这些都绕开了。
 */
import { readFileSync } from 'node:fs';
import { UwError } from '@uw/core';
import { describe, expect, it } from 'vitest';
import { ContentType } from './content-types.ts';
import { OpcPackage } from './opc.ts';
import { RelType } from './rels.ts';
import { parseXml, serializeXml } from './xml.ts';

const FIXTURE = new URL('../../../apps/fidelity/fixtures/gongwen-01.docx', import.meta.url);
const bytes = new Uint8Array(readFileSync(FIXTURE));
const open = (): OpcPackage => OpcPackage.open(bytes);

describe('OpcPackage.open', () => {
  it('列出全部部件，且不把 [Content_Types].xml 当部件', () => {
    const names = open().partNames();
    expect(names).toContain('/word/document.xml');
    expect(names).toContain('/word/styles.xml');
    expect(names).toContain('/docProps/app.xml');
    expect(names).not.toContain('/[Content_Types].xml');
  });

  it('不是 zip 就抛 NOT_A_ZIP', () => {
    const notZip = new TextEncoder().encode('这显然不是 zip');
    expect(() => OpcPackage.open(notZip)).toThrow(UwError);
    try {
      OpcPackage.open(notZip);
    } catch (e) {
      expect((e as UwError).code).toBe('NOT_A_ZIP');
    }
  });

  it('接受 ArrayBuffer', () => {
    const ab = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
    expect(OpcPackage.open(ab as ArrayBuffer).partNames().length).toBeGreaterThan(0);
  });
});

describe('内容类型', () => {
  it('Override 压过 Default —— document.xml 是主文档类型而不是 application/xml', () => {
    expect(open().part('/word/document.xml')?.contentType).toBe(ContentType.MAIN_DOCUMENT);
    expect(open().part('/word/styles.xml')?.contentType).toBe(ContentType.STYLES);
  });
});

describe('关系', () => {
  it('包级关系定位主文档', () => {
    expect(open().mainDocumentPartName()).toBe('/word/document.xml');
  });

  it('主文档的相对 Target 按源部件所在目录解析', () => {
    const pkg = open();
    const styles = pkg.mainDocumentRels().byType(RelType.STYLES)[0];
    // Target="styles.xml" → /word/styles.xml，而不是 /word/_rels/styles.xml
    expect(styles?.target).toBe('/word/styles.xml');
    expect(pkg.has('/word/styles.xml')).toBe(true);
  });

  it('包级关系里的 docProps 目标落在根目录', () => {
    const rels = open().rels().all();
    const targets = rels.map((r) => r.target);
    expect(targets).toContain('/docProps/app.xml');
    expect(targets).toContain('/docProps/core.xml');
  });

  it('按 rId 跳转', () => {
    const pkg = open();
    const rel = pkg.mainDocumentRels().all()[0];
    expect(rel).toBeDefined();
    if (rel) expect(pkg.resolveRel('/word/document.xml', rel.id)).toBe(rel.target);
  });

  it('没有 .rels 的部件返回空表，不报错', () => {
    expect(open().rels('/word/styles.xml').all()).toEqual([]);
  });

  it('按类型找部件：有的给部件名，没有的给 undefined', () => {
    const pkg = open();
    expect(pkg.partNameByRelType(RelType.STYLES)).toBe('/word/styles.xml');
    expect(pkg.partNameByRelType(RelType.THEME)).toBe('/word/theme/theme1.xml');
    // 这份公文没有编号定义 —— 这不是错误，是「没定义任何列表」
    expect(pkg.partNameByRelType(RelType.NUMBERING)).toBeUndefined();
  });
});

describe('部件读取', () => {
  it('XML 树缓存，同一部件只解析一次', () => {
    const pkg = open();
    expect(pkg.xml('/word/document.xml')).toBe(pkg.xml('/word/document.xml'));
  });

  it('主文档根元素是 w:document', () => {
    expect(open().xml('/word/document.xml').root.name).toBe('w:document');
  });

  it('缺部件抛 PART_NOT_FOUND 且带上部件名', () => {
    try {
      open().requirePart('/word/numbering.xml');
      expect.unreachable('应当抛错');
    } catch (e) {
      expect((e as UwError).code).toBe('PART_NOT_FOUND');
      expect((e as UwError).part).toBe('/word/numbering.xml');
    }
  });
});

describe('round-trip（原则 1.4）', () => {
  const xmlPartsOf = (pkg: OpcPackage): string[] =>
    pkg.partNames().filter((n) => n.endsWith('.xml') || n.endsWith('.rels'));

  it('包里每个 XML 部件序列化再解析，树完全一致', () => {
    const pkg = open();
    const parts = xmlPartsOf(pkg);
    // 这份 fixture 有 11 个部件（含 rels），少于这个数说明 fixture 被换了
    expect(parts.length).toBeGreaterThanOrEqual(10);
    for (const name of parts) {
      const tree = pkg.xml(name);
      expect(parseXml(serializeXml(tree), name), `${name} 走一趟序列化后语义变了`).toEqual(tree);
    }
  });

  it('XML 树可结构化克隆（原则 1.1：能过 Worker 边界、能当 golden file）', () => {
    const pkg = open();
    for (const name of xmlPartsOf(pkg)) {
      const tree = pkg.xml(name);
      expect(structuredClone(tree), `${name} 克隆后不等`).toEqual(tree);
    }
  });
});
