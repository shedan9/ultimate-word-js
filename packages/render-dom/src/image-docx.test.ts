/**
 * 端到端：一份**带图的 docx** 走完 解包 → 解析 → 收图片字节 → 排版 → 画。
 *
 * 为什么是合成的包而不是 fixture：`apps/fidelity/fixtures` 里的每一份 docx 都配着
 * Word 导出的真值，而图片的**几何**至今没有真值（图的底边到底坐在基线上还是差一点，
 * 见 `@uw/layout` 的 `OBJECT_SITS_ON_BASELINE`）。这里要证的不是几何精度，而是
 * **那条链一个环都没断**：`r:embed` → 关系表 → media 部件 → data URI → `<image href>`。
 * 中间任何一环写错（尤其是页眉那份 id 前缀），单元测试都照不出来 —— 它们各自都是绿的。
 */
import { createDiagnosticSink } from '@uw/core';
import { createTextMeasurer, FontRegistry } from '@uw/fonts';
import { loadBundledPacks } from '@uw/fonts/node';
import { layoutDocument } from '@uw/layout';
import { fontNameCandidates, loadDocument } from '@uw/model';
import { OpcPackage } from '@uw/ooxml';
import { zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { imageHrefResolver } from './image.ts';
import { buildDocument } from './paint.ts';
import type { RElement } from './tree.ts';

/** 只当字节用，不解码 —— 这里验的是「字节有没有原样走到 href」 */
const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
const PNG_B64 = 'iVBORw0KGgo=';

const encoder = new TextEncoder();

/** 100pt × 50pt 的内嵌图（EMU：1pt = 12700） */
function drawing(relId: string): string {
  return `<w:r><w:drawing><wp:inline>
    <wp:extent cx="1270000" cy="635000"/>
    <wp:docPr id="1" name="图片 1" descr="红头"/>
    <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
      <pic:pic><pic:blipFill><a:blip r:embed="${relId}"/></pic:blipFill></pic:pic>
    </a:graphicData></a:graphic>
  </wp:inline></w:drawing></w:r>`;
}

const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function rels(entries: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${RELS_NS}">${entries}</Relationships>`;
}

/**
 * 一份最小的 docx：正文一张图、页眉一张图，**两边的关系 id 都叫 rId1**。
 *
 * 故意撞车：`ImageRef.id` 少了部件前缀的话，页眉那张会把正文那张顶掉（与页脚页码
 * 画进正文是同一个坑，见 `@uw/model` 的 nodes.ts）。两张图的字节不同，一眼能分出来。
 */
function makeDocx(): Uint8Array {
  const body = `<w:document><w:body>
    <w:p>${drawing('rId1')}<w:r><w:t>甲</w:t></w:r></w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId9"/>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body></w:document>`;

  const header = `<w:hdr><w:p>${drawing('rId1')}</w:p></w:hdr>`;

  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': encoder.encode(
      `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Default Extension="png" ContentType="image/png"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
      </Types>`,
    ),
    '_rels/.rels': encoder.encode(
      rels(`<Relationship Id="rId1" Type="${OFFICE_REL}/officeDocument" Target="word/document.xml"/>`),
    ),
    'word/document.xml': encoder.encode(body),
    'word/_rels/document.xml.rels': encoder.encode(
      rels(
        `<Relationship Id="rId1" Type="${OFFICE_REL}/image" Target="media/body.png"/>` +
          `<Relationship Id="rId9" Type="${OFFICE_REL}/header" Target="header1.xml"/>`,
      ),
    ),
    'word/header1.xml': encoder.encode(header),
    'word/_rels/header1.xml.rels': encoder.encode(
      rels(`<Relationship Id="rId1" Type="${OFFICE_REL}/image" Target="media/header.png"/>`),
    ),
    'word/media/body.png': PNG,
    // 页眉那张多一个字节，于是两份 base64 不同
    'word/media/header.png': new Uint8Array([...PNG, 1]),
  };
  return zipSync(files);
}

function collect(node: RElement, tag: string, out: RElement[] = []): RElement[] {
  if (node.tag === tag) out.push(node);
  for (const c of node.children) collect(c, tag, out);
  return out;
}

describe('带图的 docx 端到端', () => {
  const sink = createDiagnosticSink();
  const doc = loadDocument(OpcPackage.open(makeDocx()), sink);
  const registry = new FontRegistry();
  for (const pack of loadBundledPacks()) registry.registerMetrics(pack);
  const measurer = createTextMeasurer(registry, {
    candidates: (family) => fontNameCandidates(doc.fonts, family),
    diagnostics: sink,
  });
  const layout = layoutDocument(doc.resolved, {
    measurer,
    settings: doc.cascade.settings,
    headerFooters: doc.headerFooters,
  });
  const root = buildDocument(layout, { imageHref: imageHrefResolver(doc.images) });
  const images = collect(root, 'image');

  it('图片字节按「部件前缀 + 关系 id」收，正文与页眉的 rId1 不会互相顶掉', () => {
    expect(Object.keys(doc.images).sort()).toEqual(['rId1', 'rId9:rId1']);
    expect(doc.images.rId1?.part).toBe('/word/media/body.png');
    expect(doc.images['rId9:rId1']?.part).toBe('/word/media/header.png');
    // 内容类型来自 [Content_Types].xml 的 Default 项，不是猜的
    expect(doc.images.rId1?.contentType).toBe('image/png');
  });

  it('两张图都画出来了，各自拿到自己的字节', () => {
    expect(images).toHaveLength(2);
    const hrefs = images.map((i) => i.attrs.href);
    expect(hrefs).toContain(`data:image/png;base64,${PNG_B64}`);
    expect(new Set(hrefs).size).toBe(2);
  });

  it('外框尺寸一路从 wp:extent 走到属性：100 × 50pt', () => {
    for (const img of images) {
      expect(img.attrs.width).toBe('100');
      expect(img.attrs.height).toBe('50');
    }
  });

  it('内嵌图占着文字流：同一行里「甲」排在图的右边', () => {
    const text = collect(root, 'text')[0];
    // 图宽 100pt，文字紧随其后（版心原点在 `<g>` 上，所以这里的 x 是相对版心的）
    expect(Number(text?.attrs.x)).toBeGreaterThanOrEqual(100);
  });

  it('可选文本进了 <title>，没有变成能被复制走的正文', () => {
    expect(collect(root, 'title').map((t) => t.text)).toEqual(['红头', '红头']);
  });

  it('一条诊断都没有 —— 图片这条链上没有「不认识的元素」', () => {
    expect(sink.list().filter((d) => d.code === 'image-missing')).toEqual([]);
  });
});
