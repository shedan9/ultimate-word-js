/**
 * 端到端：一份**带文本框的 docx** 走完 解包 → 解析（内容摊进旁表）→ 收图片字节 → 排版 → 画。
 *
 * 与 image-docx.test.ts 同一个理由：各层单测都是绿的，照不出「环断在哪儿」——
 * 这里最容易断的两环是**框里的图**（id 带着文本框自己的前缀，却要按所在部件的关系表解引用）
 * 与**页脚里的 VML 框**（WPS 从 PDF 转出来的文件就是这种，原来当内嵌挤在页脚那一行里）。
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

const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
const encoder = new TextEncoder();
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function rels(entries: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${RELS_NS}">${entries}</Relationships>`;
}

/** 正文：一个页面坐标 (72pt, 144pt)、200×100pt 的文本框，带黑边，框里一段字 + 一张小图 */
const BODY_BOX = `<w:r><w:drawing><wp:anchor behindDoc="0" relativeHeight="5">
  <wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH>
  <wp:positionV relativeFrom="page"><wp:posOffset>1828800</wp:posOffset></wp:positionV>
  <wp:extent cx="2540000" cy="1270000"/><wp:wrapNone/><wp:docPr id="1" name="文本框 1"/>
  <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
    <wps:wsp>
      <wps:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>
      <wps:txbx><w:txbxContent><w:p><w:r><w:t>框内文字</w:t></w:r>
        <w:r><w:drawing><wp:inline><wp:extent cx="127000" cy="127000"/>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill></pic:pic>
          </a:graphicData></a:graphic></wp:inline></w:drawing></w:r>
      </w:p></w:txbxContent></wps:txbx>
      <wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/>
    </wps:wsp>
  </a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;

/** 页脚：一个绝对定位的 VML 文本框（无填充无边框），相对纸 (400pt, 780pt) */
const FOOTER_BOX = `<w:r><w:pict><v:shape type="#_x0000_t202" filled="f" stroked="f"
  style="position:absolute;margin-left:400pt;margin-top:780pt;width:80pt;height:14pt;z-index:3;mso-position-horizontal-relative:page;mso-position-vertical-relative:page">
  <v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:r><w:t>页脚框</w:t></w:r></w:p></w:txbxContent></v:textbox>
</v:shape></w:pict></w:r>`;

function makeDocx(): Uint8Array {
  const body = `<w:document><w:body>
    <w:p>${BODY_BOX}<w:r><w:t>正文</w:t></w:r></w:p>
    <w:sectPr>
      <w:footerReference w:type="default" r:id="rId9"/>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body></w:document>`;
  return zipSync({
    '[Content_Types].xml': encoder.encode(
      `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Default Extension="png" ContentType="image/png"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
      </Types>`,
    ),
    '_rels/.rels': encoder.encode(
      rels(`<Relationship Id="rId1" Type="${OFFICE_REL}/officeDocument" Target="word/document.xml"/>`),
    ),
    'word/document.xml': encoder.encode(body),
    'word/_rels/document.xml.rels': encoder.encode(
      rels(
        `<Relationship Id="rId1" Type="${OFFICE_REL}/image" Target="media/box.png"/>` +
          `<Relationship Id="rId9" Type="${OFFICE_REL}/footer" Target="footer1.xml"/>`,
      ),
    ),
    'word/footer1.xml': encoder.encode(`<w:ftr><w:p>${FOOTER_BOX}<w:r><w:t>第1页</w:t></w:r></w:p></w:ftr>`),
    'word/media/box.png': PNG,
  });
}

function collect(node: RElement, pred: (n: RElement) => boolean, out: RElement[] = []): RElement[] {
  if (pred(node)) out.push(node);
  for (const c of node.children) collect(c, pred, out);
  return out;
}

const textOf = (n: RElement): string => n.text ?? n.children.map(textOf).join('');

describe('带文本框的 docx 端到端', () => {
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
    textBoxes: doc.textBoxes,
  });
  const root = buildDocument(layout, { imageHref: imageHrefResolver(doc.images) });
  const page = layout.pages[0];

  it('两个框的内容各摊一份，id 带部件前缀；框里的图按所在部件的关系表解引用', () => {
    expect(Object.keys(doc.textBoxes).sort()).toEqual(['rId9:tb0', 'tb0']);
    expect(doc.images['tb0:rId1']?.part).toBe('/word/media/box.png');
  });

  it('正文框落在纸坐标上，内容从框的左上角起摞；页脚的 VML 框是浮动的，不挤页脚那一行', () => {
    const [body] = (page?.floats ?? []).filter((f) => f.runId === 'r0');
    expect(body).toMatchObject({ x: 1440, y: 2880, width: 4000, height: 2000 });
    expect(body?.textBox).toMatchObject({ x: 1440, y: 2880, width: 4000 });

    const footerBox = (page?.floats ?? []).find((f) => f.runId.startsWith('rId9:'));
    expect(footerBox).toMatchObject({ x: 8000, y: 15600, width: 1600, height: 280 });
    // 页脚那一行只剩「第1页」：框不占行宽
    const footerLine = page?.footer?.blocks[0];
    expect(footerLine?.kind === 'paragraph' && footerLine.lines).toHaveLength(1);
  });

  it('画出来：框里的字、框里的图、黑边都在，一个虚线占位框都没有', () => {
    const texts = collect(root, (n) => n.attrs.class === 'uw-text-box').map(textOf);
    expect(texts.some((t) => t.includes('框内文字'))).toBe(true);
    expect(texts.some((t) => t.includes('页脚框'))).toBe(true);
    expect(collect(root, (n) => n.tag === 'image')).toHaveLength(1);
    expect(collect(root, (n) => (n.attrs.class ?? '').includes('uw-shape'))).toHaveLength(1);
    expect(collect(root, (n) => (n.attrs.class ?? '').includes('object-placeholder'))).toHaveLength(0);
  });
});
