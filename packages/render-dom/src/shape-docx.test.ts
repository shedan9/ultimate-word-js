/**
 * 端到端：一份**公文版头**样式的 docx —— 发文机关标志下面一条红色分隔线（直线形状），
 * 走完 解包 → 解析 → 排版 → 画。
 *
 * 各层单测照不出的两环：① 线包在 `mc:AlternateContent` 里，Choice 是 DrawingML、
 * Fallback 是同一条线的 VML —— 只能画一条；② 线的外框是**零高**的，原来渲染层
 * 「零尺寸对象一律不画」的那条规矩会把它整条吃掉，这正是要修的那个洞。
 * 页脚里另放一条老式的 VML `v:line`，看 from / to 折成的外框落在纸上的哪儿。
 */
import { createDiagnosticSink } from '@uw/core';
import { createTextMeasurer, FontRegistry } from '@uw/fonts';
import { loadBundledPacks } from '@uw/fonts/node';
import { layoutDocument } from '@uw/layout';
import { fontNameCandidates, loadDocument } from '@uw/model';
import { OpcPackage } from '@uw/ooxml';
import { zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { buildDocument } from './paint.ts';
import type { RElement } from './tree.ts';

const encoder = new TextEncoder();
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function rels(entries: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${RELS_NS}">${entries}</Relationships>`;
}

/** 版心通栏（441pt）、相对段落往下 7.8pt、红色 2.25pt 的横线；Fallback 是同一条线的 VML */
const RED_RULE = `<w:r><mc:AlternateContent>
  <mc:Choice Requires="wps"><w:drawing><wp:anchor behindDoc="0" relativeHeight="2" allowOverlap="1">
    <wp:simplePos x="0" y="0"/>
    <wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>
    <wp:positionV relativeFrom="paragraph"><wp:posOffset>99060</wp:posOffset></wp:positionV>
    <wp:extent cx="5600700" cy="0"/><wp:effectExtent l="0" t="12700" r="19050" b="31750"/><wp:wrapNone/>
    <wp:docPr id="1" name="直接连接符 1"/>
    <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
      <wps:wsp><wps:cNvCnPr/>
        <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5600700" cy="0"/></a:xfrm>
          <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
          <a:ln w="28575"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></wps:spPr>
        <wps:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef></wps:style>
        <wps:bodyPr/>
      </wps:wsp>
    </a:graphicData></a:graphic>
  </wp:anchor></w:drawing></mc:Choice>
  <mc:Fallback><w:pict><v:line style="position:absolute;z-index:2;mso-position-horizontal-relative:margin"
    from="0,7.8pt" to="441pt,7.8pt" strokecolor="red" strokeweight="2.25pt"/></w:pict></mc:Fallback>
</mc:AlternateContent></w:r>`;

/** 页脚：一条相对纸 (90pt, 760pt) 起、往右上斜到 (300pt, 750pt) 的 VML 线 */
const FOOTER_LINE = `<w:r><w:pict><v:line style="position:absolute;z-index:1;mso-position-horizontal-relative:page;mso-position-vertical-relative:page"
  from="90pt,760pt" to="300pt,750pt" strokecolor="#0000ff" strokeweight="1pt"><v:stroke dashstyle="dash"/></v:line></w:pict></w:r>`;

function makeDocx(): Uint8Array {
  const body = `<w:document><w:body>
    <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>某某市人民政府文件</w:t></w:r>${RED_RULE}</w:p>
    <w:p><w:r><w:t>正文</w:t></w:r></w:p>
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
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
      </Types>`,
    ),
    '_rels/.rels': encoder.encode(
      rels(`<Relationship Id="rId1" Type="${OFFICE_REL}/officeDocument" Target="word/document.xml"/>`),
    ),
    'word/document.xml': encoder.encode(body),
    'word/_rels/document.xml.rels': encoder.encode(
      rels(`<Relationship Id="rId9" Type="${OFFICE_REL}/footer" Target="footer1.xml"/>`),
    ),
    'word/footer1.xml': encoder.encode(`<w:ftr><w:p>${FOOTER_LINE}<w:r><w:t>第1页</w:t></w:r></w:p></w:ftr>`),
  });
}

function collect(node: RElement, pred: (n: RElement) => boolean, out: RElement[] = []): RElement[] {
  if (pred(node)) out.push(node);
  for (const c of node.children) collect(c, pred, out);
  return out;
}

describe('带分隔线的公文版头 docx 端到端', () => {
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
  const root = buildDocument(layout);
  const page = layout.pages[0];
  const lines = collect(root, (n) => n.tag === 'line' && (n.attrs.class ?? '').includes('uw-shape'));

  it('红线只画一条（Choice 与 Fallback 不叠画），通栏、零高、红色 2.25pt', () => {
    const rule = (page?.floats ?? []).find((f) => !f.runId.includes(':'));
    expect(rule).toMatchObject({ x: 1440, width: 8820, height: 0 });
    const red = lines.filter((l) => l.attrs.stroke === '#ff0000');
    expect(red).toHaveLength(1);
    expect(red[0]?.attrs).toMatchObject({ x1: '72', x2: '513', 'stroke-width': '2.25', fill: 'none' });
    expect(red[0]?.attrs.y1).toBe(red[0]?.attrs.y2);
  });

  it('页脚的 VML 线：外框取两个端点围成的矩形，往右上斜是竖直翻转，虚线按线宽', () => {
    const blue = lines.find((l) => l.attrs.stroke === '#0000ff');
    expect(blue?.attrs).toMatchObject({
      x1: '90',
      y1: '750',
      x2: '300',
      y2: '760',
      'stroke-dasharray': '4 3',
    });
    expect(blue?.attrs.transform).toBe('translate(0 1510) scale(1 -1)');
  });

  it('一个虚线占位框都没有', () => {
    expect(collect(root, (n) => (n.attrs.class ?? '').includes('object-placeholder'))).toHaveLength(0);
  });
});
