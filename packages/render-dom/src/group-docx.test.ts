/**
 * 端到端：一份**带流程图组合的 docx** 走完 解包 → 解析（成员摊平、框里的字摊进旁表）→ 排版 → 画。
 *
 * 最容易断的三环：组合外面那层 `mc:AlternateContent`（Choice 是 wpg、Fallback 是 VML 组合，
 * 框里的字只能摊一份）、成员里的图（id 带正文的前缀，按正文的关系表解引用）、
 * 成员文本框的字（分页排完才摞，借住在组合后面）。
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

/** 1pt = 12700 EMU */
const P = (pt: number): number => pt * 12700;

function xfrm(x: number, y: number, w: number, h: number, ch = ''): string {
  return `<a:xfrm><a:off x="${P(x)}" y="${P(y)}"/><a:ext cx="${P(w)}" cy="${P(h)}"/>${ch}</a:xfrm>`;
}

const LN = '<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>';

/**
 * 页面坐标 (72pt, 144pt)、300 × 200pt 的组合，子坐标是它的两倍（chExt 600 × 400pt）：
 * 左上一个带字的「开始」框、右下一张图、中间一条带箭头的肘形连接符
 */
const GROUP = `<wpg:wgp>
  <wpg:cNvGrpSpPr/>
  <wpg:grpSpPr>${xfrm(0, 0, 300, 200, `<a:chOff x="0" y="0"/><a:chExt cx="${P(600)}" cy="${P(400)}"/>`)}</wpg:grpSpPr>
  <wps:wsp><wps:cNvSpPr/><wps:spPr>${xfrm(0, 0, 200, 100)}<a:prstGeom prst="flowChartTerminator"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="DDEEFF"/></a:solidFill>${LN}</wps:spPr>
    <wps:txbx><w:txbxContent><w:p><w:r><w:t>开始</w:t></w:r></w:p></w:txbxContent></wps:txbx>
    <wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"/></wps:wsp>
  <wps:wsp><wps:cNvCnPr/><wps:spPr>${xfrm(200, 50, 200, 250)}<a:prstGeom prst="bentConnector3"><a:avLst/></a:prstGeom>
    <a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:tailEnd type="triangle"/></a:ln></wps:spPr><wps:bodyPr/></wps:wsp>
  <pic:pic><pic:nvPicPr/><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill>
    <pic:spPr>${xfrm(400, 300, 200, 100)}<a:prstGeom prst="rect"/></pic:spPr></pic:pic>
</wpg:wgp>`;

/** 同一个组合的 VML 兼容写法：框里也有同一段字 —— 只收 Choice 的话这一份不该摊进旁表 */
const FALLBACK = `<w:pict><v:group style="position:absolute;margin-left:72pt;margin-top:144pt;width:300pt;height:200pt;z-index:5;mso-position-horizontal-relative:page;mso-position-vertical-relative:page" coordsize="6000,4000">
  <v:roundrect style="position:absolute;left:0;top:0;width:2000;height:1000">
    <v:textbox><w:txbxContent><w:p><w:r><w:t>开始</w:t></w:r></w:p></w:txbxContent></v:textbox>
  </v:roundrect>
</v:group></w:pict>`;

const ANCHOR = `<w:r><mc:AlternateContent><mc:Choice Requires="wpg"><w:drawing><wp:anchor behindDoc="0" relativeHeight="5">
  <wp:positionH relativeFrom="page"><wp:posOffset>${P(72)}</wp:posOffset></wp:positionH>
  <wp:positionV relativeFrom="page"><wp:posOffset>${P(144)}</wp:posOffset></wp:positionV>
  <wp:extent cx="${P(300)}" cy="${P(200)}"/><wp:wrapNone/><wp:docPr id="1" name="组合 1"/>
  <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">${GROUP}</a:graphicData></a:graphic>
</wp:anchor></w:drawing></mc:Choice><mc:Fallback>${FALLBACK}</mc:Fallback></mc:AlternateContent></w:r>`;

function makeDocx(): Uint8Array {
  return zipSync({
    '[Content_Types].xml': encoder.encode(
      `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Default Extension="png" ContentType="image/png"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      </Types>`,
    ),
    '_rels/.rels': encoder.encode(
      rels(`<Relationship Id="rId1" Type="${OFFICE_REL}/officeDocument" Target="word/document.xml"/>`),
    ),
    'word/document.xml': encoder.encode(`<w:document><w:body>
      <w:p>${ANCHOR}<w:r><w:t>正文</w:t></w:r></w:p>
      <w:sectPr>
        <w:pgSz w:w="11906" w:h="16838"/>
        <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
      </w:sectPr>
    </w:body></w:document>`),
    'word/_rels/document.xml.rels': encoder.encode(
      rels(`<Relationship Id="rId1" Type="${OFFICE_REL}/image" Target="media/chart.png"/>`),
    ),
    'word/media/chart.png': PNG,
  });
}

function collect(node: RElement, pred: (n: RElement) => boolean, out: RElement[] = []): RElement[] {
  if (pred(node)) out.push(node);
  for (const c of node.children) collect(c, pred, out);
  return out;
}

const textOf = (n: RElement): string => n.text ?? n.children.map(textOf).join('');
const hasClass = (n: RElement, cls: string): boolean => (n.attrs.class ?? '').split(' ').includes(cls);

describe('带流程图组合的 docx 端到端', () => {
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
    textBoxes: doc.textBoxes,
  });
  const root = buildDocument(layout, { imageHref: imageHrefResolver(doc.images) });
  const floats = layout.pages[0]?.floats ?? [];

  it('框里的字只摊一份（Fallback 的 VML 组合不收）', () => {
    expect(Object.keys(doc.textBoxes)).toEqual(['tb0']);
  });

  it('组合落在纸坐标上，成员的字借住在它后面，框是成员缩一半之后的', () => {
    expect(floats).toHaveLength(2);
    const [group, member] = floats;
    expect(group).toMatchObject({ x: 1440, y: 2880, width: 6000, height: 4000 });
    expect(group?.group).toHaveLength(3);
    // 成员 (0, 0, 200, 100)pt 缩一半 → 100 × 50pt
    expect(member).toMatchObject({ x: 1440, y: 2880, width: 2000, height: 1000, inline: true });
  });

  it('画出来：终止符框、肘形连接符带箭头、成员图片拿到正文的字节，字在框里', () => {
    const [g] = collect(root, (n) => hasClass(n, 'uw-group'));
    expect(g).toBeDefined();
    const paths = collect(g as RElement, (n) => n.tag === 'path');
    // 终止符（两头半椭圆）、连接符的线、箭头
    expect(paths.some((p) => (p.attrs.d ?? '').includes('A'))).toBe(true);
    expect(paths.some((p) => p.attrs.fill === 'none' && (p.attrs.d ?? '').split('L').length === 4)).toBe(
      true,
    );
    expect(paths.some((p) => p.attrs.fill === '#000000')).toBe(true);
    const [img] = collect(g as RElement, (n) => n.tag === 'image');
    expect(img?.attrs).toMatchObject({ x: '272', y: '294', width: '100', height: '50' });
    expect(img?.attrs.href).toMatch(/^data:image\/png;base64,/);
    const [box] = collect(root, (n) => hasClass(n, 'uw-text-box'));
    expect(textOf(box as RElement)).toContain('开始');
  });

  it('一条诊断都没有', () => {
    expect(sink.list().filter((d) => d.code === 'image-missing')).toEqual([]);
  });
});
