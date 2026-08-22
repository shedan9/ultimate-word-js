/**
 * `w:drawing` / `w:pict` 的解析。
 *
 * 用手写片段，每个用例打一个点。重点全在那些**画错了才看得出来**的地方：
 * 单位换算（EMU / 1/60000 度 / 千分之一百分点 / CSS 长度）、blip 藏得深浅、
 * 以及 id 前缀 —— 前缀错了页眉里的图会画到正文里去。
 */
import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import { collectImages } from './images.ts';
import type { ObjectContent, Paragraph } from './nodes.ts';
import { walkBlocks } from './nodes.ts';
import { parseBody, parseHeaderFooter } from './parse-body.ts';

/** 一张 100×50 pt 的内嵌图（EMU：1pt = 12700） */
const INLINE_PIC = `
<w:r><w:drawing>
  <wp:inline distT="0" distB="0">
    <wp:extent cx="1270000" cy="635000"/>
    <wp:docPr id="1" name="图片 1" descr="红头"/>
    <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
      <pic:pic>
        <pic:blipFill>
          <a:blip r:embed="rId5"/>
          <a:srcRect l="10000" b="25000"/>
        </pic:blipFill>
        <pic:spPr><a:xfrm rot="5400000" flipH="1"><a:ext cx="635000" cy="1270000"/></a:xfrm></pic:spPr>
      </pic:pic>
    </a:graphicData></a:graphic>
  </wp:inline>
</w:drawing></w:r>`;

const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>';

/** 前缀走 `parseHeaderFooter`（正文那条路的前缀恒为空串，见 parse-body.ts 的 `Ctx`） */
function objects(bodyXml: string, prefix = ''): ObjectContent[] {
  const sink = createDiagnosticSink();
  const blocks =
    prefix === ''
      ? parseBody(
          parseXml(`<w:document><w:body>${bodyXml}${SECT}</w:body></w:document>`, 'document.xml'),
          sink,
          'document.xml',
        ).sections.flatMap((s) => s.blocks)
      : parseHeaderFooter(parseXml(`<w:hdr>${bodyXml}</w:hdr>`, 'header1.xml'), sink, 'header1.xml', prefix);
  const out: ObjectContent[] = [];
  for (const b of walkBlocks(blocks)) {
    if (b.kind !== 'paragraph') continue;
    for (const run of (b as Paragraph).runs) {
      for (const c of run.content) if (c.kind === 'object') out.push(c);
    }
  }
  return out;
}

describe('w:drawing', () => {
  it('外框取 wp:extent（EMU → twips），可选文本取 descr', () => {
    const [obj] = objects(`<w:p>${INLINE_PIC}</w:p>`);
    // 1270000 EMU / 635 = 2000 twips = 100pt
    expect(obj?.width).toBe(2000);
    expect(obj?.height).toBe(1000);
    expect(obj?.alt).toBe('红头');
    expect(obj?.graphic).toBe('picture');
    // 内嵌图没有 anchor —— 有没有它就是「参不参与文字流」的判据
    expect(obj?.anchor).toBeUndefined();
  });

  it('外框取的是 extent 而不是 a:ext —— 旋转 90° 的图两者横竖相反', () => {
    const [obj] = objects(`<w:p>${INLINE_PIC}</w:p>`);
    // a:ext 是 635000×1270000（竖），extent 是 1270000×635000（横）。
    // 占位与行高要的是「占多大地方」，也就是转完之后的外接矩形
    expect(obj?.width).toBeGreaterThan(obj?.height ?? 0);
    expect(obj?.image?.rotation).toBe(90); // 5400000 / 60000
    expect(obj?.image?.flipH).toBe(true);
    expect(obj?.image?.flipV).toBeUndefined();
  });

  it('裁剪从千分之一百分点换算成比例，没写的边是 0', () => {
    const [obj] = objects(`<w:p>${INLINE_PIC}</w:p>`);
    expect(obj?.image?.crop).toEqual({ left: 0.1, top: 0, right: 0, bottom: 0.25 });
  });

  it('图片 id 带部件前缀 —— 页眉里的 rId5 与正文里的 rId5 不是同一张图', () => {
    const [body] = objects(`<w:p>${INLINE_PIC}</w:p>`);
    const [header] = objects(`<w:p>${INLINE_PIC}</w:p>`, 'rId7:');
    expect(body?.image?.id).toBe('rId5');
    expect(header?.image?.id).toBe('rId7:rId5');
    // relId 原文两边一样：回写时要用它，不能带前缀
    expect(header?.image?.relId).toBe('rId5');
  });

  it('r:link（外链图）标出来，因为包里没有它的字节', () => {
    const [obj] = objects(
      `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="635" cy="635"/>
        <a:graphic><a:graphicData><pic:pic><pic:blipFill>
          <a:blip r:link="rId9"/>
        </pic:blipFill></pic:pic></a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p>`,
    );
    expect(obj?.image).toEqual({ id: 'rId9', relId: 'rId9', linked: true });
  });

  it('图表 / SmartArt 没有 blip：外框在、图片引用不在 —— 这正是画占位框的判据', () => {
    const [obj] = objects(
      `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1270000" cy="635000"/>
        <wp:docPr id="2" name="图表 2"/>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rId4"/>
        </a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p>`,
    );
    expect(obj?.width).toBe(2000);
    expect(obj?.image).toBeUndefined();
    expect(obj?.graphic).toBe('chart');
    expect(obj?.alt).toBe('图表 2');
  });

  it('blip 藏在 mc:AlternateContent 的 Choice 里也找得到（深搜，不按固定路径）', () => {
    const [obj] = objects(
      `<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps">
        <w:drawing><wp:anchor behindDoc="1" relativeHeight="3">
          <wp:extent cx="1270000" cy="635000"/>
          <wp:wrapNone/>
          <wp:positionH relativeFrom="page"><wp:posOffset>635000</wp:posOffset></wp:positionH>
          <wp:positionV relativeFrom="paragraph"><wp:align>bottom</wp:align></wp:positionV>
          <a:graphic><a:graphicData><wps:wsp><wps:spPr><a:blipFill>
            <a:blip r:embed="rId3"/>
          </a:blipFill></wps:spPr></wps:wsp></a:graphicData></a:graphic>
        </wp:anchor></w:drawing>
      </mc:Choice></mc:AlternateContent></w:r></w:p>`,
    );
    expect(obj?.image?.relId).toBe('rId3');
    expect(obj?.anchor).toEqual({
      wrap: 'none',
      behindDoc: true,
      z: 3,
      h: { relativeFrom: 'page', offset: 1000 },
      v: { relativeFrom: 'paragraph', align: 'bottom' },
      dist: { top: 0, bottom: 0, left: 0, right: 0 },
    });
  });

  it('simplePos="1" 时忽略 positionH/V，改用那一对绝对坐标', () => {
    const [obj] = objects(
      `<w:p><w:r><w:drawing><wp:anchor simplePos="1" distT="12700">
        <wp:simplePos x="635000" y="1270000"/>
        <wp:extent cx="635" cy="635"/>
        <wp:wrapSquare/>
        <wp:positionH relativeFrom="column"><wp:posOffset>99</wp:posOffset></wp:positionH>
      </wp:anchor></w:drawing></w:r></w:p>`,
    );
    expect(obj?.anchor?.h).toEqual({ relativeFrom: 'page', offset: 1000 });
    expect(obj?.anchor?.v).toEqual({ relativeFrom: 'page', offset: 2000 });
    expect(obj?.anchor?.wrap).toBe('square');
    expect(obj?.anchor?.dist.top).toBe(20); // 12700 EMU = 1pt
  });
});

describe('w:pict（VML）', () => {
  it('尺寸从 style 里取，图片引用取 v:imagedata', () => {
    const [obj] = objects(
      `<w:p><w:r><w:pict>
        <v:shape style="position:absolute;width:100pt;height:2cm" alt="印章">
          <v:imagedata r:id="rId8"/>
        </v:shape>
      </w:pict></w:r></w:p>`,
    );
    expect(obj?.objectKind).toBe('picture');
    expect(obj?.width).toBe(2000);
    expect(obj?.height).toBeCloseTo(1133.86, 1); // 2cm
    expect(obj?.alt).toBe('印章');
    expect(obj?.image?.relId).toBe('rId8');
  });

  it('不带单位的数字按 px 算（CSS 的默认单位），不是 pt', () => {
    const [obj] = objects(`<w:p><w:r><w:pict><v:shape style="width:96;height:48"/></w:pict></w:r></w:p>`);
    expect(obj?.width).toBe(1440); // 96px = 1in = 1440 twips
    expect(obj?.height).toBe(720);
  });

  it('带 imagedata 的形状优先 —— 装饰性形状排在前面时不能被它顶掉', () => {
    const [obj] = objects(
      `<w:p><w:r><w:pict>
        <v:shape style="width:10pt;height:10pt"/>
        <v:shape style="width:50pt;height:20pt"><v:imagedata r:id="rId2"/></v:shape>
      </w:pict></w:r></w:p>`,
    );
    expect(obj?.image?.relId).toBe('rId2');
    expect(obj?.width).toBe(1000);
  });

  it('认不出尺寸时留 0 —— 零尺寸占位比编一个错尺寸强（行不会跟着错位）', () => {
    const [obj] = objects(`<w:p><w:r><w:pict><v:rect style="width:50%"/></w:pict></w:r></w:p>`);
    expect(obj).toEqual({ kind: 'object', objectKind: 'picture', width: 0, height: 0 });
  });
});

describe('collectImages', () => {
  /** 一个够用的假包：只要 `rels` 与 `part` 两个方法（见 `ImagePartSource`） */
  function fakePkg(): Parameters<typeof collectImages>[0] {
    const rels = [
      {
        id: 'rId5',
        type: 'image',
        rawTarget: 'media/image1.png',
        target: '/word/media/image1.png',
        targetMode: 'Internal' as const,
      },
      {
        id: 'rId9',
        type: 'image',
        rawTarget: 'https://example.com/logo.gif',
        target: undefined,
        targetMode: 'External' as const,
      },
    ];
    return {
      rels: () => ({
        byId: (id: string) => rels.find((r) => r.id === id),
        byType: () => [],
        all: () => rels,
      }),
      part: (name: string) =>
        name === '/word/media/image1.png'
          ? { name, contentType: 'image/png', bytes: new Uint8Array([1, 2, 3]) }
          : undefined,
    };
  }

  function bodyBlocks(xml: string) {
    const sink = createDiagnosticSink();
    const doc = parseXml(`<w:document><w:body>${xml}${SECT}</w:body></w:document>`, 'document.xml');
    return parseBody(doc, sink, 'document.xml').sections.flatMap((s) => s.blocks);
  }

  it('只收被引用到的图，同一张图只读一份字节', () => {
    const sink = createDiagnosticSink();
    const blocks = bodyBlocks(`<w:p>${INLINE_PIC}${INLINE_PIC}</w:p>`);
    const images = collectImages(fakePkg(), [{ part: '/word/document.xml', idPrefix: '', blocks }], sink);
    expect(Object.keys(images)).toEqual(['rId5']);
    expect(images.rId5?.bytes).toEqual(new Uint8Array([1, 2, 3]));
    expect(images.rId5?.contentType).toBe('image/png');
    expect(images.rId5?.part).toBe('/word/media/image1.png');
  });

  it('外链图只给地址，不去发网络请求 —— 发不发是宿主的事', () => {
    const sink = createDiagnosticSink();
    const blocks = bodyBlocks(
      `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="635" cy="635"/>
        <a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:link="rId9"/>
        </pic:blipFill></pic:pic></a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p>`,
    );
    const images = collectImages(fakePkg(), [{ part: '/word/document.xml', idPrefix: '', blocks }], sink);
    expect(images.rId9?.url).toBe('https://example.com/logo.gif');
    expect(images.rId9?.bytes).toBeUndefined();
    // 内容类型按扩展名兜底：外链没有 [Content_Types].xml 可查
    expect(images.rId9?.contentType).toBe('image/gif');
  });

  it('悬空引用记诊断继续走，不抛 —— 少一张图好过白屏（原则 1.5）', () => {
    const sink = createDiagnosticSink();
    const blocks = bodyBlocks(
      `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="635" cy="635"/>
        <a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rIdX"/>
        </pic:blipFill></pic:pic></a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p>`,
    );
    const images = collectImages(fakePkg(), [{ part: '/word/document.xml', idPrefix: '', blocks }], sink);
    expect(images).toEqual({});
    expect(sink.list().map((d) => d.code)).toContain('image-missing');
  });
});
