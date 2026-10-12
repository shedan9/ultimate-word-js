/**
 * 组合与画布（`wpg:wgp` / `wpc:wpc` / `v:group`）→ `ObjectContent.group`。
 *
 * 要钉住的是**坐标系**：每层组合把子坐标（`chOff` / `chExt`、`coordsize`）映到自己的框上，
 * 再绕中心翻转、旋转；成员的框与叠完的旋转翻转都要落到对象外框里。另外两处容易错的：
 * 组合本身不能再深搜出一张图（那是第一个成员的），成员里的文本框照样摊进 `textBoxes`。
 */
import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import type { Block, ObjectContent, Paragraph } from './nodes.ts';
import { walkBlocks } from './nodes.ts';
import { parseBody } from './parse-body.ts';

const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>';

function parse(bodyXml: string): { objects: ObjectContent[]; textBoxes: Record<string, Block[]> } {
  const textBoxes: Record<string, Block[]> = {};
  const body = parseBody(
    parseXml(`<w:document><w:body>${bodyXml}${SECT}</w:body></w:document>`, 'document.xml'),
    createDiagnosticSink(),
    'document.xml',
    undefined,
    { textBoxes },
  );
  const objects: ObjectContent[] = [];
  for (const b of walkBlocks(body.sections.flatMap((s) => s.blocks))) {
    if (b.kind !== 'paragraph') continue;
    for (const run of (b as Paragraph).runs) {
      for (const c of run.content) if (c.kind === 'object') objects.push(c);
    }
  }
  return { objects, textBoxes };
}

/** 1 twip = 635 EMU。测试里的数都按 twips 写，EMU 现乘 */
const E = (twips: number): number => twips * 635;

function xfrm(x: number, y: number, w: number, h: number, extra = '', ch = ''): string {
  return `<a:xfrm ${extra}><a:off x="${E(x)}" y="${E(y)}"/><a:ext cx="${E(w)}" cy="${E(h)}"/>${ch}</a:xfrm>`;
}

function ch(x: number, y: number, w: number, h: number): string {
  return `<a:chOff x="${E(x)}" y="${E(y)}"/><a:chExt cx="${E(w)}" cy="${E(h)}"/>`;
}

function anchored(width: number, height: number, graphic: string, uri: string): string {
  return `<w:p><w:r><w:drawing><wp:anchor behindDoc="0" relativeHeight="5">
    <wp:simplePos x="0" y="0"/>
    <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
    <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
    <wp:extent cx="${E(width)}" cy="${E(height)}"/><wp:wrapNone/>
    <wp:docPr id="7" name="组合 7"/>
    <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/${uri}">${graphic}</a:graphicData></a:graphic>
  </wp:anchor></w:drawing></w:r></w:p>`;
}

const BLACK_LN = '<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>';

function rect(frame: string, inner = ''): string {
  return `<wps:wsp><wps:cNvSpPr/><wps:spPr>${frame}<a:prstGeom prst="flowChartProcess"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>${BLACK_LN}</wps:spPr>${inner}<wps:bodyPr/></wps:wsp>`;
}

function connector(frame: string, prst: string, av = ''): string {
  return `<wps:wsp><wps:cNvCnPr/><wps:spPr>${frame}<a:prstGeom prst="${prst}"><a:avLst>${av}</a:avLst></a:prstGeom>
    <a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:tailEnd type="triangle"/></a:ln></wps:spPr><wps:bodyPr/></wps:wsp>`;
}

const TXBX = '<wps:txbx><w:txbxContent><w:p><w:r><w:t>开始</w:t></w:r></w:p></w:txbxContent></wps:txbx>';

describe('DrawingML 组合', () => {
  it('子坐标映到组合的框上：chExt 是 off/ext 的两倍，成员缩一半', () => {
    const group = `<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr>${xfrm(0, 0, 4000, 2000, '', ch(0, 0, 8000, 4000))}</wpg:grpSpPr>
      ${rect(xfrm(4000, 0, 4000, 2000))}
    </wpg:wgp>`;
    const [obj] = parse(anchored(4000, 2000, group, 'wordprocessingGroup')).objects;
    expect(obj?.graphic).toBe('wordprocessingGroup');
    expect(obj?.group).toEqual([
      {
        x: 2000,
        y: 0,
        width: 2000,
        height: 1000,
        shape: { geometry: 'rect', fill: 'FFFFFF', stroke: { color: '000000', width: 20 } },
      },
    ]);
  });

  it('组合本身不深搜图片 —— 那是第一个成员的；成员里的图与文本框各归各的', () => {
    const pic = `<pic:pic><pic:nvPicPr/><pic:blipFill><a:blip r:embed="rId9"/></pic:blipFill>
      <pic:spPr>${xfrm(0, 0, 1000, 1000)}<a:prstGeom prst="rect"/></pic:spPr></pic:pic>`;
    const group = `<wpg:wgp><wpg:grpSpPr>${xfrm(0, 0, 3000, 1000, '', ch(0, 0, 3000, 1000))}</wpg:grpSpPr>
      ${pic}${rect(xfrm(2000, 0, 1000, 1000), TXBX)}</wpg:wgp>`;
    const { objects, textBoxes } = parse(anchored(3000, 1000, group, 'wordprocessingGroup'));
    const obj = objects[0];
    expect(obj?.image).toBeUndefined();
    expect(obj?.shape).toBeUndefined();
    expect(obj?.group?.[0]).toMatchObject({ x: 0, width: 1000, image: { id: 'rId9', relId: 'rId9' } });
    const box = obj?.group?.[1];
    expect(box).toMatchObject({ x: 2000, y: 0, width: 1000, height: 1000 });
    expect(box?.textBox?.id).toBeDefined();
    expect(Object.keys(textBoxes)).toEqual([box?.textBox?.id]);
  });

  it('肘形连接符：几何、调整值（可以出界）与箭头', () => {
    const group = `<wpg:wgp><wpg:grpSpPr>${xfrm(0, 0, 2000, 2000, '', ch(0, 0, 2000, 2000))}</wpg:grpSpPr>
      ${connector(xfrm(0, 0, 2000, 1000), 'bentConnector3', '<a:gd name="adj1" fmla="val -12000"/>')}
    </wpg:wgp>`;
    const [obj] = parse(anchored(2000, 2000, group, 'wordprocessingGroup')).objects;
    expect(obj?.group?.[0]?.shape).toEqual({
      geometry: 'bentConnector3',
      adjust: -12000,
      stroke: { color: '000000', width: 20, tail: { type: 'triangle', w: 'med', len: 'med' } },
    });
  });

  it('嵌套组合的水平翻转：成员的框镜像过去、自己的 flipH 取反', () => {
    // 内层组合占外层的右半边并水平翻转；里面一条从左上到右下的线贴着内层的左边
    const inner = `<wpg:grpSp><wpg:grpSpPr>${xfrm(2000, 0, 2000, 1000, 'flipH="1"', ch(0, 0, 2000, 1000))}</wpg:grpSpPr>
      ${connector(xfrm(0, 0, 500, 1000), 'straightConnector1')}</wpg:grpSp>`;
    const group = `<wpg:wgp><wpg:grpSpPr>${xfrm(0, 0, 4000, 1000, '', ch(0, 0, 4000, 1000))}</wpg:grpSpPr>${inner}</wpg:wgp>`;
    const [obj] = parse(anchored(4000, 1000, group, 'wordprocessingGroup')).objects;
    const line = obj?.group?.[0];
    // 镜像之后贴着内层的右边：2000 + 2000 − 500
    expect(line).toMatchObject({ x: 3500, y: 0, width: 500, height: 1000 });
    expect(line?.shape?.flipH).toBe(true);
    expect(line?.shape?.flipV).toBeUndefined();
    expect(line?.shape?.rotation).toBeUndefined();
  });

  it('组合的竖直翻转折成「转 180° + 水平翻转」，成员自己的旋转跟着变号', () => {
    const group = `<wpg:wgp><wpg:grpSpPr>${xfrm(0, 0, 2000, 2000, 'flipV="1"', ch(0, 0, 2000, 2000))}</wpg:grpSpPr>
      ${rect(xfrm(0, 0, 1000, 500, 'rot="1800000"'))}</wpg:wgp>`;
    const [obj] = parse(anchored(2000, 2000, group, 'wordprocessingGroup')).objects;
    const m = obj?.group?.[0];
    expect(m).toMatchObject({ x: 0, y: 1500, width: 1000, height: 500 });
    // 180 − 30
    expect(m?.shape?.rotation).toBe(150);
    expect(m?.shape?.flipH).toBe(true);
  });

  it('组合旋转 90°：成员的中心绕组合中心转过去，成员的旋转叠上 90', () => {
    const group = `<wpg:wgp><wpg:grpSpPr>${xfrm(0, 0, 2000, 2000, '', ch(0, 0, 2000, 2000))}</wpg:grpSpPr>
      <wpg:grpSp><wpg:grpSpPr>${xfrm(0, 0, 2000, 2000, 'rot="5400000"', ch(0, 0, 2000, 2000))}</wpg:grpSpPr>
      ${rect(xfrm(0, 0, 1000, 500))}</wpg:grpSp></wpg:wgp>`;
    const [obj] = parse(anchored(2000, 2000, group, 'wordprocessingGroup')).objects;
    const m = obj?.group?.[0];
    // 中心 (500, 250) 绕 (1000, 1000) 顺时针转 90° → (1750, 500)
    expect(m).toMatchObject({ x: 1250, y: 250, width: 1000, height: 500 });
    expect(m?.shape?.rotation).toBe(90);
  });

  it('画布：成员的 off 直接是相对画布左上角的 EMU', () => {
    const canvas = `<wpc:wpc><wpc:bg/><wpc:whole/>${rect(xfrm(300, 200, 1000, 400))}</wpc:wpc>`;
    const [obj] = parse(anchored(4000, 3000, canvas, 'wordprocessingCanvas')).objects;
    expect(obj?.group).toHaveLength(1);
    expect(obj?.group?.[0]).toMatchObject({ x: 300, y: 200, width: 1000, height: 400 });
  });

  it('认不出的成员照样占一个位置（画占位框）；一个都没有时整个对象退回占位', () => {
    const custom = `<wps:wsp><wps:spPr>${xfrm(0, 0, 500, 500)}<a:custGeom/></wps:spPr><wps:bodyPr/></wps:wsp>`;
    const mixed = `<wpg:wgp><wpg:grpSpPr>${xfrm(0, 0, 1000, 500, '', ch(0, 0, 1000, 500))}</wpg:grpSpPr>${custom}${rect(xfrm(500, 0, 500, 500))}</wpg:wgp>`;
    const [obj] = parse(anchored(1000, 500, mixed, 'wordprocessingGroup')).objects;
    expect(obj?.group?.[0]).toEqual({ x: 0, y: 0, width: 500, height: 500 });
    const empty = `<wpg:wgp><wpg:grpSpPr>${xfrm(0, 0, 1000, 500)}</wpg:grpSpPr></wpg:wgp>`;
    const [none] = parse(anchored(1000, 500, empty, 'wordprocessingGroup')).objects;
    expect(none?.group).toBeUndefined();
    expect(none?.shape).toBeUndefined();
  });
});

describe('VML 组合', () => {
  function pict(inner: string): string {
    return `<w:p><w:r><w:pict>${inner}</w:pict></w:r></w:p>`;
  }

  it('coordsize 是组合内部的格子：成员的 left 不是 px', () => {
    const { objects, textBoxes } = parse(
      pict(`<v:group style="position:absolute;margin-left:0;margin-top:0;width:200pt;height:100pt;z-index:3"
          coordorigin="1000,0" coordsize="4000,2000">
        <v:rect style="position:absolute;left:3000;top:1000;width:2000;height:1000" fillcolor="#ff0" strokecolor="red"/>
        <v:shape type="#_x0000_t202" style="position:absolute;left:1000;top:0;width:1000;height:500">
          <v:textbox><w:txbxContent><w:p><w:r><w:t>判断</w:t></w:r></w:p></w:txbxContent></v:textbox>
        </v:shape>
        <v:shape type="#_x0000_t34" adj="2700" style="position:absolute;left:1000;top:500;width:2000;height:1000">
          <v:stroke endarrow="block"/>
        </v:shape>
        <v:shape style="position:absolute;left:1000;top:0;width:10;height:10">
          <v:imagedata r:id="rId4"/>
        </v:shape>
      </v:group>`),
    );
    const obj = objects[0];
    expect(obj?.anchor).toBeDefined();
    expect(obj?.image).toBeUndefined();
    expect(obj?.width).toBe(4000);
    // 4000 twips / 4000 格 = 1；原点在 1000
    expect(obj?.group?.[0]).toMatchObject({
      x: 2000,
      y: 1000,
      width: 2000,
      height: 1000,
      shape: { fill: 'FFFF00', stroke: { color: 'FF0000', width: 15 } },
    });
    expect(obj?.group?.[1]).toMatchObject({ x: 0, y: 0, width: 1000, height: 500 });
    expect(obj?.group?.[1]?.textBox?.id).toBe(Object.keys(textBoxes)[0]);
    expect(obj?.group?.[2]?.shape).toMatchObject({
      geometry: 'bentConnector3',
      adjust: 12500,
      stroke: { tail: { type: 'triangle' } },
    });
    expect(obj?.group?.[3]?.image).toEqual({ id: 'rId4', relId: 'rId4' });
  });

  it('组合里的 v:line：端点是组合坐标，方向折成翻转', () => {
    const [obj] = parse(
      pict(`<v:group style="position:absolute;width:100pt;height:100pt" coordsize="1000,1000">
        <v:line from="800,0" to="200,500" strokeweight="1pt"/>
      </v:group>`),
    ).objects;
    expect(obj?.group?.[0]).toMatchObject({
      x: 400,
      y: 0,
      width: 1200,
      height: 1000,
      shape: { geometry: 'line', flipH: true },
    });
  });
});
