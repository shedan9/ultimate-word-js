/**
 * 纯形状（`wps:wsp` 没有文本框 / VML 的基本形状）的外观。
 *
 * 要钉住的是**画错了才看得出来**的几处：直线的方向全靠翻转（从左下到右上是 `flipV`）、
 * 认不出的几何**不收**（照旧画占位框，不画一个错形状）、线没有填充、VML 的 `v:line`
 * 外框不在 style 里而在两个端点上、VML 与 DrawingML 的虚线 / 箭头名字不同但折成同一套。
 */
import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import type { Block, ObjectContent, Paragraph } from './nodes.ts';
import { walkBlocks } from './nodes.ts';
import { parseBody } from './parse-body.ts';

const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>';

function objects(bodyXml: string): ObjectContent[] {
  const textBoxes: Record<string, Block[]> = {};
  const body = parseBody(
    parseXml(`<w:document><w:body>${bodyXml}${SECT}</w:body></w:document>`, 'document.xml'),
    createDiagnosticSink(),
    'document.xml',
    undefined,
    { textBoxes },
  );
  const out: ObjectContent[] = [];
  for (const b of walkBlocks(body.sections.flatMap((s) => s.blocks))) {
    if (b.kind !== 'paragraph') continue;
    for (const run of (b as Paragraph).runs) {
      for (const c of run.content) if (c.kind === 'object') out.push(c);
    }
  }
  return out;
}

/** 一个浮动的 DrawingML 形状，外框 cx × cy（EMU），spPr 的内容自己给 */
function wsp(spPr: string, opts: { cx?: number; cy?: number; style?: string } = {}): string {
  return `<w:p><w:r><w:drawing><wp:anchor behindDoc="0" relativeHeight="2">
    <wp:simplePos x="0" y="0"/>
    <wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>
    <wp:positionV relativeFrom="paragraph"><wp:posOffset>127000</wp:posOffset></wp:positionV>
    <wp:extent cx="${opts.cx ?? 5600700}" cy="${opts.cy ?? 0}"/><wp:wrapNone/>
    <wp:docPr id="3" name="直接连接符 3"/>
    <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
      <wps:wsp><wps:cNvCnPr/><wps:spPr>${spPr}</wps:spPr>${opts.style ?? ''}<wps:bodyPr/></wps:wsp>
    </a:graphicData></a:graphic>
  </wp:anchor></w:drawing></w:r></w:p>`;
}

/** 公文红头下那条线的典型写法：通栏、零高、红色 2.25pt */
const RED_RULE = `<a:xfrm><a:off x="0" y="0"/><a:ext cx="5600700" cy="0"/></a:xfrm>
  <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
  <a:ln w="28575"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>`;

describe('DrawingML 纯形状', () => {
  it('红头分隔线：直线几何、零高的外框、红色轮廓，没有填充', () => {
    const [obj] = objects(wsp(RED_RULE));
    expect(obj?.width).toBe(8820);
    expect(obj?.height).toBe(0);
    expect(obj?.shape).toEqual({ geometry: 'line', stroke: { color: 'FF0000', width: 45 } });
    expect(obj?.textBox).toBeUndefined();
  });

  it('样式引用的 fillRef 对直线不起作用（线没有内部）', () => {
    const style = `<wps:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>
      <a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef></wps:style>`;
    const [obj] = objects(wsp('<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>', { style }));
    expect(obj?.shape).toEqual({ geometry: 'line', stroke: { color: '4472C4', width: 10 } });
  });

  it('直线的方向全靠翻转：flipV 是从左下画到右上', () => {
    const [obj] = objects(
      wsp(
        '<a:xfrm flipV="1"/><a:prstGeom prst="line"/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln>',
        {
          cy: 635000,
        },
      ),
    );
    expect(obj?.shape?.flipV).toBe(true);
    expect(obj?.shape?.flipH).toBeUndefined();
  });

  it('虚线与线端：prstDash 认预设名，箭头带三档大小，缺省 med', () => {
    const [obj] = objects(
      wsp(`<a:prstGeom prst="straightConnector1"/>
        <a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="sysDash"/>
          <a:headEnd type="oval" w="sm"/><a:tailEnd type="triangle" w="lg" len="lg"/></a:ln>`),
    );
    expect(obj?.shape?.stroke).toEqual({
      color: '000000',
      width: 20,
      dash: 'sysDash',
      head: { type: 'oval', w: 'sm', len: 'med' },
      tail: { type: 'triangle', w: 'lg', len: 'lg' },
    });
  });

  it('流程图符号折成同形的基本形状，调整值从 avLst 读', () => {
    const geom = (prst: string, av = '') =>
      objects(wsp(`<a:prstGeom prst="${prst}"><a:avLst>${av}</a:avLst></a:prstGeom>`, { cy: 635000 }))[0]
        ?.shape;
    expect(geom('flowChartDecision')?.geometry).toBe('diamond');
    expect(geom('flowChartProcess')?.geometry).toBe('rect');
    expect(geom('roundRect', '<a:gd name="adj" fmla="val 30000"/>')).toMatchObject({
      geometry: 'roundRect',
      adjust: 30000,
    });
  });

  it('认不出的几何不收 —— 渲染层照旧画占位框，不画一个错形状', () => {
    const [arrow] = objects(wsp('<a:prstGeom prst="rightArrow"/>', { cy: 635000 }));
    expect(arrow?.shape).toBeUndefined();
    const [custom] = objects(wsp('<a:custGeom><a:pathLst/></a:custGeom>', { cy: 635000 }));
    expect(custom?.shape).toBeUndefined();
  });

  it('透明的矩形也有 shape（不是占位框），只是什么都不画', () => {
    const [obj] = objects(
      wsp('<a:prstGeom prst="rect"/><a:noFill/><a:ln><a:noFill/></a:ln>', { cy: 635000 }),
    );
    expect(obj?.shape).toEqual({ geometry: 'rect' });
  });

  it('旋转从 a:xfrm@rot 换成度', () => {
    const [obj] = objects(wsp('<a:xfrm rot="2700000"/><a:prstGeom prst="ellipse"/>', { cy: 635000 }));
    expect(obj?.shape).toMatchObject({ geometry: 'ellipse', rotation: 45 });
  });
});

describe('VML 基本形状', () => {
  it('v:line：外框是两个端点围成的矩形，偏移加上端点的较小坐标，线不填充', () => {
    const [obj] = objects(`<w:p><w:r><w:pict>
      <v:line style="position:absolute;z-index:1;mso-position-horizontal-relative:margin;mso-position-vertical-relative:text"
        from="0,7.8pt" to="441pt,7.8pt" strokecolor="red" strokeweight="2.25pt"/>
    </w:pict></w:r></w:p>`);
    expect(obj?.width).toBe(8820);
    expect(obj?.height).toBe(0);
    expect(obj?.anchor?.h).toEqual({ relativeFrom: 'margin', offset: 0 });
    expect(obj?.anchor?.v).toEqual({ relativeFrom: 'paragraph', offset: 156 });
    expect(obj?.shape).toEqual({ geometry: 'line', stroke: { color: 'FF0000', width: 45 } });
  });

  it('v:line 从右上画到左下 = 水平翻转（与 DrawingML 的直线同一种说法）', () => {
    const [obj] = objects(`<w:p><w:r><w:pict>
      <v:line style="position:absolute" from="100pt,0" to="0,50pt"/>
    </w:pict></w:r></w:p>`);
    expect(obj?.shape).toMatchObject({ geometry: 'line', flipH: true });
    expect(obj?.shape?.flipV).toBeUndefined();
  });

  it('直线连接符（_x0000_t32）的方向在 style 的 flip 里；虚线与箭头折成 DrawingML 的名字', () => {
    const [obj] = objects(`<w:p><w:r><w:pict>
      <v:shape type="#_x0000_t32" style="position:absolute;width:100pt;height:50pt;flip:y" o:connectortype="straight">
        <v:stroke dashstyle="longDash" endarrow="block" endarrowwidth="wide"/>
      </v:shape>
    </w:pict></w:r></w:p>`);
    expect(obj?.shape).toEqual({
      geometry: 'line',
      flipV: true,
      stroke: { color: '000000', width: 15, dash: 'lgDash', tail: { type: 'triangle', w: 'lg', len: 'med' } },
    });
  });

  it('v:roundrect 的 arcsize（占短边一半的比例）换成 adj；缺省填白描黑', () => {
    const [obj] = objects(`<w:p><w:r><w:pict>
      <v:roundrect style="width:100pt;height:50pt" arcsize="10923f"/>
    </w:pict></w:r></w:p>`);
    expect(obj?.shape).toEqual({
      geometry: 'roundRect',
      adjust: 8334,
      fill: 'FFFFFF',
      stroke: { color: '000000', width: 15 },
    });
  });

  it('自定义路径的 v:shape 认不出，装着图片的形状不叠形状', () => {
    const [path] = objects(`<w:p><w:r><w:pict>
      <v:shape style="width:10pt;height:10pt" path="m,l10,10e"/>
    </w:pict></w:r></w:p>`);
    expect(path?.shape).toBeUndefined();
    const [pic] = objects(`<w:p><w:r><w:pict>
      <v:shape type="#_x0000_t75" style="width:10pt;height:10pt"><v:imagedata r:id="rId4"/></v:shape>
    </w:pict></w:r></w:p>`);
    expect(pic?.shape).toBeUndefined();
    expect(pic?.image?.relId).toBe('rId4');
  });
});
