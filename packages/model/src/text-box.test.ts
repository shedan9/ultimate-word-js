/**
 * 文本框（`wps:txbx` / `v:textbox`）与 VML 的绝对定位。
 *
 * 三类要钉住的：内容摊进旁表且 id 不撞车（尤其是**不挪动正文的 id** —— 回写靠两次解析 id 一致）、
 * 内边距 / 填充 / 轮廓的单位与缺省值（VML 的缺省是「填白描黑」，与 DrawingML 相反）、
 * 以及 VML 的 `position:absolute` 折成浮动定位。
 */
import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import type { Block, ObjectContent, Paragraph } from './nodes.ts';
import { walkBlocks } from './nodes.ts';
import { parseBody } from './parse-body.ts';
import { parseTheme } from './theme.ts';

const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>';

function parse(bodyXml: string, withSink = true, themeColors?: Record<string, string>) {
  const textBoxes: Record<string, Block[]> = {};
  const body = parseBody(
    parseXml(`<w:document><w:body>${bodyXml}${SECT}</w:body></w:document>`, 'document.xml'),
    createDiagnosticSink(),
    'document.xml',
    undefined,
    withSink ? { textBoxes, ...(themeColors === undefined ? {} : { themeColors }) } : {},
  );
  const objects: ObjectContent[] = [];
  const ids: string[] = [];
  for (const b of walkBlocks(body.sections.flatMap((s) => s.blocks))) {
    ids.push(b.id);
    if (b.kind !== 'paragraph') continue;
    for (const run of (b as Paragraph).runs) {
      ids.push(run.id);
      for (const c of run.content) if (c.kind === 'object') objects.push(c);
    }
  }
  return { objects, textBoxes, ids };
}

/** 一个 100×50pt 的浮动 DrawingML 文本框，内容两段 */
function wpsBox(opts: { bodyPr?: string; spPr?: string; style?: string } = {}): string {
  return `<w:r><w:drawing><wp:anchor behindDoc="0" relativeHeight="3">
    <wp:simplePos x="0" y="0"/>
    <wp:positionH relativeFrom="page"><wp:posOffset>12700</wp:posOffset></wp:positionH>
    <wp:positionV relativeFrom="page"><wp:posOffset>25400</wp:posOffset></wp:positionV>
    <wp:extent cx="1270000" cy="635000"/><wp:wrapNone/>
    <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
      <wps:wsp>
        <wps:spPr>${opts.spPr ?? ''}</wps:spPr>
        ${opts.style ?? ''}
        <wps:txbx><w:txbxContent>
          <w:p><w:r><w:t>框里第一段</w:t></w:r></w:p>
          <w:p><w:r><w:t>第二段</w:t></w:r></w:p>
        </w:txbxContent></wps:txbx>
        <wps:bodyPr ${opts.bodyPr ?? ''}/>
      </wps:wsp>
    </a:graphicData></a:graphic>
  </wp:anchor></w:drawing></w:r>`;
}

describe('DrawingML 文本框', () => {
  it('内容摊进旁表，节点 id 带文本框自己的前缀', () => {
    const { objects, textBoxes } = parse(`<w:p>${wpsBox()}</w:p>`);
    const ref = objects[0]?.textBox;
    expect(ref?.id).toBe('tb0');
    const blocks = textBoxes.tb0 ?? [];
    expect(blocks.map((b) => b.id)).toEqual(['tb0:p0', 'tb0:p1']);
    const first = blocks[0] as Paragraph;
    expect(first.runs[0]?.id).toBe('tb0:r0');
    expect(first.runs[0]?.content[0]).toEqual({ kind: 'text', text: '框里第一段' });
  });

  it('不要内容时（回写重新解析）照样分配 id，正文的 id 与加载时一字不差', () => {
    const xml = `<w:p><w:r><w:t>前</w:t></w:r>${wpsBox()}</w:p><w:p><w:r><w:t>后</w:t></w:r></w:p>`;
    const withSink = parse(xml);
    const without = parse(xml, false);
    expect(without.ids).toEqual(withSink.ids);
    // 引用照留（回写不看它），只是旁表里没有内容
    expect(without.objects[0]?.textBox).toEqual(withSink.objects[0]?.textBox);
    expect(without.textBoxes).toEqual({});
    // 正文里文本框后面那一段不会因为框里多了两段而挪号
    expect(withSink.ids).toContain('p1');
  });

  it('内边距缺省左右 0.1 英寸、上下 0.05 英寸；写了按 EMU 换算', () => {
    const def = parse(`<w:p>${wpsBox()}</w:p>`).objects[0]?.textBox;
    expect(def?.inset).toEqual({ left: 144, top: 72, right: 144, bottom: 72 });
    expect(def?.vAlign).toBe('top');
    const set = parse(
      `<w:p>${wpsBox({ bodyPr: 'lIns="0" tIns="12700" rIns="0" bIns="0" anchor="ctr"' })}</w:p>`,
    ).objects[0]?.textBox;
    expect(set?.inset).toEqual({ left: 0, top: 20, right: 0, bottom: 0 });
    expect(set?.vAlign).toBe('center');
  });

  it('填充与轮廓：srgb / 预设色，线宽按 EMU；noFill 不画', () => {
    const [obj] = parse(
      `<w:p>${wpsBox({
        spPr: '<a:solidFill><a:srgbClr val="ffeeaa"/></a:solidFill><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln>',
      })}</w:p>`,
    ).objects;
    expect(obj?.shape).toEqual({ fill: 'FFEEAA', stroke: { color: '000000', width: 20 } });

    const [none] = parse(`<w:p>${wpsBox({ spPr: '<a:noFill/><a:ln><a:noFill/></a:ln>' })}</w:p>`).objects;
    expect(none?.shape).toBeUndefined();
  });

  it('主题色按文档主题换算，lumMod / lumOff 调亮度；主题没配色时退到 Office 默认', () => {
    const spPr =
      '<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr></a:solidFill>';
    const themed = parse(`<w:p>${wpsBox({ spPr })}</w:p>`, true, { accent1: 'FF0000' }).objects[0];
    expect(themed?.shape?.fill).toBe('800000');
    const fallback = parse(
      `<w:p>${wpsBox({ spPr: '<a:solidFill><a:schemeClr val="bg1"/></a:solidFill>' })}</w:p>`,
    ).objects[0];
    expect(fallback?.shape?.fill).toBe('FFFFFF');
  });

  it('spPr 没写时退到形状样式的 fillRef / lnRef，idx 0 是「无」', () => {
    const style = `<wps:style>
      <a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>
      <a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>
    </wps:style>`;
    const [obj] = parse(`<w:p>${wpsBox({ style })}</w:p>`).objects;
    expect(obj?.shape).toEqual({ stroke: { color: '4472C4', width: 20 } });
  });

  it('框里插的图是框的内容，不是框的填充 —— a:blip 深搜不钻进 w:txbxContent', () => {
    const inner = `<w:r><w:drawing><wp:inline><wp:extent cx="12700" cy="12700"/>
      <a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill></pic:pic>
      </a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
    const { objects, textBoxes } = parse(
      `<w:p>${wpsBox().replace('<w:t>框里第一段</w:t></w:r>', `<w:t>框里第一段</w:t></w:r>${inner}`)}</w:p>`,
    );
    expect(objects[0]?.image).toBeUndefined();
    const first = textBoxes.tb0?.[0] as Paragraph;
    const pic = first.runs[1]?.content[0];
    // 框里的图 id 带文本框自己的前缀
    expect(pic?.kind === 'object' && pic.image?.id).toBe('tb0:rId7');
  });

  it('mc:AlternateContent 只收 Choice —— 同一个框不摊出两份', () => {
    const xml = `<w:p><w:r><mc:AlternateContent>
      <mc:Choice Requires="wps">${wpsBox().replace('<w:r>', '').replace('</w:r>', '')}</mc:Choice>
      <mc:Fallback><w:pict><v:shape style="position:absolute;width:100pt;height:50pt">
        <v:textbox><w:txbxContent><w:p><w:r><w:t>框里第一段</w:t></w:r></w:p></w:txbxContent></v:textbox>
      </v:shape></w:pict></mc:Fallback>
    </mc:AlternateContent></w:r></w:p>`;
    const { objects, textBoxes } = parse(xml);
    expect(objects).toHaveLength(1);
    expect(Object.keys(textBoxes)).toEqual(['tb0']);
  });
});

describe('VML 文本框与绝对定位', () => {
  /** WPS 从 PDF 转出来的页脚里就是这种写法 */
  const VML = `<w:p><w:r><w:pict><v:shape id="_x0000_s1" type="#_x0000_t202"
      style="position:absolute;left:0pt;margin-left:183.45pt;margin-top:3.4pt;height:13.3pt;width:49pt;z-index:251686912;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;v-text-anchor:middle"
      filled="f" stroked="f">
      <v:imagedata o:title=""/>
      <v:textbox inset="0mm,0mm,0mm,0mm"><w:txbxContent><w:p><w:r><w:t>2026-1（4）</w:t></w:r></w:p></w:txbxContent></v:textbox>
    </v:shape></w:pict></w:r></w:p>`;

  it('position:absolute 折成浮动定位：偏移 = left + margin-left，参照物换成 DrawingML 的名字', () => {
    const [obj] = parse(VML).objects;
    expect(obj?.width).toBe(980);
    expect(obj?.anchor).toMatchObject({
      wrap: 'none',
      behindDoc: false,
      z: 251686912,
      h: { relativeFrom: 'page', offset: 3669 },
      v: { relativeFrom: 'page', offset: 68 },
    });
  });

  it('内边距按 CSS 长度读，v-text-anchor 给纵向对齐；filled / stroked 为 f 时不画', () => {
    const [obj] = parse(VML).objects;
    expect(obj?.textBox).toMatchObject({ inset: { left: 0, top: 0, right: 0, bottom: 0 }, vAlign: 'center' });
    expect(obj?.shape).toBeUndefined();
  });

  it('VML 的缺省是填白、描黑 0.75pt（与 DrawingML 相反），v:stroke 的颜色与粗细盖过属性', () => {
    const [obj] = parse(
      `<w:p><w:r><w:pict><v:rect style="width:100pt;height:20pt" strokecolor="red">
        <v:stroke weight="1.5pt" color="#00ff00"/>
        <v:textbox><w:txbxContent><w:p/></w:txbxContent></v:textbox>
      </v:rect></w:pict></w:r></w:p>`,
    ).objects;
    expect(obj?.shape).toEqual({ fill: 'FFFFFF', stroke: { color: '00FF00', width: 30 } });
    // 缺省内边距与 DrawingML 相同
    expect(obj?.textBox?.inset).toEqual({ left: 144, top: 72, right: 144, bottom: 72 });
    // 没有 position:absolute 就是内嵌
    expect(obj?.anchor).toBeUndefined();
  });

  it('负的 z-index 是衬于文字下方；w10:wrap 给环绕方式，缺省参照物是栏与段落', () => {
    const [obj] = parse(
      `<w:p><w:r><w:pict><v:shape style="position:absolute;margin-left:10pt;margin-top:5pt;width:50pt;height:50pt;z-index:-3;mso-position-horizontal:center">
        <w10:wrap type="square" side="left"/>
      </v:shape></w:pict></w:r></w:p>`,
    ).objects;
    expect(obj?.anchor).toMatchObject({
      wrap: 'square',
      wrapText: 'left',
      behindDoc: true,
      z: 3,
      h: { relativeFrom: 'column', align: 'center' },
      v: { relativeFrom: 'paragraph', offset: 100 },
    });
    // 缺省环绕距离：左右 9pt
    expect(obj?.anchor?.dist).toEqual({ top: 0, bottom: 0, left: 180, right: 180 });
  });
});

describe('主题配色', () => {
  it('a:clrScheme：srgbClr 取 val，sysClr 取存盘时的 lastClr', () => {
    const theme = parseTheme(
      parseXml(
        `<a:theme><a:themeElements><a:clrScheme name="Office">
          <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
          <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
          <a:accent1><a:srgbClr val="5b9bd5"/></a:accent1>
        </a:clrScheme></a:themeElements></a:theme>`,
      ),
    );
    expect(theme.colors).toEqual({ dk1: '000000', lt1: 'FFFFFF', accent1: '5B9BD5' });
  });
});
