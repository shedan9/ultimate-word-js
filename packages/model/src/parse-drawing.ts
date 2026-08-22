/**
 * `w:drawing`（DrawingML）与 `w:pict`（VML）→ `ObjectContent`。
 *
 * 单独一个文件，是因为这里要跨**四个命名空间**下钻（`wp:` 定位与外框、`a:` 图形本身、
 * `pic:` 图片、`r:` 关系），与 `parse-body.ts` 那种「一层 switch 收元素」不是一回事。
 *
 * ## 三条容易搞反的
 *
 * ① **尺寸取 `wp:extent`，不取图片自己的像素尺寸。** extent 是**显示**尺寸：用户在 Word 里
 *    把图拖小，extent 跟着变、图片字节一个都不变。按像素尺寸画会让每张被拖过的图都错。
 *    `a:ext`（`pic:spPr/a:xfrm/a:ext`）与 extent 通常相等，但**旋转之后 extent 是外接矩形**，
 *    两者会差一截 —— 占位与行高要的是 extent（它占多大地方），画图要的是 `a:ext`。
 * ② **`a:blip` 要往深处找，不能按固定路径取。** 规范路径是
 *    `a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip`，但同一张图也可能包在
 *    `mc:AlternateContent` 的 Choice 里、或者挂在形状的填充上（`wps:wsp/wps:spPr/a:blipFill`）。
 *    深搜第一个 `a:blip` 把这些写法一网打尽；而图表 / SmartArt 里根本没有 blip ——
 *    「找不到 blip 就是画不出来」正好是我们要的判据。
 * ③ **裁剪（`a:srcRect`）不改外框。** 它的四个值是「从这条边往里裁掉百分之多少」，
 *    裁剩的那一块被**拉伸**回原来的外框。当成「外框变小」会让图整个缩水。
 *
 * ## VML（`w:pict`）只取两样：图片引用与外框尺寸
 *
 * VML 图形本身是非目标（开发计划 §5），但老文件（以及 Word 为兼容写的 Fallback）里的
 * 图片都走这条路 —— 公文的红头与印章尤其常见。只解析 `v:shape@style` 的 width / height
 * 与 `v:imagedata@r:id`，够把图画在对的地方；`v:line` / `v:rect` 这些真的 VML 图形
 * 仍然只得到一个占位框。
 */
import { emuToTwips, inchToTwips, mmToTwips, ptToTwips, pxToTwips, type Twips } from '@uw/core';
import type { XmlElement } from '@uw/ooxml';
import { attr, child, children, textContent } from '@uw/ooxml';
import type { AnchorPos, DrawingAnchor, ImageRef, ObjectContent } from './nodes.ts';
import { attrInt, attrOnOff, put } from './xml-values.ts';

/** `a:xfrm@rot` 的单位是 1/60000 度 */
const ROT_PER_DEGREE = 60000;
/** `a:srcRect` 的单位是千分之一个百分点（100% = 100000） */
const PERCENT_1000 = 100000;

const WRAP_ELEMENTS: Record<string, DrawingAnchor['wrap']> = {
  'wp:wrapNone': 'none',
  'wp:wrapSquare': 'square',
  'wp:wrapTight': 'tight',
  'wp:wrapThrough': 'through',
  'wp:wrapTopAndBottom': 'topAndBottom',
};

/**
 * `w:drawing` → 一个对象片段。
 *
 * `idPrefix` 与节点 id 用的是同一个（页眉页脚各带一个，正文是空串）：图片资源按
 * 「前缀 + 关系 id」索引，不带前缀时页眉里的 rId1 会顶掉正文里的 rId1。
 */
export function parseDrawing(drawing: XmlElement, idPrefix: string): ObjectContent {
  const inline = child(drawing, 'wp:inline');
  const frame = inline ?? child(drawing, 'wp:anchor');
  const out: ObjectContent = {
    kind: 'object',
    objectKind: 'drawing',
    ...frameExtent(frame),
  };
  if (frame === undefined) return out;

  put(out, 'alt', altText(frame));
  put(out, 'graphic', graphicUri(frame));
  put(out, 'image', blipRef(frame, idPrefix));
  if (inline === undefined) out.anchor = anchorOf(frame);
  return out;
}

/**
 * `w:pict` → 一个对象片段（VML）。
 *
 * 尺寸藏在 `style="width:100pt;height:50pt"` 里，单位可以是 pt / px / in / cm / mm ——
 * CSS 的默认单位是 **px 不是 pt**（VML 照 CSS 写），漏掉这条会让所有不带单位的图
 * 小一截（1px = 0.75pt）。
 */
export function parsePict(pict: XmlElement, idPrefix: string): ObjectContent {
  const shape = findShape(pict);
  const style = parseStyle(shape === undefined ? undefined : attr(shape, 'style'));
  const out: ObjectContent = {
    kind: 'object',
    objectKind: 'picture',
    width: cssLength(style.width) ?? 0,
    height: cssLength(style.height) ?? 0,
  };
  if (shape === undefined) return out;

  put(out, 'alt', nonEmpty(attr(shape, 'alt') ?? attr(shape, 'o:title')));
  const data = child(shape, 'v:imagedata');
  const relId = data === undefined ? undefined : (attr(data, 'r:id') ?? attr(data, 'o:relid'));
  if (relId !== undefined) out.image = { id: `${idPrefix}${relId}`, relId };
  return out;
}

// ── DrawingML 的各个零件 ──────────────────────────────────────────────────────

/**
 * 外框尺寸 —— `wp:extent`，也就是这个对象在版面上**占多大地方**。
 *
 * 缺席时是 0：布局层按「零尺寸占位」处理（行不会错位，图不见了），这不是 bug，
 * 是「宁可少画一张图，也不要整段文字挪错」。
 */
function frameExtent(frame: XmlElement | undefined): { width: Twips; height: Twips } {
  const extent = frame === undefined ? undefined : child(frame, 'wp:extent');
  return {
    width: emuToTwips(attrInt(extent, 'cx') ?? 0),
    height: emuToTwips(attrInt(extent, 'cy') ?? 0),
  };
}

/**
 * 可选文本。`descr` 是「说明」、`title` 是「标题」，两个都是人写的；退到 `name`
 * 只是为了占位框上有字可显 —— 它是 Word 自动起的「图片 3」这种编号名。
 */
function altText(frame: XmlElement): string | undefined {
  const docPr = child(frame, 'wp:docPr');
  if (docPr === undefined) return undefined;
  return nonEmpty(attr(docPr, 'descr') ?? attr(docPr, 'title') ?? attr(docPr, 'name'));
}

/** `a:graphicData@uri` 的最后一段：`…/picture` → `picture`、`…/chart` → `chart` */
function graphicUri(frame: XmlElement): string | undefined {
  const graphic = child(frame, 'a:graphic');
  const data = graphic === undefined ? undefined : child(graphic, 'a:graphicData');
  const uri = data === undefined ? undefined : attr(data, 'uri');
  return uri === undefined ? undefined : nonEmpty(uri.slice(uri.lastIndexOf('/') + 1));
}

function blipRef(frame: XmlElement, idPrefix: string): ImageRef | undefined {
  const blip = descendant(frame, 'a:blip');
  if (blip === undefined) return undefined;
  const embed = attr(blip, 'r:embed');
  const link = attr(blip, 'r:link');
  const relId = embed ?? link;
  // 两个都没有的 blip 是「有图片填充，但引用丢了」—— Word 自己也画不出来，当没有图
  if (relId === undefined) return undefined;

  const ref: ImageRef = { id: `${idPrefix}${relId}`, relId };
  if (embed === undefined) ref.linked = true;
  put(ref, 'crop', cropOf(frame));

  const xfrm = descendant(frame, 'a:xfrm');
  if (xfrm !== undefined) {
    const rot = attrInt(xfrm, 'rot');
    if (rot !== undefined && rot !== 0) ref.rotation = rot / ROT_PER_DEGREE;
    if (attrOnOff(xfrm, 'flipH') === true) ref.flipH = true;
    if (attrOnOff(xfrm, 'flipV') === true) ref.flipV = true;
  }
  return ref;
}

/**
 * `a:srcRect`。四个属性缺省为 0，**允许负值**（往外扩，Word 的「裁剪成留白」），
 * 所以不夹到 0 —— 夹了那圈留白会被吃掉，图跟着偏。
 */
function cropOf(frame: XmlElement): ImageRef['crop'] | undefined {
  const rect = descendant(frame, 'a:srcRect');
  if (rect === undefined) return undefined;
  const side = (name: string) => (attrInt(rect, name) ?? 0) / PERCENT_1000;
  const crop = { left: side('l'), top: side('t'), right: side('r'), bottom: side('b') };
  const empty = crop.left === 0 && crop.top === 0 && crop.right === 0 && crop.bottom === 0;
  return empty ? undefined : crop;
}

function anchorOf(anchor: XmlElement): DrawingAnchor {
  const dist = (name: string): Twips => emuToTwips(attrInt(anchor, name) ?? 0);
  const out: DrawingAnchor = {
    wrap: wrapOf(anchor),
    behindDoc: attrOnOff(anchor, 'behindDoc') === true,
    z: attrInt(anchor, 'relativeHeight') ?? 0,
    h: positionOf(anchor, 'wp:positionH'),
    v: positionOf(anchor, 'wp:positionV'),
    dist: { top: dist('distT'), bottom: dist('distB'), left: dist('distL'), right: dist('distR') },
  };
  // simplePos="1" 时规范说忽略 positionH / V，改用这一对绝对坐标（相对纸左上角）
  if (attrOnOff(anchor, 'simplePos') === true) {
    const simple = child(anchor, 'wp:simplePos');
    if (simple !== undefined) {
      out.h = { relativeFrom: 'page', offset: emuToTwips(attrInt(simple, 'x') ?? 0) };
      out.v = { relativeFrom: 'page', offset: emuToTwips(attrInt(simple, 'y') ?? 0) };
    }
  }
  return out;
}

/** 环绕方式由**哪个元素在场**决定（规范里是 choice），一个都没有时按 `none` */
function wrapOf(anchor: XmlElement): DrawingAnchor['wrap'] {
  for (const el of children(anchor)) {
    const wrap = WRAP_ELEMENTS[el.name];
    if (wrap !== undefined) return wrap;
  }
  return 'none';
}

function positionOf(anchor: XmlElement, name: string): AnchorPos {
  const el = child(anchor, name);
  // 缺席按「相对栏 / 段落偏移 0」处理，也就是落在文字流当前的位置上
  if (el === undefined) return { relativeFrom: name === 'wp:positionH' ? 'column' : 'paragraph' };
  const pos: AnchorPos = { relativeFrom: attr(el, 'relativeFrom') ?? 'column' };
  const offset = child(el, 'wp:posOffset');
  const align = child(el, 'wp:align');
  if (offset !== undefined) {
    const n = Number.parseInt(textContent(offset).trim(), 10);
    if (!Number.isNaN(n)) pos.offset = emuToTwips(n);
  } else if (align !== undefined) {
    put(pos, 'align', nonEmpty(textContent(align).trim()));
  }
  return pos;
}

// ── VML 的两样 ────────────────────────────────────────────────────────────────

/** 带 `v:imagedata` 的形状优先；一个都没有时退回第一个形状（只为拿尺寸画占位框） */
function findShape(pict: XmlElement): XmlElement | undefined {
  const shapes = descendants(pict, 'v:shape');
  return shapes.find((s) => child(s, 'v:imagedata') !== undefined) ?? shapes[0];
}

function parseStyle(style: string | undefined): Record<string, string> {
  const out: Record<string, string> = {};
  if (style === undefined) return out;
  for (const decl of style.split(';')) {
    const i = decl.indexOf(':');
    if (i < 0) continue;
    out[decl.slice(0, i).trim().toLowerCase()] = decl.slice(i + 1).trim();
  }
  return out;
}

const CSS_UNITS: Record<string, (n: number) => Twips> = {
  pt: ptToTwips,
  px: (n) => pxToTwips(n),
  in: inchToTwips,
  cm: (n) => mmToTwips(n * 10),
  mm: mmToTwips,
};

/** 认不出的单位当 px —— CSS 的默认单位就是它，VML 里不带单位的数字也是 px */
function cssLength(value: string | undefined): Twips | undefined {
  if (value === undefined) return undefined;
  const m = /^(-?[\d.]+)\s*([a-z%]*)$/i.exec(value.trim());
  if (m === null) return undefined;
  const n = Number.parseFloat(m[1] as string);
  if (Number.isNaN(n)) return undefined;
  // 百分比要有参照物（父形状的尺寸），这一层没有 —— 留空让它变成零尺寸占位，
  // 而不是一个编出来的错尺寸
  const unit = (m[2] as string).toLowerCase();
  if (unit === '%') return undefined;
  return (CSS_UNITS[unit] ?? pxToTwips)(n);
}

// ── 小工具 ────────────────────────────────────────────────────────────────────

function nonEmpty(s: string | undefined): string | undefined {
  return s === undefined || s === '' ? undefined : s;
}

/** 深度优先找第一个同名后代。`a:blip` 的容器有好几种写法，按固定路径取会漏 */
function descendant(root: XmlElement, name: string): XmlElement | undefined {
  for (const el of children(root)) {
    if (el.name === name) return el;
    const found = descendant(el, name);
    if (found !== undefined) return found;
  }
  return undefined;
}

function descendants(root: XmlElement, name: string): XmlElement[] {
  const out: XmlElement[] = [];
  const walk = (el: XmlElement): void => {
    for (const c of children(el)) {
      if (c.name === name) out.push(c);
      walk(c);
    }
  };
  walk(root);
  return out;
}
