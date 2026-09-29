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
 * 图片都走这条路 —— 公文的红头与印章尤其常见。只解析 `style` 的外框与定位
 * 与 `v:imagedata@r:id`，够把图画在对的地方；`v:line` / `v:rect` 这些真的 VML 图形
 * 仍然只得到一个占位框。
 *
 * `position:absolute` 的形状是**浮动**的，折成与 `wp:anchor` 同一份 `DrawingAnchor`
 * （`vmlAnchor`）。原来一律当内嵌：WPS 从 PDF 转出来的文件把页脚里的每一小段字都放进
 * 绝对定位的 VML 文本框，当内嵌就是一串 40pt 宽的框挤在页脚那一行里。
 *
 * ## 文本框（`wps:txbx` / `v:textbox`）
 *
 * 内容是一列块（`w:txbxContent`），交给调用方解析（`DrawingParseOptions.textBox`，
 * 块的解析在 parse-body.ts，这里 import 它会成环），对象上只留 id、内边距与纵向对齐。
 * 外观只认矩形的填充与轮廓（`ShapeStyle`）—— 文本框就是矩形，别的几何仍是非目标。
 * `mc:AlternateContent` 里 Choice（wps）与 Fallback（VML）各有一份同样的内容，
 * 调用方只收 Choice，所以一个文本框只摊出一份。
 */
import { emuToTwips, inchToTwips, mmToTwips, ptToTwips, pxToTwips, type Twips } from '@uw/core';
import type { XmlElement } from '@uw/ooxml';
import { attr, child, children, textContent } from '@uw/ooxml';
import type { AnchorPos, DrawingAnchor, ImageRef, NodeId, ObjectContent, ShapeStyle } from './nodes.ts';
import { attrInt, attrOnOff, put } from './xml-values.ts';

/** 解析对象时要从外面借的两样 */
export interface DrawingParseOptions {
  /**
   * `w:txbxContent` → 摊进 `LoadedDocument.textBoxes` 的 id。不给 = 不认文本框
   * （只剩外框，照旧画占位）—— 单测与只要外框的调用方不必关心内容怎么解析
   */
  textBox?: (content: XmlElement) => NodeId;
  /** 主题配色（`Theme.colors`），`a:schemeClr` 换算成 RGB 用。缺席时退到 Office 默认主题 */
  themeColors?: Readonly<Record<string, string>>;
}

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
export function parseDrawing(
  drawing: XmlElement,
  idPrefix: string,
  opts: DrawingParseOptions = {},
): ObjectContent {
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
  drawingTextBox(frame, opts, out);
  return out;
}

/**
 * `w:pict` → 一个对象片段（VML）。
 *
 * 尺寸藏在 `style="width:100pt;height:50pt"` 里，单位可以是 pt / px / in / cm / mm ——
 * CSS 的默认单位是 **px 不是 pt**（VML 照 CSS 写），漏掉这条会让所有不带单位的图
 * 小一截（1px = 0.75pt）。
 */
export function parsePict(pict: XmlElement, idPrefix: string, opts: DrawingParseOptions = {}): ObjectContent {
  // 外框与定位看**最外层**那个形状（组合的话就是组合本身）：组合里的子形状的 style
  // 用的是组合自己的坐标系（`coordsize`），不是 pt
  const shape = findShape(pict);
  const root = vmlRoot(pict, shape);
  const style = parseStyle(attr(root ?? shape ?? pict, 'style'));
  const out: ObjectContent = {
    kind: 'object',
    objectKind: 'picture',
    width: cssLength(style.width) ?? 0,
    height: cssLength(style.height) ?? 0,
  };
  if (root !== undefined) put(out, 'anchor', vmlAnchor(root, style));
  // 文本框常是 `v:rect` / `v:roundrect`，不是 `v:shape` —— 不能等找到图片形状才认
  if (root !== undefined) vmlTextBox(root, style, opts, out);
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
  const side = wrapTextOf(anchor);
  if (side !== undefined) out.wrapText = side;
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

const WRAP_TEXT = new Set(['bothSides', 'left', 'right', 'largest']);

/** 只有方形 / 紧密型 / 穿越型带 `@wrapText`；写错的值当没写（按规范默认的 bothSides） */
function wrapTextOf(anchor: XmlElement): DrawingAnchor['wrapText'] {
  for (const el of children(anchor)) {
    if (el.name !== 'wp:wrapSquare' && el.name !== 'wp:wrapTight' && el.name !== 'wp:wrapThrough') continue;
    const v = attr(el, 'wrapText');
    return v !== undefined && WRAP_TEXT.has(v) ? (v as DrawingAnchor['wrapText']) : undefined;
  }
  return undefined;
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

// ── 文本框 ────────────────────────────────────────────────────────────────────

/** `wps:bodyPr` 的内边距缺省值（EMU）：左右 0.1 英寸、上下 0.05 英寸 */
const BODY_INSET_EMU = { l: 91440, t: 45720, r: 91440, b: 45720 };

/**
 * DrawingML 文本框：`a:graphicData/wps:wsp/wps:txbx/w:txbxContent`。
 *
 * 只认 `wps:wsp` 直接挂在 graphicData 下的那种 —— 组合（`wpg:wgp`）与画布（`wpc:wpc`）
 * 里的子形状各有自己的坐标系，那是组合的几何，还是非目标
 */
function drawingTextBox(frame: XmlElement, opts: DrawingParseOptions, out: ObjectContent): void {
  if (opts.textBox === undefined) return;
  const graphic = child(frame, 'a:graphic');
  const data = graphic && child(graphic, 'a:graphicData');
  const wsp = data && child(data, 'wps:wsp');
  const txbx = wsp && child(wsp, 'wps:txbx');
  const content = txbx && child(txbx, 'w:txbxContent');
  if (wsp === undefined || content === undefined) return;

  const bodyPr = child(wsp, 'wps:bodyPr');
  const inset = (name: keyof typeof BODY_INSET_EMU): Twips =>
    emuToTwips(attrInt(bodyPr, `${name}Ins`) ?? BODY_INSET_EMU[name]);
  const anchor = bodyPr === undefined ? undefined : attr(bodyPr, 'anchor');
  out.textBox = {
    id: opts.textBox(content),
    inset: { left: inset('l'), top: inset('t'), right: inset('r'), bottom: inset('b') },
    vAlign: anchor === 'ctr' ? 'center' : anchor === 'b' ? 'bottom' : 'top',
  };
  put(out, 'shape', drawingShapeStyle(wsp, opts.themeColors));
}

/**
 * `wps:spPr` 的填充与轮廓；spPr 没写的退到 `wps:style` 的 `a:fillRef` / `a:lnRef`
 * （形状样式：「这一类形状默认用主题的第几号填充」，颜色写在引用里面）。
 *
 * 渐变 / 图案 / 图片填充不认（不画底色）—— 画错一个底色会盖住字，不画只是少了装饰
 */
function drawingShapeStyle(
  wsp: XmlElement,
  themeColors: DrawingParseOptions['themeColors'],
): ShapeStyle | undefined {
  const spPr = child(wsp, 'wps:spPr');
  const style = child(wsp, 'wps:style');
  const out: ShapeStyle = {};

  const fill = spPr === undefined ? undefined : fillOf(spPr);
  if (fill === undefined) {
    // spPr 一个填充都没写：看形状样式。idx 0 是「无填充」
    const ref = style && child(style, 'a:fillRef');
    if (ref !== undefined && (attrInt(ref, 'idx') ?? 0) > 0) put(out, 'fill', drawingColor(ref, themeColors));
  } else if (fill !== 'none') {
    put(out, 'fill', drawingColor(fill, themeColors));
  }

  const ln = spPr === undefined ? undefined : child(spPr, 'a:ln');
  const lnFill = ln === undefined ? undefined : fillOf(ln);
  const lnRef = style && child(style, 'a:lnRef');
  const refIdx = lnRef === undefined ? 0 : (attrInt(lnRef, 'idx') ?? 0);
  // 线宽：写了用写的；没写按样式引用的那一档（Office 主题的线型表 0.5 / 1 / 1.5pt）
  const width = emuToTwips(attrInt(ln, 'w') ?? THEME_LINE_WIDTHS_EMU[refIdx - 1] ?? DEFAULT_LINE_WIDTH_EMU);
  const color =
    lnFill === undefined
      ? refIdx > 0 && lnRef !== undefined
        ? drawingColor(lnRef, themeColors)
        : undefined
      : lnFill === 'none'
        ? undefined
        : drawingColor(lnFill, themeColors);
  if (color !== undefined && width > 0) out.stroke = { color, width };

  return out.fill === undefined && out.stroke === undefined ? undefined : out;
}

/**
 * Office 主题线型表（`a:lnStyleLst`）的三档线宽。我们不解析主题的这一段 ——
 * 自带主题从 2007 到 365 这三档没变过，换主题改线宽的文件罕见到不值得多解析一张表
 */
const THEME_LINE_WIDTHS_EMU = [6350, 12700, 19050];
/** spPr 写了线色却没写线宽：DrawingML 的缺省 0.75pt */
const DEFAULT_LINE_WIDTH_EMU = 9525;

/** 一个填充容器里写的是哪种填充：`none`、`a:solidFill` 元素，或者没写（undefined） */
function fillOf(parent: XmlElement): XmlElement | 'none' | undefined {
  for (const el of children(parent)) {
    if (el.name === 'a:noFill') return 'none';
    if (el.name === 'a:solidFill') return el;
    // 认不出的填充当「不画」，理由见 drawingShapeStyle
    if (el.name === 'a:gradFill' || el.name === 'a:pattFill' || el.name === 'a:blipFill') return 'none';
  }
  return undefined;
}

/** Office 2013 起的默认主题配色，文档没带主题（或主题里没配色）时拿它换算 `a:schemeClr` */
const DEFAULT_THEME_COLORS: Readonly<Record<string, string>> = {
  dk1: '000000',
  lt1: 'FFFFFF',
  dk2: '44546A',
  lt2: 'E7E6E6',
  accent1: '4472C4',
  accent2: 'ED7D31',
  accent3: 'A5A5A5',
  accent4: 'FFC000',
  accent5: '5B9BD5',
  accent6: '70AD47',
  hlink: '0563C1',
  folHlink: '954F72',
};

/** `bg1` / `tx1` 这类别名：Word 的颜色映射（`w:clrSchemeMapping`）默认就是这一套 */
const SCHEME_ALIASES: Readonly<Record<string, string>> = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };

const PRESET_COLORS: Readonly<Record<string, string>> = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  green: '008000',
  blue: '0000FF',
  yellow: 'FFFF00',
  gray: '808080',
  grey: '808080',
};

/**
 * 颜色元素（`a:srgbClr` / `a:schemeClr` / `a:prstClr` / `a:sysClr`）→ RRGGBB。
 * 修饰只认 `lumMod` / `lumOff`（Word 调色板上「深色 25%」「浅色 40%」就是这两个）
 */
function drawingColor(
  parent: XmlElement,
  themeColors: DrawingParseOptions['themeColors'],
): string | undefined {
  for (const el of children(parent)) {
    const val = attr(el, 'val');
    let rgb: string | undefined;
    switch (el.name) {
      case 'a:srgbClr':
        rgb = val;
        break;
      case 'a:schemeClr': {
        const slot = val === undefined ? undefined : (SCHEME_ALIASES[val] ?? val);
        rgb = slot === undefined ? undefined : (themeColors?.[slot] ?? DEFAULT_THEME_COLORS[slot]);
        break;
      }
      case 'a:prstClr':
        rgb = val === undefined ? undefined : PRESET_COLORS[val];
        break;
      case 'a:sysClr':
        rgb = attr(el, 'lastClr');
        break;
      default:
        continue;
    }
    if (rgb === undefined || !HEX6.test(rgb)) return undefined;
    return applyLum(rgb.toUpperCase(), el);
  }
  return undefined;
}

const HEX6 = /^[0-9A-Fa-f]{6}$/;

function applyLum(rgb: string, color: XmlElement): string {
  const mod = attrInt(child(color, 'a:lumMod'), 'val');
  const off = attrInt(child(color, 'a:lumOff'), 'val');
  if (mod === undefined && off === undefined) return rgb;
  const [h, s, l] = rgbToHsl(rgb);
  const lum = Math.min(1, Math.max(0, l * ((mod ?? 100000) / 100000) + (off ?? 0) / 100000));
  return hslToRgb(h, s, lum);
}

function rgbToHsl(rgb: string): [number, number, number] {
  const [r, g, b] = [0, 2, 4].map((i) => Number.parseInt(rgb.slice(i, i + 2), 16) / 255) as [
    number,
    number,
    number,
  ];
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}

function hslToRgb(h: number, s: number, l: number): string {
  const hue = (p: number, q: number, t0: number): number => {
    const t = t0 < 0 ? t0 + 1 : t0 > 1 ? t0 - 1 : t0;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  const channels =
    s === 0
      ? [l, l, l]
      : (() => {
          const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
          const p = 2 * l - q;
          return [hue(p, q, h + 1 / 3), hue(p, q, h), hue(p, q, h - 1 / 3)];
        })();
  return channels
    .map((c) =>
      Math.round(c * 255)
        .toString(16)
        .padStart(2, '0'),
    )
    .join('')
    .toUpperCase();
}

/** `v:textbox@inset` 的缺省值：与 DrawingML 相同，左右 0.1 英寸、上下 0.05 英寸 */
const VML_INSET_DEFAULT = [inchToTwips(0.1), inchToTwips(0.05), inchToTwips(0.1), inchToTwips(0.05)] as const;

/** VML 文本框：形状直下的 `v:textbox/w:txbxContent`。组合里的不认（见 `drawingTextBox`） */
function vmlTextBox(
  root: XmlElement,
  style: Record<string, string>,
  opts: DrawingParseOptions,
  out: ObjectContent,
): void {
  if (opts.textBox === undefined) return;
  const textbox = child(root, 'v:textbox');
  const content = textbox && child(textbox, 'w:txbxContent');
  if (textbox === undefined || content === undefined) return;

  // inset="l,t,r,b"，任何一项可以空着（空着的取缺省）
  const parts = (attr(textbox, 'inset') ?? '').split(',');
  const inset = (i: 0 | 1 | 2 | 3): Twips => cssLength(nonEmpty(parts[i]?.trim())) ?? VML_INSET_DEFAULT[i];
  const anchor = style['v-text-anchor'] ?? '';
  out.textBox = {
    id: opts.textBox(content),
    inset: { left: inset(0), top: inset(1), right: inset(2), bottom: inset(3) },
    vAlign: anchor.startsWith('middle') ? 'center' : anchor.startsWith('bottom') ? 'bottom' : 'top',
  };
  put(out, 'shape', vmlShapeStyle(root));
}

/**
 * VML 的填充与轮廓。**缺省是填白、描黑 0.75pt**（VML 规范的默认值）——
 * 所以 `filled="f"` / `stroked="f"` 要显式认；`v:fill` / `v:stroke` 子元素的 `on="f"` 同样关掉
 */
function vmlShapeStyle(shape: XmlElement): ShapeStyle | undefined {
  const out: ShapeStyle = {};
  const fillEl = child(shape, 'v:fill');
  const strokeEl = child(shape, 'v:stroke');
  const filled = vmlBool(attr(shape, 'filled')) ?? true;
  if (filled && vmlBool(fillEl && attr(fillEl, 'on')) !== false) {
    put(out, 'fill', vmlColor(attr(shape, 'fillcolor') ?? (fillEl && attr(fillEl, 'color')) ?? 'white'));
  }
  const stroked = vmlBool(attr(shape, 'stroked')) ?? true;
  if (stroked && vmlBool(strokeEl && attr(strokeEl, 'on')) !== false) {
    const color = vmlColor((strokeEl && attr(strokeEl, 'color')) ?? attr(shape, 'strokecolor') ?? 'black');
    const width = cssLength(
      (strokeEl && attr(strokeEl, 'weight')) ?? attr(shape, 'strokeweight') ?? '0.75pt',
      ptToTwips,
    );
    if (color !== undefined && width !== undefined && width > 0) out.stroke = { color, width };
  }
  return out.fill === undefined && out.stroke === undefined ? undefined : out;
}

/** VML 的布尔：`t` / `true` / `f` / `false`（**不是** ST_OnOff —— 那一套里 `f` 算真） */
function vmlBool(v: string | undefined): boolean | undefined {
  if (v === undefined) return undefined;
  const s = v.trim().toLowerCase();
  if (s === 't' || s === 'true' || s === '1') return true;
  if (s === 'f' || s === 'false' || s === '0') return false;
  return undefined;
}

/**
 * VML 颜色：`#rrggbb` / `#rgb` / 颜色名，后面可能跟着 ` [3213]`（调色板序号，忽略）。
 * 系统色（`window` / `windowText`）按白底黑字取
 */
function vmlColor(value: string): string | undefined {
  const v = (value.trim().split(/\s+/)[0] ?? '').toLowerCase();
  if (/^#[0-9a-f]{6}$/.test(v)) return v.slice(1).toUpperCase();
  if (/^#[0-9a-f]{3}$/.test(v)) {
    return [...v.slice(1)]
      .map((c) => c + c)
      .join('')
      .toUpperCase();
  }
  if (v === 'window') return 'FFFFFF';
  if (v === 'windowtext') return '000000';
  return PRESET_COLORS[v];
}

// ── VML 的定位 ────────────────────────────────────────────────────────────────

/** `mso-position-horizontal-relative` → `wp:positionH@relativeFrom`。缺省 `text`（= 栏） */
const VML_H_RELATIVE: Readonly<Record<string, string>> = {
  page: 'page',
  margin: 'margin',
  text: 'column',
  char: 'character',
  'left-margin-area': 'leftMargin',
  'right-margin-area': 'rightMargin',
  'inner-margin-area': 'insideMargin',
  'outer-margin-area': 'outsideMargin',
};

/** `mso-position-vertical-relative` → `wp:positionV@relativeFrom`。缺省 `text`（= 段落） */
const VML_V_RELATIVE: Readonly<Record<string, string>> = {
  page: 'page',
  margin: 'margin',
  text: 'paragraph',
  line: 'line',
  'top-margin-area': 'topMargin',
  'bottom-margin-area': 'bottomMargin',
  'inner-margin-area': 'insideMargin',
  'outer-margin-area': 'outsideMargin',
};

const VML_WRAP: Readonly<Record<string, DrawingAnchor['wrap']>> = {
  none: 'none',
  square: 'square',
  tight: 'tight',
  through: 'through',
  topAndBottom: 'topAndBottom',
};

const VML_WRAP_SIDE: Readonly<Record<string, NonNullable<DrawingAnchor['wrapText']>>> = {
  both: 'bothSides',
  left: 'left',
  right: 'right',
  largest: 'largest',
};

/**
 * `position:absolute` 的 VML 形状 → 与 `wp:anchor` 同一份定位。不是绝对定位的是内嵌，返回 undefined。
 *
 * 三处与 DrawingML 写法不同：
 * ① 偏移是 `left` + `margin-left`（Word 写 `left:0` 加 `margin-left`，别的生成器反过来，两个都加）；
 * ② 环绕在子元素 `w10:wrap` 上，**没有它就是「浮于 / 衬于文字」**（不绕）；
 * ③ 衬于文字下方写成**负的 `z-index`**，不是一个单独的开关
 */
function vmlAnchor(shape: XmlElement, style: Record<string, string>): DrawingAnchor | undefined {
  if (style.position !== 'absolute') return undefined;
  const z = Number.parseInt(style['z-index'] ?? '0', 10);
  const wrapEl = child(shape, 'w10:wrap');
  const wrapType = wrapEl === undefined ? undefined : attr(wrapEl, 'type');
  const side = wrapEl === undefined ? undefined : attr(wrapEl, 'side');
  const dist = (name: string, fallback: Twips): Twips =>
    cssLength(style[`mso-wrap-distance-${name}`]) ?? fallback;
  const out: DrawingAnchor = {
    wrap: (wrapType === undefined ? undefined : VML_WRAP[wrapType]) ?? 'none',
    behindDoc: z < 0,
    z: Number.isNaN(z) ? 0 : Math.abs(z),
    h: vmlAxis(style, 'horizontal', 'left', VML_H_RELATIVE, 'column'),
    v: vmlAxis(style, 'vertical', 'top', VML_V_RELATIVE, 'paragraph'),
    // VML 规范的缺省环绕距离：左右 9pt、上下 0
    dist: {
      top: dist('top', 0),
      bottom: dist('bottom', 0),
      left: dist('left', 180),
      right: dist('right', 180),
    },
  };
  const wrapText = side === undefined ? undefined : VML_WRAP_SIDE[side];
  if (wrapText !== undefined) out.wrapText = wrapText;
  return out;
}

function vmlAxis(
  style: Record<string, string>,
  axis: 'horizontal' | 'vertical',
  edge: 'left' | 'top',
  relatives: Readonly<Record<string, string>>,
  fallback: string,
): AnchorPos {
  const rel = style[`mso-position-${axis}-relative`];
  const pos: AnchorPos = { relativeFrom: (rel === undefined ? undefined : relatives[rel]) ?? fallback };
  const align = style[`mso-position-${axis}`];
  if (align !== undefined && align !== 'absolute') {
    pos.align = align;
    return pos;
  }
  pos.offset = (cssLength(style[edge]) ?? 0) + (cssLength(style[`margin-${edge}`]) ?? 0);
  return pos;
}

// ── VML 的外框与图片 ──────────────────────────────────────────────────────────

/** VML 里能当「一个对象」的元素（`v:shapetype` 只是定义，不画） */
const VML_ROOTS = new Set([
  'v:shape',
  'v:rect',
  'v:roundrect',
  'v:oval',
  'v:line',
  'v:polyline',
  'v:arc',
  'v:curve',
  'v:group',
  'v:image',
]);

/**
 * `w:pict` / `w:object` 直下的那个形状 —— 外框与定位都看它。优先装着图片的那一个
 * （同一个 pict 里排在前面的装饰性形状不能把印章顶掉），没有图片就取第一个
 */
function vmlRoot(pict: XmlElement, withImage: XmlElement | undefined): XmlElement | undefined {
  const roots = children(pict).filter((el) => VML_ROOTS.has(el.name));
  const holds = (el: XmlElement): boolean => el === withImage || children(el).some(holds);
  return (withImage === undefined ? undefined : roots.find(holds)) ?? roots[0];
}

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

/**
 * 认不出的单位当 px —— CSS 的默认单位就是它，VML 里不带单位的数字也是 px。
 * 例外是描边粗细（`strokeweight`），不带单位时按 pt（`bare`）
 */
function cssLength(value: string | undefined, bare: (n: number) => Twips = pxToTwips): Twips | undefined {
  if (value === undefined) return undefined;
  const m = /^(-?[\d.]+)\s*([a-z%]*)$/i.exec(value.trim());
  if (m === null) return undefined;
  const n = Number.parseFloat(m[1] as string);
  if (Number.isNaN(n)) return undefined;
  // 百分比要有参照物（父形状的尺寸），这一层没有 —— 留空让它变成零尺寸占位，
  // 而不是一个编出来的错尺寸
  const unit = (m[2] as string).toLowerCase();
  if (unit === '%') return undefined;
  if (unit === '') return bare(n);
  return (CSS_UNITS[unit] ?? pxToTwips)(n);
}

// ── 小工具 ────────────────────────────────────────────────────────────────────

function nonEmpty(s: string | undefined): string | undefined {
  return s === undefined || s === '' ? undefined : s;
}

/**
 * 深度优先找第一个同名后代。`a:blip` 的容器有好几种写法，按固定路径取会漏。
 *
 * **不钻进文本框的内容**（`w:txbxContent`）：那里面是另一列段落，框里插的图有自己的
 * `w:drawing`。钻进去的话框里的第一张图会被当成整个框的填充，画成一张盖满框的图
 */
function descendant(root: XmlElement, name: string): XmlElement | undefined {
  for (const el of children(root)) {
    if (el.name === name) return el;
    if (el.name === 'w:txbxContent') continue;
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
      // 同 `descendant`：文本框里的形状是框的内容，不是这个 pict 的
      if (c.name !== 'w:txbxContent') walk(c);
    }
  };
  walk(root);
  return out;
}
