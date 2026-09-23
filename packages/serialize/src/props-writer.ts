/**
 * `ParaProps` / `RunProps` → `w:pPr` / `w:rPr`，**以原元素为底打补丁**。
 *
 * 不从模型整个重新生成属性容器，因为模型只认识 Word 属性的一个子集：`w:pBdr`（段落边框）、
 * `w:shd`（底纹）、`w:framePr`、`w:rPrChange`（格式修订）、`w:eastAsianLayout`（双行合一）……
 * 这些我们不解析，但它们在用户的文档里。重新生成等于把它们全删了（原则 1.4）。
 *
 * 所以做法是**按字段组比对**：解析原元素得到「原来的值」，与模型上的值逐组比，
 * 只有变了的那一组才去动它对应的那一个子元素，而且动的时候也只改我们认识的那几个属性
 * （`patchAttrs`）。没变的组连同它的原始写法（`w:jc="start"`、`w:ind w:start=`）一字不改 ——
 * 解析是有损的（`start` 读成 `left`），只有「没变就不碰」才能保证不改写用户的原文。
 */
import { twipsToHalfPt } from '@uw/core';
import type { ParaProps, RunProps, TabStop } from '@uw/model';
import { parseParaProps, parseRunProps } from '@uw/model';
import type { XmlElement, XmlNode } from '@uw/ooxml';
import { child } from '@uw/ooxml';
import { el, insertOrdered, patchAttrs, same } from './xml-edit.ts';

/**
 * `CT_RPr` 的子元素顺序（ECMA-376 §17.3.2，`EG_RPrBase` + `w:rPrChange`）。
 * 前四个是段落标记专用的（`CT_ParaRPr`），出现在 `w:pPr/w:rPr` 里，放在同一张表不碍事。
 */
export const RPR_ORDER = [
  'w:ins',
  'w:del',
  'w:moveFrom',
  'w:moveTo',
  'w:rStyle',
  'w:rFonts',
  'w:b',
  'w:bCs',
  'w:i',
  'w:iCs',
  'w:caps',
  'w:smallCaps',
  'w:strike',
  'w:dstrike',
  'w:outline',
  'w:shadow',
  'w:emboss',
  'w:imprint',
  'w:noProof',
  'w:snapToGrid',
  'w:vanish',
  'w:webHidden',
  'w:color',
  'w:spacing',
  'w:w',
  'w:kern',
  'w:position',
  'w:sz',
  'w:szCs',
  'w:highlight',
  'w:u',
  'w:effect',
  'w:bdr',
  'w:shd',
  'w:fitText',
  'w:vertAlign',
  'w:rtl',
  'w:cs',
  'w:em',
  'w:lang',
  'w:eastAsianLayout',
  'w:specVanish',
  'w:oMath',
  'w:rPrChange',
] as const;

/** `CT_PPr` 的子元素顺序（§17.3.1，`CT_PPrBase` + `w:rPr` / `w:sectPr` / `w:pPrChange`） */
export const PPR_ORDER = [
  'w:pStyle',
  'w:keepNext',
  'w:keepLines',
  'w:pageBreakBefore',
  'w:framePr',
  'w:widowControl',
  'w:numPr',
  'w:suppressLineNumbers',
  'w:pBdr',
  'w:shd',
  'w:tabs',
  'w:suppressAutoHyphens',
  'w:kinsoku',
  'w:wordWrap',
  'w:overflowPunct',
  'w:topLinePunct',
  'w:autoSpaceDE',
  'w:autoSpaceDN',
  'w:bidi',
  'w:adjustRightInd',
  'w:snapToGrid',
  'w:spacing',
  'w:ind',
  'w:contextualSpacing',
  'w:mirrorIndents',
  'w:suppressOverlap',
  'w:jc',
  'w:textDirection',
  'w:textAlignment',
  'w:textboxTightWrap',
  'w:outlineLvl',
  'w:divId',
  'w:cnfStyle',
  'w:rPr',
  'w:sectPr',
  'w:pPrChange',
] as const;

/**
 * 一组字段 ↔ 一个子元素。`write` 拿到原元素（可能没有）与模型上的值，给出新元素；
 * 返回 undefined 表示这层不再写它（回到样式值）。
 */
interface Field<P> {
  keys: readonly (keyof P)[];
  name: string;
  write(old: XmlElement | undefined, props: P): XmlElement | undefined;
}

function patchContainer<P>(
  base: XmlElement | undefined,
  name: string,
  from: P,
  to: P,
  fields: readonly Field<P>[],
  order: readonly string[],
): XmlElement | undefined {
  let children: XmlNode[] = base?.children ?? [];
  for (const field of fields) {
    if (field.keys.every((k) => same(from[k], to[k]))) continue;
    const old = base === undefined ? undefined : child(base, field.name);
    const next = field.write(old, to);
    const at = old === undefined ? -1 : children.indexOf(old);
    if (at >= 0) {
      children = next === undefined ? children.toSpliced(at, 1) : children.toSpliced(at, 1, next);
    } else if (next !== undefined) {
      children = insertOrdered(children, next, order);
    }
  }
  if (children.length === 0 && Object.keys(base?.attrs ?? {}).length === 0) return undefined;
  return base === undefined ? el(name, {}, children) : { ...base, children };
}

// ── 取值的几种写法 ────────────────────────────────────────────────────────────

/** ST_OnOff：true 写光秃秃的元素（Word 自己的写法），false 写 `w:val="0"` —— 明确关掉与「没意见」是两回事 */
function onOff<P>(key: keyof P, name: string): Field<P> {
  return {
    keys: [key],
    name,
    write: (_old, p) => {
      const v = p[key];
      if (v === undefined) return undefined;
      return el(name, v ? {} : { 'w:val': '0' });
    },
  };
}

function val<P>(
  key: keyof P,
  name: string,
  format: (v: NonNullable<P[keyof P]>) => string = String,
): Field<P> {
  return {
    keys: [key],
    name,
    write: (old, p) => {
      const v = p[key];
      if (v === undefined || v === null) return undefined;
      return el(name, patchAttrs(old?.attrs ?? {}, ['w:val'], { 'w:val': format(v) }));
    },
  };
}

/** 模型里是 twips，文件里是半磅（`w:sz` / `w:kern` / `w:position`） */
const halfPt = (v: unknown): string => String(Math.round(twipsToHalfPt(v as number)));
const onOffAttr = (v: boolean | undefined): string | undefined =>
  v === undefined ? undefined : v ? '1' : '0';
const num = (v: number | undefined): string | undefined =>
  v === undefined ? undefined : String(Math.round(v));

// ── 字符属性 ──────────────────────────────────────────────────────────────────

const FONT_ATTRS = [
  'w:ascii',
  'w:hAnsi',
  'w:eastAsia',
  'w:cs',
  'w:hint',
  'w:asciiTheme',
  'w:hAnsiTheme',
  'w:eastAsiaTheme',
  'w:cstheme',
] as const;

const RUN_FIELDS: readonly Field<RunProps>[] = [
  val('styleId', 'w:rStyle'),
  {
    keys: ['fonts', 'fontThemes'],
    name: 'w:rFonts',
    write(old, p) {
      const f = p.fonts ?? {};
      const t = p.fontThemes ?? {};
      // 同一个桶既有显式名又有主题引用时**只写显式名**：Word 按规范让主题引用压过显式名
      // （§17.3.2.26「asciiTheme 在时 ascii 被忽略」），而我们的级联是反过来的（cascade.ts，
      // 显式名压过主题）。写两个等于让 Word 显示的字体与屏幕上的不一样 —— 改过的 run
      // 以模型为准，丢掉那个被压住的主题引用
      const theme = (name: string | undefined, ref: string | undefined) =>
        name !== undefined && name !== '' ? undefined : ref;
      const attrs = patchAttrs(old?.attrs ?? {}, FONT_ATTRS, {
        'w:ascii': f.ascii,
        'w:hAnsi': f.hAnsi,
        'w:eastAsia': f.eastAsia,
        'w:cs': f.cs,
        'w:hint': f.hint,
        'w:asciiTheme': theme(f.ascii, t.ascii),
        'w:hAnsiTheme': theme(f.hAnsi, t.hAnsi),
        'w:eastAsiaTheme': theme(f.eastAsia, t.eastAsia),
        'w:cstheme': theme(f.cs, t.cs),
      });
      return Object.keys(attrs).length === 0 ? undefined : el('w:rFonts', attrs);
    },
  },
  onOff('bold', 'w:b'),
  onOff('boldCs', 'w:bCs'),
  onOff('italic', 'w:i'),
  onOff('italicCs', 'w:iCs'),
  onOff('caps', 'w:caps'),
  onOff('smallCaps', 'w:smallCaps'),
  onOff('strike', 'w:strike'),
  onOff('doubleStrike', 'w:dstrike'),
  onOff('snapToGrid', 'w:snapToGrid'),
  onOff('hidden', 'w:vanish'),
  {
    keys: ['color', 'themeColor'],
    name: 'w:color',
    write(old, p) {
      // 颜色改了就丢掉主题色：Word 里主题色压过 w:val，而渲染层画的是 w:val（主题色未解析，
      // props.ts）—— 两个都留着，屏幕上是新颜色、Word 里还是旧的主题色
      const theme = p.color !== undefined && old?.attrs['w:val'] !== p.color ? undefined : p.themeColor;
      if (p.color === undefined && theme === undefined) return undefined;
      // 深浅（themeShade / themeTint）是修饰主题色的，主题色换了它们就不再成立
      const keepShade = theme !== undefined && old?.attrs['w:themeColor'] === theme;
      return el(
        'w:color',
        patchAttrs(old?.attrs ?? {}, ['w:val', 'w:themeColor', 'w:themeShade', 'w:themeTint'], {
          // w:val 是必填属性；只给了主题色时写 auto，颜色由主题色决定
          'w:val': p.color ?? 'auto',
          'w:themeColor': theme,
          'w:themeShade': keepShade ? old?.attrs['w:themeShade'] : undefined,
          'w:themeTint': keepShade ? old?.attrs['w:themeTint'] : undefined,
        }),
      );
    },
  },
  val('charSpacing', 'w:spacing', (v) => String(Math.round(v as number))),
  val('scale', 'w:w', (v) => String(Math.round(v as number))),
  val('kerning', 'w:kern', halfPt),
  val('position', 'w:position', halfPt),
  val('size', 'w:sz', halfPt),
  val('sizeCs', 'w:szCs', halfPt),
  val('underline', 'w:u'),
  val('vertAlign', 'w:vertAlign'),
  {
    keys: ['langEastAsia'],
    name: 'w:lang',
    write(old, p) {
      const attrs = patchAttrs(old?.attrs ?? {}, ['w:eastAsia'], { 'w:eastAsia': p.langEastAsia });
      return Object.keys(attrs).length === 0 ? undefined : el('w:lang', attrs);
    },
  },
];

/** `base` 是原来的 `w:rPr`（可能没有），`name` 区分 run 的与段落标记的（两者元素名相同，顺序表共用） */
export function patchRunProps(base: XmlElement | undefined, to: RunProps): XmlElement | undefined {
  return patchContainer(base, 'w:rPr', parseRunProps(base), to, RUN_FIELDS, RPR_ORDER);
}

// ── 段落属性 ──────────────────────────────────────────────────────────────────

const IND_ATTRS = [
  'w:left',
  'w:start',
  'w:right',
  'w:end',
  'w:firstLine',
  'w:hanging',
  'w:leftChars',
  'w:startChars',
  'w:rightChars',
  'w:endChars',
  'w:firstLineChars',
  'w:hangingChars',
] as const;

const SPACING_ATTRS = [
  'w:before',
  'w:after',
  'w:beforeLines',
  'w:afterLines',
  'w:line',
  'w:lineRule',
  'w:beforeAutospacing',
  'w:afterAutospacing',
] as const;

function tabElement(t: TabStop): XmlElement {
  const attrs: Record<string, string> = { 'w:val': t.alignment };
  if (t.leader !== 'none') attrs['w:leader'] = t.leader;
  attrs['w:pos'] = String(Math.round(t.pos));
  return el('w:tab', attrs);
}

const PARA_FIELDS: readonly Field<ParaProps>[] = [
  val('styleId', 'w:pStyle'),
  onOff('keepNext', 'w:keepNext'),
  onOff('keepLines', 'w:keepLines'),
  onOff('pageBreakBefore', 'w:pageBreakBefore'),
  onOff('widowControl', 'w:widowControl'),
  {
    keys: ['numbering'],
    name: 'w:numPr',
    write(old, p) {
      const n = p.numbering;
      if (n === undefined || (n.level === undefined && n.numId === undefined)) return undefined;
      // 我们只认 ilvl / numId，`w:numberingChange`（编号修订）之类的邻居留着
      let children: XmlNode[] = (old?.children ?? []).filter(
        (c) => c.kind !== 'element' || (c.name !== 'w:ilvl' && c.name !== 'w:numId'),
      );
      const order = ['w:ilvl', 'w:numId', 'w:numberingChange', 'w:ins'];
      if (n.level !== undefined)
        children = insertOrdered(children, el('w:ilvl', { 'w:val': String(n.level) }), order);
      if (n.numId !== undefined)
        children = insertOrdered(children, el('w:numId', { 'w:val': String(n.numId) }), order);
      return { ...(old ?? el('w:numPr')), children };
    },
  },
  {
    keys: ['tabs'],
    name: 'w:tabs',
    write: (_old, p) =>
      p.tabs === undefined || p.tabs.length === 0 ? undefined : el('w:tabs', {}, p.tabs.map(tabElement)),
  },
  onOff('overflowPunct', 'w:overflowPunct'),
  onOff('autoSpaceDE', 'w:autoSpaceDE'),
  onOff('autoSpaceDN', 'w:autoSpaceDN'),
  onOff('snapToGrid', 'w:snapToGrid'),
  {
    keys: ['spacing'],
    name: 'w:spacing',
    write(old, p) {
      const s = p.spacing ?? {};
      const attrs = patchAttrs(old?.attrs ?? {}, SPACING_ATTRS, {
        'w:before': num(s.before),
        'w:after': num(s.after),
        'w:beforeLines': num(s.beforeLines),
        'w:afterLines': num(s.afterLines),
        'w:line': num(s.line),
        'w:lineRule': s.lineRule,
        'w:beforeAutospacing': onOffAttr(s.beforeAutospacing),
        'w:afterAutospacing': onOffAttr(s.afterAutospacing),
      });
      return Object.keys(attrs).length === 0 ? undefined : el('w:spacing', attrs);
    },
  },
  {
    keys: ['indent'],
    name: 'w:ind',
    write(old, p) {
      const i = p.indent ?? {};
      // `w:start` / `w:end` 是新名字，解析时两种都认；写回统一用 left / right（过渡格式里两者等价），
      // 两个都留着会让旧值与新值打架
      const attrs = patchAttrs(old?.attrs ?? {}, IND_ATTRS, {
        'w:left': num(i.left),
        'w:right': num(i.right),
        'w:firstLine': num(i.firstLine),
        'w:hanging': num(i.hanging),
        'w:leftChars': num(i.leftChars),
        'w:rightChars': num(i.rightChars),
        'w:firstLineChars': num(i.firstLineChars),
        'w:hangingChars': num(i.hangingChars),
      });
      return Object.keys(attrs).length === 0 ? undefined : el('w:ind', attrs);
    },
  },
  val('justification', 'w:jc'),
  val('outlineLevel', 'w:outlineLvl'),
  {
    keys: ['markRunProps'],
    name: 'w:rPr',
    write: (old, p) => patchRunProps(old, p.markRunProps ?? {}),
  },
];

/**
 * 段落属性的补丁。`w:sectPr` 不归这里管：它跟着「这一段是不是本节最后一段」走，
 * 那是正文结构的事（body-writer.ts），这里原样带着、由调用方摘掉或补上。
 */
export function patchParaProps(base: XmlElement | undefined, to: ParaProps): XmlElement | undefined {
  return patchContainer(base, 'w:pPr', parseParaProps(base), to, PARA_FIELDS, PPR_ORDER);
}
