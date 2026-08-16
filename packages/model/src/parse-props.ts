/**
 * `w:pPr` / `w:rPr` → 部分属性对象。
 *
 * 这一层只做**翻译**，不做任何合并、不填任何默认值 —— 缺席就是缺席，
 * 「这层没意见」的信息必须原样传给级联（cascade.ts），否则继承关系就断了。
 *
 * 单位在这里一次性转成 twips（原则 1.3）。哪些**不能**转、为什么，见 props.ts 的注释。
 */
import { halfPtToTwips } from '@uw/core';
import type { XmlElement } from '@uw/ooxml';
import { attr, child, children } from '@uw/ooxml';
import type {
  Indent,
  Justification,
  LineRule,
  NumberingRef,
  ParagraphSpacing,
  ParaProps,
  RunFonts,
  RunFontThemes,
  RunProps,
  TabStop,
} from './props.ts';
import { attrInt, attrOnOff, enumVal, intVal, onOff, put, valOf } from './xml-values.ts';

const VERT_ALIGNS = ['baseline', 'superscript', 'subscript'] as const;
const LINE_RULES = ['auto', 'exact', 'atLeast'] as const;
const HINTS = ['default', 'eastAsia', 'cs'] as const;
const TAB_ALIGNMENTS = ['left', 'center', 'right', 'decimal', 'bar', 'clear'] as const;
const TAB_LEADERS = ['none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot'] as const;

/**
 * `w:jc` 的取值。`start` / `end` 是较新的写法，等价于 left / right ——
 * 我们不做 RTL（非目标），所以直接映射掉，不必带着两套名字进布局层。
 */
export function parseJustification(v: string | undefined): Justification | undefined {
  if (v === 'start') return 'left';
  if (v === 'end') return 'right';
  return enumVal(v, ['left', 'center', 'right', 'both', 'distribute']);
}

export function parseRunProps(rPr: XmlElement | undefined): RunProps {
  const out: RunProps = {};
  if (rPr === undefined) return out;

  put(out, 'styleId', valOf(rPr, 'w:rStyle'));

  const rFonts = child(rPr, 'w:rFonts');
  if (rFonts !== undefined) {
    const fonts: RunFonts = {};
    put(fonts, 'ascii', attr(rFonts, 'w:ascii'));
    put(fonts, 'hAnsi', attr(rFonts, 'w:hAnsi'));
    put(fonts, 'eastAsia', attr(rFonts, 'w:eastAsia'));
    put(fonts, 'cs', attr(rFonts, 'w:cs'));
    put(fonts, 'hint', enumVal(attr(rFonts, 'w:hint'), HINTS));
    if (Object.keys(fonts).length > 0) out.fonts = fonts;

    // 主题字体引用单独存：级联时它和显式字体名同层竞争，展开成真实名字是最后一步
    const themes: RunFontThemes = {};
    put(themes, 'ascii', attr(rFonts, 'w:asciiTheme'));
    put(themes, 'hAnsi', attr(rFonts, 'w:hAnsiTheme'));
    put(themes, 'eastAsia', attr(rFonts, 'w:eastAsiaTheme'));
    // 注意大小写：这个属性是 `w:cstheme`，不是 `w:csTheme`。规范就这么不一致
    put(themes, 'cs', attr(rFonts, 'w:cstheme'));
    if (Object.keys(themes).length > 0) out.fontThemes = themes;
  }

  put(out, 'bold', onOff(rPr, 'w:b'));
  put(out, 'boldCs', onOff(rPr, 'w:bCs'));
  put(out, 'italic', onOff(rPr, 'w:i'));
  put(out, 'italicCs', onOff(rPr, 'w:iCs'));
  put(out, 'caps', onOff(rPr, 'w:caps'));
  put(out, 'smallCaps', onOff(rPr, 'w:smallCaps'));
  put(out, 'strike', onOff(rPr, 'w:strike'));
  put(out, 'doubleStrike', onOff(rPr, 'w:dstrike'));
  put(out, 'hidden', onOff(rPr, 'w:vanish'));
  put(out, 'snapToGrid', onOff(rPr, 'w:snapToGrid'));

  // w:sz / w:szCs / w:position / w:kern 都是半磅
  put(out, 'size', halfPt(intVal(rPr, 'w:sz')));
  put(out, 'sizeCs', halfPt(intVal(rPr, 'w:szCs')));
  put(out, 'position', halfPt(intVal(rPr, 'w:position')));
  put(out, 'kerning', halfPt(intVal(rPr, 'w:kern')));
  // w:spacing（字符间距）本来就是 twips，不用转
  put(out, 'charSpacing', intVal(rPr, 'w:spacing'));
  put(out, 'scale', intVal(rPr, 'w:w'));

  put(out, 'underline', valOf(rPr, 'w:u'));
  put(out, 'vertAlign', enumVal(valOf(rPr, 'w:vertAlign'), VERT_ALIGNS));

  const color = child(rPr, 'w:color');
  if (color !== undefined) {
    put(out, 'color', attr(color, 'w:val'));
    put(out, 'themeColor', attr(color, 'w:themeColor'));
  }

  const lang = child(rPr, 'w:lang');
  if (lang !== undefined) put(out, 'langEastAsia', attr(lang, 'w:eastAsia'));

  return out;
}

export function parseParaProps(pPr: XmlElement | undefined): ParaProps {
  const out: ParaProps = {};
  if (pPr === undefined) return out;

  put(out, 'styleId', valOf(pPr, 'w:pStyle'));
  put(out, 'justification', parseJustification(valOf(pPr, 'w:jc')));
  put(out, 'keepNext', onOff(pPr, 'w:keepNext'));
  put(out, 'keepLines', onOff(pPr, 'w:keepLines'));
  put(out, 'pageBreakBefore', onOff(pPr, 'w:pageBreakBefore'));
  put(out, 'widowControl', onOff(pPr, 'w:widowControl'));
  put(out, 'snapToGrid', onOff(pPr, 'w:snapToGrid'));
  put(out, 'autoSpaceDE', onOff(pPr, 'w:autoSpaceDE'));
  put(out, 'autoSpaceDN', onOff(pPr, 'w:autoSpaceDN'));
  put(out, 'overflowPunct', onOff(pPr, 'w:overflowPunct'));
  put(out, 'outlineLevel', intVal(pPr, 'w:outlineLvl'));

  const ind = child(pPr, 'w:ind');
  if (ind !== undefined) {
    const indent: Indent = {};
    // w:start / w:end 是 w:left / w:right 的新名字，两者都可能出现
    put(indent, 'left', attrInt(ind, 'w:left') ?? attrInt(ind, 'w:start'));
    put(indent, 'right', attrInt(ind, 'w:right') ?? attrInt(ind, 'w:end'));
    put(indent, 'firstLine', attrInt(ind, 'w:firstLine'));
    put(indent, 'hanging', attrInt(ind, 'w:hanging'));
    put(indent, 'leftChars', attrInt(ind, 'w:leftChars') ?? attrInt(ind, 'w:startChars'));
    put(indent, 'rightChars', attrInt(ind, 'w:rightChars') ?? attrInt(ind, 'w:endChars'));
    put(indent, 'firstLineChars', attrInt(ind, 'w:firstLineChars'));
    put(indent, 'hangingChars', attrInt(ind, 'w:hangingChars'));
    if (Object.keys(indent).length > 0) out.indent = indent;
  }

  const sp = child(pPr, 'w:spacing');
  if (sp !== undefined) {
    const spacing: ParagraphSpacing = {};
    put(spacing, 'before', attrInt(sp, 'w:before'));
    put(spacing, 'after', attrInt(sp, 'w:after'));
    put(spacing, 'beforeLines', attrInt(sp, 'w:beforeLines'));
    put(spacing, 'afterLines', attrInt(sp, 'w:afterLines'));
    // line 的刻度取决于 lineRule，故意不转单位 —— 见 props.ts 的 LineRule
    put(spacing, 'line', attrInt(sp, 'w:line'));
    put(spacing, 'lineRule', enumVal<LineRule>(attr(sp, 'w:lineRule'), LINE_RULES));
    put(spacing, 'beforeAutospacing', attrOnOff(sp, 'w:beforeAutospacing'));
    put(spacing, 'afterAutospacing', attrOnOff(sp, 'w:afterAutospacing'));
    if (Object.keys(spacing).length > 0) out.spacing = spacing;
  }

  const tabsEl = child(pPr, 'w:tabs');
  if (tabsEl !== undefined) {
    const tabs: TabStop[] = [];
    for (const t of children(tabsEl, 'w:tab')) {
      const pos = attrInt(t, 'w:pos');
      // w:pos 缺席的制表位没有意义（对齐到哪儿？），Word 也不写这种；跳过而不是当 0
      if (pos === undefined) continue;
      tabs.push({
        pos,
        // `start` / `end` 同 w:jc，是 left / right 的新名字
        alignment: parseTabAlignment(attr(t, 'w:val')) ?? 'left',
        leader: enumVal(attr(t, 'w:leader'), TAB_LEADERS) ?? 'none',
      });
    }
    if (tabs.length > 0) out.tabs = tabs;
  }

  const numPr = child(pPr, 'w:numPr');
  if (numPr !== undefined) {
    const numbering: NumberingRef = {};
    put(numbering, 'numId', intVal(numPr, 'w:numId'));
    put(numbering, 'level', intVal(numPr, 'w:ilvl'));
    if (Object.keys(numbering).length > 0) out.numbering = numbering;
  }

  // 段落标记自己的字符属性：空段落的行高全靠它
  const markRPr = child(pPr, 'w:rPr');
  if (markRPr !== undefined) out.markRunProps = parseRunProps(markRPr);

  return out;
}

function parseTabAlignment(v: string | undefined): TabStop['alignment'] | undefined {
  if (v === 'start') return 'left';
  if (v === 'end') return 'right';
  // `num` 是编号专用的左对齐制表位，排版行为与 left 一致，直接归并
  if (v === 'num') return 'left';
  return enumVal(v, TAB_ALIGNMENTS);
}

function halfPt(v: number | undefined): number | undefined {
  return v === undefined ? undefined : halfPtToTwips(v);
}
