/**
 * 富文本剪贴板：复制把模型片段写成 HTML，粘贴把 HTML 读回「段落 → 带格式补丁的文字」。
 *
 * 复制与粘贴都读取样式原文，粘贴同时读取 Word 类样式；不走 CSSOM：Word 写的 `mso-fareast-font-family`
 * 这类私有属性 CSSOM 一律丢掉，而「汉字用仿宋、数字用 Times」恰恰只靠它们表达。
 * 复制照 Word 自己的 HTML 写法出（`font-family` + `mso-ascii/hansi/fareast-font-family`），
 * 于是我们复制的内容 Word 认得出字体，Word 复制的内容我们也认得出 —— 同一套解析。
 *
 * 粘贴分两种来源，取舍不同（Word 的「保留源格式 / 合并格式」）：
 * - **可信**：我们自己（`data-uw-fragment`）或 Word / WPS（`mso-` 属性、Office 命名空间）。
 *   带字体、字号、颜色与段落对齐，没写的开关（加粗 / 斜体…）按「没有」写 —— 否则粘进加粗的
 *   标题里，源里不加粗的字也跟着加粗，与复制时看到的不一致
 * - **网页**：只带加粗 / 斜体 / 下划线 / 删除线 / 上下标，且只加不减。网页的字体（Arial、
 *   微软雅黑）和 px 字号进了公文就是一段排版不对的字，度量包里多半还没有这款字体
 */

import { TWIP_PER_PT } from '@uw/core';
import { emphasisOf } from '@uw/layout';
import type {
  EmphasisMark,
  FragmentRun,
  Justification,
  RichFragment,
  RunPropsPatch,
  Shading,
} from '@uw/model';
import { HIGHLIGHT_COLORS, highlightFill, shadingFill } from '@uw/render-dom';
import { clipboardStyleReader } from './clipboard-styles.ts';

/** 粘贴进来的一段文字；补丁在光标处继承来的格式之上再改。 */
export interface PasteRun {
  text: string;
  patch: RunPropsPatch;
}
/** 粘贴进来的一段；`justification` 只在可信来源里读得到。 */
export interface PasteParagraph {
  justification?: Justification;
  runs: PasteRun[];
}

const CSS_ALIGN: Readonly<Record<Justification, string>> = {
  left: 'left',
  center: 'center',
  right: 'right',
  both: 'justify',
  distribute: 'justify',
};

const escapeText = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const escapeAttr = (s: string) => escapeText(s).replace(/"/g, '&quot;');
/** 字体名放进双引号括起的 style 属性里，只能用单引号括 —— 「Times New Roman」有空格。 */
const cssFont = (name: string) => `'${name.replace(/['\\]/g, '')}'`;

function runStyle(props: FragmentRun['props']): string {
  const { ascii, hAnsi, eastAsia } = props.fonts;
  const decorations = [
    ...(props.underline !== 'none' && props.underline !== '' ? ['underline'] : []),
    ...(props.strike || props.doubleStrike ? ['line-through'] : []),
  ];
  const style = [
    ...(ascii ? [`font-family:${cssFont(ascii)}`, `mso-ascii-font-family:${cssFont(ascii)}`] : []),
    ...(hAnsi ? [`mso-hansi-font-family:${cssFont(hAnsi)}`] : []),
    ...(eastAsia ? [`mso-fareast-font-family:${cssFont(eastAsia)}`] : []),
    `font-size:${props.size / TWIP_PER_PT}pt`,
    `font-weight:${props.bold ? 'bold' : 'normal'}`,
    `font-style:${props.italic ? 'italic' : 'normal'}`,
    `text-decoration:${decorations.length ? decorations.join(' ') : 'none'}`,
    // Word 的 HTML 用这个私有属性区分双线 / 波浪线；浏览器照样只画单线。
    ...(props.underline !== 'none' && props.underline !== 'single' && props.underline !== ''
      ? [`text-underline:${props.underline}`]
      : []),
    ...(/^[0-9a-f]{6}$/i.test(props.color) ? [`color:#${props.color}`] : []),
    ...(props.vertAlign === 'baseline'
      ? []
      : [`vertical-align:${props.vertAlign === 'superscript' ? 'super' : 'sub'}`]),
    ...backgroundStyle(props),
    ...emphasisStyle(props),
  ];
  return style.join(';');
}

/**
 * 着重号写成 CSS `text-emphasis`，上下与屏幕上一致（中文的点在字下）。浏览器照着画，粘回来也认它。
 * Word 自己的 HTML 怎么写着重号没有对过，粘进 Word 多半丢掉 —— 丢的只是这一样
 */
const CSS_EMPHASIS = { dot: 'filled dot', comma: 'filled sesame', circle: 'open circle' } as const;

function emphasisStyle(props: FragmentRun['props']): string[] {
  const emphasis = emphasisOf(props);
  if (emphasis === undefined) return [];
  return [
    `text-emphasis:${CSS_EMPHASIS[emphasis.mark]}`,
    `text-emphasis-position:${emphasis.below ? 'under' : 'over'} right`,
  ];
}

/** CSS 的 `text-emphasis(-style)` → `w:em`。上下不管：中文版 Word 的着重号写的就是 `dot`，上下由语言定 */
function readEmphasis(value: string): EmphasisMark | undefined {
  const v = value.toLowerCase();
  if (/\bnone\b/.test(v)) return 'none';
  if (v.includes('sesame')) return 'comma';
  if (v.includes('circle')) return 'circle';
  if (v.includes('dot')) return 'dot';
  return undefined;
}

/**
 * 高亮照 Word 的写法出 `background` + `mso-highlight`：浏览器只认前者，粘回来靠后者分辨
 * 「这是高亮」而不是一块底纹。两者都有时看得见的是高亮（与画法一致），底纹不再写。
 */
function backgroundStyle(props: FragmentRun['props']): string[] {
  const highlight = highlightFill(props.highlight);
  if (highlight !== undefined) return [`background:${highlight}`, `mso-highlight:${props.highlight}`];
  const shading = shadingFill(props.shading);
  return shading === undefined ? [] : [`background:${shading}`];
}

/** Word 写分页符的原样：粘进 Word 仍是分页符，浏览器与我们自己都认 `page-break-before`。 */
const PAGE_BREAK_HTML = '<br clear="all" style="mso-special-character:line-break;page-break-before:always">';
/** 显式标记真实软换行，避免段末或纯换行片段被当成网页编辑器的占位 br。 */
const SOFT_BREAK_HTML = '<br style="mso-special-character:line-break">';
/** Word 用空白 span 表示制表位；保留原始 tab，让 pre-wrap 下的浏览器也能显示。 */
const TAB_HTML = '<span style="mso-tab-count:1">\t</span>';

/**
 * 片段写成剪贴板 HTML。一段一个 `<p>`（含空段 —— 选到下一段开头时末尾那个空 `<p>`
 * 正是纯文本里的那个换行），`pre-wrap` 保住公文里「版    本」那种连续空格。
 */
export function fragmentToHtml(fragment: RichFragment): string {
  const body = fragment.paragraphs
    .map((p) => {
      const runs = p.runs
        .map(
          (r) =>
            `<span style="${escapeAttr(runStyle(r.props))}">${escapeText(r.text).replace(/\t/g, TAB_HTML).replace(/\n/g, SOFT_BREAK_HTML).replace(/\f/g, PAGE_BREAK_HTML)}</span>`,
        )
        .join('');
      return `<p style="margin:0;white-space:pre-wrap;text-align:${CSS_ALIGN[p.justification]}">${runs}</p>`;
    })
    .join('');
  return `<meta charset="utf-8"><div data-uw-fragment="1">${body}</div>`;
}

// ── 粘贴 ─────────────────────────────────────────────────────────────────────

/** 产生段落边界的元素；表格单元格也算 —— 与纯文本复制「单元格之间换行」一致。 */
const BLOCKS = new Set([
  'P',
  'DIV',
  'H1',
  'H2',
  'H3',
  'H4',
  'H5',
  'H6',
  'LI',
  'BLOCKQUOTE',
  'PRE',
  'TR',
  'TD',
  'TH',
  'DT',
  'DD',
  'SECTION',
  'ARTICLE',
  'HEADER',
  'FOOTER',
  'TABLE',
  'UL',
  'OL',
  'DL',
  'ADDRESS',
  'FIGURE',
]);
/** 写了就算一段的元素，哪怕是空的；`<div>` 这类纯容器空着不出段，免得嵌套几层就多出几个空段。 */
const EXPLICIT = new Set(['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI', 'PRE']);
const SKIP = new Set([
  'HEAD',
  'STYLE',
  'SCRIPT',
  'TITLE',
  'META',
  'LINK',
  'TEMPLATE',
  'IMG',
  'SVG',
  'OBJECT',
]);
/** CSS 的具名色里常见的那几个；Word 的 HTML 会写 `color:red`、`color:windowtext`。 */
const NAMED_COLORS: Readonly<Record<string, string>> = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  green: '008000',
  blue: '0000FF',
  yellow: 'FFFF00',
  gray: '808080',
  grey: '808080',
  silver: 'C0C0C0',
  maroon: '800000',
  navy: '000080',
  purple: '800080',
  teal: '008080',
  olive: '808000',
  lime: '00FF00',
  aqua: '00FFFF',
  fuchsia: 'FF00FF',
  orange: 'FFA500',
  windowtext: 'auto',
};

interface Inherited {
  bold: boolean;
  italic: boolean;
  underline: string | undefined;
  strike: boolean;
  vertAlign: 'baseline' | 'superscript' | 'subscript';
  color?: string | undefined;
  size?: number;
  ascii?: string | undefined;
  hAnsi?: string | undefined;
  eastAsia?: string | undefined;
  family?: string | undefined;
  /** 高亮的 `w:highlight` 名字；与 `shading` 互斥（后写的那个说了算） */
  highlight?: string | undefined;
  emphasis?: EmphasisMark | undefined;
  shading?: Shading | undefined;
  pre: boolean;
  /** Word 的显式空格与 pre 不同：不把 HTML 源码换行拆成段落。 */
  spaces: boolean;
  align?: Justification | undefined;
}

const firstFamily = (value: string) =>
  value
    .split(',')[0]
    ?.trim()
    .replace(/^['"]|['"]$/g, '')
    .trim() || undefined;

/** 只认 pt 与 px（px 按 CSS 的 96dpi 折 pt）；em / % 要知道父级字号，交给继承。 */
function parseSize(value: string): number | undefined {
  const m = /^([\d.]+)(pt|px)$/i.exec(value.trim());
  if (!m?.[1] || !m[2]) return undefined;
  const pt = Number(m[1]) * (m[2].toLowerCase() === 'px' ? 0.75 : 1);
  // 半磅是 w:sz 的精度，落在半磅上免得写出 Word 存不下的字号。
  return pt > 0 ? Math.round(pt * 2) * (TWIP_PER_PT / 2) : undefined;
}

function parseColor(value: string): string | undefined {
  const v = value.trim().toLowerCase();
  const hex = /^#([0-9a-f]{3}|[0-9a-f]{6})$/.exec(v)?.[1];
  if (hex) return (hex.length === 3 ? [...hex].map((c) => c + c).join('') : hex).toUpperCase();
  const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/.exec(v);
  if (rgb)
    return rgb
      .slice(1, 4)
      .map((n) => Math.min(255, Number(n)).toString(16).padStart(2, '0'))
      .join('')
      .toUpperCase();
  return NAMED_COLORS[v];
}

const CJK = '[\\u2e80-\\u9fff\\uf900-\\ufaff\\u3000-\\u303f\\uff00-\\uffef]';
const CJK_WRAP = new RegExp(`(${CJK})[ \\t]*[\\r\\n]+[ \\t]*(?=${CJK})`, 'g');

const CSS_JUSTIFY: Readonly<Record<string, Justification>> = {
  left: 'left',
  start: 'left',
  center: 'center',
  right: 'right',
  end: 'right',
  justify: 'both',
};

function inherit(el: Element, from: Inherited, style: Map<string, string>): Inherited {
  const tag = el.tagName.toUpperCase();
  const next: Inherited = { ...from };
  if (tag === 'B' || tag === 'STRONG' || /^H[1-6]$/.test(tag) || tag === 'TH') next.bold = true;
  if (tag === 'I' || tag === 'EM' || tag === 'CITE') next.italic = true;
  if (tag === 'U' || tag === 'INS') next.underline = 'single';
  if (tag === 'S' || tag === 'STRIKE' || tag === 'DEL') next.strike = true;
  if (tag === 'SUP') next.vertAlign = 'superscript';
  if (tag === 'SUB') next.vertAlign = 'subscript';
  if (tag === 'PRE') next.pre = true;
  const align = el.getAttribute('align');
  if (align && CSS_JUSTIFY[align.toLowerCase()]) next.align = CSS_JUSTIFY[align.toLowerCase()];
  const weight = style.get('font-weight');
  if (weight) next.bold = weight === 'bold' || weight === 'bolder' || Number(weight) >= 600;
  const fontStyle = style.get('font-style');
  if (fontStyle) next.italic = fontStyle === 'italic' || fontStyle === 'oblique';
  const decoration = style.get('text-decoration') ?? style.get('text-decoration-line');
  if (decoration) {
    next.underline = decoration.includes('underline') ? (next.underline ?? 'single') : undefined;
    next.strike = decoration.includes('line-through');
  }
  const underline = style.get('text-underline');
  if (underline && next.underline) next.underline = underline === 'none' ? undefined : underline;
  const vertical = style.get('vertical-align');
  if (vertical)
    next.vertAlign = vertical === 'super' ? 'superscript' : vertical === 'sub' ? 'subscript' : 'baseline';
  const color = style.get('color');
  const parsed = color && parseColor(color);
  if (parsed) next.color = parsed;
  const size = style.get('font-size');
  const twips = size && parseSize(size);
  if (twips) next.size = twips;
  const family = style.get('font-family');
  if (family) next.family = firstFamily(family);
  const ascii = style.get('mso-ascii-font-family');
  if (ascii) next.ascii = firstFamily(ascii);
  const hAnsi = style.get('mso-hansi-font-family');
  if (hAnsi) next.hAnsi = firstFamily(hAnsi);
  const eastAsia = style.get('mso-fareast-font-family');
  if (eastAsia) next.eastAsia = firstFamily(eastAsia);
  // 背景只从行内元素上读：块上的背景是段落底纹，铺到每个 run 上就成了字符底纹。
  // CSS 的 background 本不继承，这里顺着 span 往下带，是因为它包住的字确实都铺着这块底
  if (!BLOCKS.has(tag)) readBackground(style, next);
  const emphasis = style.get('text-emphasis-style') ?? style.get('text-emphasis');
  if (emphasis) next.emphasis = readEmphasis(emphasis) ?? next.emphasis;
  const whiteSpace = style.get('white-space');
  if (whiteSpace) next.pre = whiteSpace.startsWith('pre') || whiteSpace === 'break-spaces';
  const spacerun = style
    .get('mso-spacerun')
    ?.replace(/^['"]|['"]$/g, '')
    .toLowerCase();
  if (spacerun === 'yes' || spacerun === 'no') next.spaces = spacerun === 'yes';
  const textAlign = style.get('text-align');
  if (textAlign && CSS_JUSTIFY[textAlign.toLowerCase()]) next.align = CSS_JUSTIFY[textAlign.toLowerCase()];
  return next;
}

/** 高亮色值 → `w:highlight` 名字。Word 的 `mso-highlight` 写 CSS 色名（`lime` / `olive`），我们写规范名 */
const HIGHLIGHT_BY_HEX: ReadonlyMap<string, string> = new Map(
  Object.entries(HIGHLIGHT_COLORS).map(([name, hex]) => [hex.slice(1).toUpperCase(), name]),
);

/** 只在顶层空白处分词：RGB 内的空格不能拆色值，URL / 渐变 / 字符串里的颜色词也不能漏出来。 */
function backgroundTokens(value: string): string[] {
  const tokens: string[] = [];
  let quote = '';
  let depth = 0;
  let start = 0;
  for (let i = 0; i < value.length; i++) {
    const ch = value[i];
    if (ch === '\\') i++;
    else if (quote) {
      if (ch === quote) quote = '';
    } else if (ch === '"' || ch === "'") quote = ch;
    else if (ch === '(') depth++;
    else if (ch === ')') depth = Math.max(0, depth - 1);
    else if (depth === 0 && ch !== undefined && /\s/.test(ch)) {
      if (i > start) tokens.push(value.slice(start, i));
      start = i + 1;
    }
  }
  if (start < value.length) tokens.push(value.slice(start));
  return tokens;
}

function readBackground(style: Map<string, string>, next: Inherited): void {
  const mso = style.get('mso-highlight');
  if (mso !== undefined) {
    const name =
      mso === 'none' || mso in HIGHLIGHT_COLORS ? mso : HIGHLIGHT_BY_HEX.get(parseColor(mso) ?? '');
    if (name !== undefined) {
      next.highlight = name;
      // none 只关闭高亮：background 仍可能表达独立的字符底纹，不能一起丢掉。
      if (name !== 'none') {
        next.shading = undefined;
        return;
      }
    }
  }
  const background = style.get('background') ?? style.get('background-color');
  if (background === undefined) return;
  // `background` 简写可能带着图片、位置；只取第一个认得出的颜色词
  const fill = backgroundTokens(background)
    .map((word) => parseColor(word))
    .find((c) => c !== undefined);
  if (fill === undefined || fill === 'auto') {
    if (/^(none|transparent)$/i.test(background.trim())) {
      next.highlight = undefined;
      next.shading = undefined;
    }
    return;
  }
  next.highlight = undefined;
  next.shading = { pattern: 'clear', color: 'auto', fill };
}

function patchOf(state: Inherited, trusted: boolean): RunPropsPatch {
  if (!trusted)
    return {
      ...(state.bold ? { bold: true } : {}),
      ...(state.italic ? { italic: true } : {}),
      ...(state.underline ? { underline: 'single' } : {}),
      ...(state.strike ? { strike: true } : {}),
      ...(state.vertAlign !== 'baseline' ? { vertAlign: state.vertAlign } : {}),
    };
  // Word 只在一款字体上写 font-family 时（纯中文或纯西文的 run），它就是三个已支持桶共同的字体；
  // 写了 mso-* 的桶以 mso-* 为准。一个都没写就不动 —— 沿用光标处的字体，而不是清成样式值。
  const ascii = state.ascii ?? state.family;
  // Word 可只写扩展拉丁桶；不能依附在 ascii 的分支上，否则整桶会静默丢失。
  const hAnsi = state.hAnsi ?? ascii;
  const eastAsia = state.eastAsia ?? state.family;
  return {
    bold: state.bold,
    italic: state.italic,
    underline: state.underline ?? 'none',
    strike: state.strike,
    vertAlign: state.vertAlign,
    ...(state.size !== undefined ? { size: state.size } : {}),
    ...(state.color !== undefined ? { color: state.color } : {}),
    // 可信粘贴要保留源的关闭状态。null 只删除直接格式，会重新露出目标样式的高亮 / 底纹 / 着重号；
    // 必须写 none / nil 才能盖掉目标的样式层，且新增的直接格式会跟随撤销与 DOCX 回写。
    highlight: state.highlight ?? 'none',
    shading: state.shading ?? { pattern: 'nil', color: 'auto', fill: 'auto' },
    emphasis: state.emphasis ?? 'none',
    ...(ascii || hAnsi || eastAsia
      ? {
          fonts: {
            ...(ascii ? { ascii } : {}),
            ...(hAnsi ? { hAnsi } : {}),
            ...(eastAsia ? { eastAsia } : {}),
          },
        }
      : {}),
  };
}

/** 我们自己、Word、WPS 写的 HTML：格式可以照单全收。 */
export function isTrustedHtml(html: string): boolean {
  return /data-uw-fragment|urn:schemas-microsoft-com:office|mso-[a-z-]+\s*:|\bclass\s*=\s*(?:"(?:[^"]*\s)?Mso|'(?:[^']*\s)?Mso|Mso)/i.test(
    html,
  );
}

const samePatch = (a: RunPropsPatch, b: RunPropsPatch) => JSON.stringify(a) === JSON.stringify(b);

/**
 * Word HTML 的 tab-count 数量不等于 span 内空格数量（微软 WDK 的 Moufiltr.htm 有 2 个 tab、11 个空格）。
 * https://github.com/microsoft/Windows-driver-samples/blob/main/input/moufiltr/Moufiltr.htm
 * 仅转换行内纯空白标记，免得异常 HTML 吞掉正文；单标记最多展开一万个，避免坏数值撑爆粘贴。
 */
function tabCount(el: Element, style: Map<string, string>): number | undefined {
  const value = style.get('mso-tab-count');
  if (el.tagName.toUpperCase() !== 'SPAN' || !value || !/^\d+$/.test(value)) return undefined;
  const count = Number(value);
  if (count < 1 || count > 10_000 || !/^[ \t\r\n\u00a0]*$/.test(el.textContent ?? '')) return undefined;
  if (
    [...el.querySelectorAll('*')].some((child) => {
      const tag = child.tagName.toUpperCase();
      return tag === 'BR' || BLOCKS.has(tag) || SKIP.has(tag);
    })
  )
    return undefined;
  return count;
}

function skippedStyle(style: Map<string, string>): boolean {
  return (
    style.get('display')?.toLowerCase() === 'none' ||
    style.get('mso-hide')?.toLowerCase() === 'all' ||
    style.get('mso-list')?.toLowerCase() === 'ignore'
  );
}

/**
 * 剪贴板 HTML 读成段落。空白按 CSS 的 `white-space: normal` 折叠（段首尾的去掉），
 * `pre` / `pre-wrap` 里原样保留、其中的换行拆段；`&nbsp;` 当普通空格 —— 浏览器与 Word
 * 都拿它表示「连续的第二个空格」，不是真的不断行空格。段里的 `<br>` 是软换行（run 文字里的
 * `\n`，控制器插成 `w:br`），`pre` 里的制表符原样是 `\t`（插成 `w:tab`）。图片与脚本跳过；可信来源的内嵌样式表只用于读取格式，不输出文字。
 */
export function htmlToParagraphs(html: string, parser: DOMParser): PasteParagraph[] {
  const trusted = isTrustedHtml(html);
  const doc = parser.parseFromString(html, 'text/html');
  const root = doc.body;
  const styleOf = clipboardStyleReader(doc, trusted);
  const rootStyle = styleOf(root);
  if (skippedStyle(rootStyle)) return [];
  const paragraphs: PasteParagraph[] = [];
  let current: PasteParagraph | undefined;
  /** 当前段最后一个字是不是空白（折叠用）；段首视为空白，于是段首的空白被吃掉。 */
  let space = true;
  /** 只裁掉段末可折叠的源码空格，不能连同前面 mso-spacerun / pre 的真实空格一起删。 */
  let trailingSpaces = 0;
  let spacesBeforePlaceholder = 0;
  /** 普通 HTML 的末尾 br 只占位；显式软换行必须留在模型里。 */
  let trailingPlaceholder = false;
  const open = (state: Inherited): PasteParagraph => {
    current = { ...(trusted && state.align ? { justification: state.align } : {}), runs: [] };
    space = true;
    trailingSpaces = 0;
    spacesBeforePlaceholder = 0;
    trailingPlaceholder = false;
    return current;
  };
  const close = (keepEmpty: boolean) => {
    if (!current) return;
    // 单个占位 br 仍代表一个空段；多个 br 只能去掉最后一个，不能拆成多个段落。
    const keepPlaceholder = trailingPlaceholder;
    const runs = current.runs;
    const trimEnd = (count: number) => {
      let last = runs.at(-1);
      while (last) {
        const removed = Math.min(count, last.text.length);
        if (removed) last.text = last.text.slice(0, -removed);
        count -= removed;
        if (last.text) return last;
        runs.pop();
        last = runs.at(-1);
      }
      return undefined;
    };
    const last = trimEnd(trailingSpaces);
    if (last && trailingPlaceholder) {
      last.text = last.text.replace(/\n$/, '');
      // 占位前的空格也在段末；可能来自另一个带格式的 run。
      trimEnd(spacesBeforePlaceholder);
    }
    if (runs.length || keepEmpty || keepPlaceholder) paragraphs.push(current);
    current = undefined;
  };
  const append = (text: string, state: Inherited) => {
    if (!text) return;
    const paragraph = current ?? open(state);
    const patch = patchOf(state, trusted);
    const last = paragraph.runs.at(-1);
    if (last && samePatch(last.patch, patch)) last.text += text;
    else paragraph.runs.push({ text, patch });
    // br 后仅有源码空白时仍是段末占位；跨格式的尾空白会在 close 中一起清掉。
    if (state.pre || state.spaces || /[^ ]/.test(text)) trailingPlaceholder = false;
    const tail = text.match(/ +$/)?.[0].length ?? 0;
    trailingSpaces = state.pre || state.spaces ? 0 : tail === text.length ? trailingSpaces + tail : tail;
    space = text.endsWith(' ');
  };
  function text(value: string, state: Inherited): void {
    if (state.pre) {
      const lines = value.replace(/\r\n?/g, '\n').split('\n');
      for (const [i, line] of lines.entries()) {
        if (i > 0) {
          const p = current ?? open(state);
          current = p;
          close(true);
        }
        append(line.replace(/\u00a0/g, ' '), state);
      }
      return;
    }
    // Word 的 HTML 源码按行宽折行，折在两个汉字之间的换行不是空格（CSS Text 3 的分段规则）。
    let collapsed = state.spaces
      ? value.replace(/[\t\n\r\f]+/g, ' ')
      : value.replace(CJK_WRAP, '$1').replace(/[ \t\n\r\f]+/g, ' ');
    if (space && !state.spaces) collapsed = collapsed.replace(/^ /, '');
    append(collapsed.replace(/\u00a0/g, ' '), state);
  }
  function walk(node: Node, state: Inherited): void {
    for (const child of node.childNodes) {
      if (child.nodeType === 3) {
        text(child.nodeValue ?? '', state);
        continue;
      }
      if (child.nodeType !== 1) continue;
      const el = child as Element;
      const tag = el.tagName.toUpperCase();
      if (SKIP.has(tag)) continue;
      const style = styleOf(el);
      // mso-list:Ignore 是 Word 把列表编号写成的正文字（「1.」+ 一串空格），编号不是正文。
      if (skippedStyle(style)) continue;
      // Word 用 <o:p>&nbsp;</o:p> 占住空段落的段落标记，那个空格不是正文。
      if (tag === 'O:P' && !el.textContent?.replace(/[\s\u00a0]/g, '')) continue;
      // 带 page-break-before 的 <br> 是 Word 写的分页符（常在下一段的段首），段首也要留住，记成 U+000C。
      if (
        tag === 'BR' &&
        /^(always|page)$/i.test(style.get('page-break-before') ?? style.get('break-before') ?? '')
      ) {
        append('\f', state);
        space = true;
        continue;
      }
      // 先记录换行，等段末再判断普通 br 是否只是占位；段首 br 后面可能还有正文。
      if (tag === 'BR') {
        const before = trailingSpaces;
        append('\n', inherit(el, state, style));
        spacesBeforePlaceholder = before;
        trailingPlaceholder = style.get('mso-special-character')?.toLowerCase() !== 'line-break';
        space = true;
        continue;
      }
      const next = inherit(el, state, style);
      const tabs = tabCount(el, style);
      if (tabs !== undefined) {
        append('\t'.repeat(tabs), next);
        // 标记后的源码换行不额外插空格；显式 spacerun 仍会原样保留。
        space = true;
        continue;
      }
      if (!BLOCKS.has(tag)) {
        walk(el, next);
        continue;
      }
      if (current?.runs.length) close(false);
      open(next);
      walk(el, next);
      close(EXPLICIT.has(tag));
    }
  }
  walk(
    root,
    inherit(
      root,
      {
        bold: false,
        italic: false,
        underline: undefined,
        strike: false,
        vertAlign: 'baseline',
        pre: false,
        spaces: false,
      },
      rootStyle,
    ),
  );
  close(false);
  return paragraphs;
}
