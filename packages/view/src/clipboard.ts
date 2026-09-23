/**
 * 富文本剪贴板：复制把模型片段写成 HTML，粘贴把 HTML 读回「段落 → 带格式补丁的文字」。
 *
 * 两个方向都**只认内联 `style` 属性的原文**，不走 CSSOM：Word 写的 `mso-fareast-font-family`
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
import type { FragmentRun, Justification, RichFragment, RunPropsPatch } from '@uw/model';

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
  ];
  return style.join(';');
}

/** Word 写分页符的原样：粘进 Word 仍是分页符，浏览器与我们自己都认 `page-break-before`。 */
const PAGE_BREAK_HTML = '<br clear="all" style="mso-special-character:line-break;page-break-before:always">';

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
            `<span style="${escapeAttr(runStyle(r.props))}">${escapeText(r.text).replace(/\n/g, '<br>').replace(/\f/g, PAGE_BREAK_HTML)}</span>`,
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
  pre: boolean;
  align?: Justification | undefined;
}

/** 手写的 `style` 属性解析：分号在引号里不算分隔。属性名一律小写。 */
function parseStyle(text: string | null): Map<string, string> {
  const out = new Map<string, string>();
  if (!text) return out;
  let quote = '';
  let start = 0;
  const flush = (end: number) => {
    const decl = text.slice(start, end);
    const colon = decl.indexOf(':');
    if (colon > 0) out.set(decl.slice(0, colon).trim().toLowerCase(), decl.slice(colon + 1).trim());
  };
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quote) {
      if (c === quote) quote = '';
    } else if (c === '"' || c === "'") quote = c;
    else if (c === ';') {
      flush(i);
      start = i + 1;
    }
  }
  flush(text.length);
  return out;
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

function inherit(el: Element, from: Inherited): Inherited {
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
  const style = parseStyle(el.getAttribute('style'));
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
  const whiteSpace = style.get('white-space');
  if (whiteSpace) next.pre = whiteSpace.startsWith('pre') || whiteSpace === 'break-spaces';
  const textAlign = style.get('text-align');
  if (textAlign && CSS_JUSTIFY[textAlign.toLowerCase()]) next.align = CSS_JUSTIFY[textAlign.toLowerCase()];
  return next;
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
  // Word 只在一款字体上写 font-family 时（纯中文或纯西文的 run），它就是四个桶共同的字体；
  // 写了 mso-* 的桶以 mso-* 为准。一个都没写就不动 —— 沿用光标处的字体，而不是清成样式值。
  const ascii = state.ascii ?? state.family;
  const eastAsia = state.eastAsia ?? state.family;
  return {
    bold: state.bold,
    italic: state.italic,
    underline: state.underline ?? 'none',
    strike: state.strike,
    vertAlign: state.vertAlign,
    ...(state.size !== undefined ? { size: state.size } : {}),
    ...(state.color !== undefined ? { color: state.color } : {}),
    ...(ascii || eastAsia
      ? {
          fonts: {
            ...(ascii ? { ascii, hAnsi: state.hAnsi ?? ascii } : {}),
            ...(eastAsia ? { eastAsia } : {}),
          },
        }
      : {}),
  };
}

/** 我们自己、Word、WPS 写的 HTML：格式可以照单全收。 */
export function isTrustedHtml(html: string): boolean {
  return /data-uw-fragment|urn:schemas-microsoft-com:office|mso-[a-z-]+\s*:|class="?Mso/i.test(html);
}

const samePatch = (a: RunPropsPatch, b: RunPropsPatch) => JSON.stringify(a) === JSON.stringify(b);

/**
 * 剪贴板 HTML 读成段落。空白按 CSS 的 `white-space: normal` 折叠（段首尾的去掉），
 * `pre` / `pre-wrap` 里原样保留、其中的换行拆段；`&nbsp;` 当普通空格 —— 浏览器与 Word
 * 都拿它表示「连续的第二个空格」，不是真的不断行空格。段里的 `<br>` 是软换行（run 文字里的
 * `\n`，控制器插成 `w:br`），`pre` 里的制表符原样是 `\t`（插成 `w:tab`）。图片、脚本、样式表跳过。
 */
export function htmlToParagraphs(html: string, parser: DOMParser): PasteParagraph[] {
  const trusted = isTrustedHtml(html);
  const root = parser.parseFromString(html, 'text/html').body;
  const paragraphs: PasteParagraph[] = [];
  let current: PasteParagraph | undefined;
  /** 当前段最后一个字是不是空白（折叠用）；段首视为空白，于是段首的空白被吃掉。 */
  let space = true;
  /** 最后一段文字来自 `pre`：段尾的空格是原文，不是排版源码里的缩进。 */
  let preserved = false;
  const open = (state: Inherited): PasteParagraph => {
    current = { ...(trusted && state.align ? { justification: state.align } : {}), runs: [] };
    space = true;
    return current;
  };
  const close = (keepEmpty: boolean) => {
    if (!current) return;
    const last = current.runs.at(-1);
    // 块末尾的 <br> 只是结束这一行（浏览器不为它多画一行），不是软换行。
    if (last) last.text = last.text.replace(/\n$/, '');
    if (last && !preserved) last.text = last.text.replace(/ +$/, '');
    if (last && !last.text) current.runs.pop();
    if (current.runs.length || keepEmpty) paragraphs.push(current);
    current = undefined;
  };
  const append = (text: string, state: Inherited) => {
    if (!text) return;
    const paragraph = current ?? open(state);
    const patch = patchOf(state, trusted);
    const last = paragraph.runs.at(-1);
    if (last && samePatch(last.patch, patch)) last.text += text;
    else paragraph.runs.push({ text, patch });
    space = text.endsWith(' ');
    preserved = state.pre;
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
    let collapsed = value.replace(CJK_WRAP, '$1').replace(/[ \t\n\r\f]+/g, ' ');
    if (space) collapsed = collapsed.replace(/^ /, '');
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
      const style = parseStyle(el.getAttribute('style'));
      // mso-list:Ignore 是 Word 把列表编号写成的正文字（「1.」+ 一串空格），编号不是正文。
      if (
        style.get('display') === 'none' ||
        style.get('mso-hide') === 'all' ||
        style.get('mso-list')?.toLowerCase() === 'ignore'
      )
        continue;
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
      // 有字的段里 <br> 是软换行；空段里的 <br> 是浏览器给空行占位的（<div><br></div>），就是一个空段。
      if (tag === 'BR') {
        if (current?.runs.length) {
          append('\n', state);
          space = true;
        } else {
          current ??= open(state);
          close(true);
        }
        continue;
      }
      const next = inherit(el, state);
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
  walk(root, {
    bold: false,
    italic: false,
    underline: undefined,
    strike: false,
    vertAlign: 'baseline',
    pre: false,
  });
  close(false);
  return paragraphs;
}
